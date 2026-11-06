"""Fill the housekeeping team's shared morning briefing workbook."""

from copy import copy
from datetime import datetime
from io import BytesIO
from pathlib import Path
import posixpath
import re
import warnings
from xml.etree import ElementTree
from zipfile import ZIP_DEFLATED, ZipFile

import pandas as pd
import openpyxl
from openpyxl.formula.translate import Translator
from ui.security import sanitize_excel_value

TEMPLATE_PATH = Path(__file__).resolve().parent / "templates" / "HK MB TEMPLATE.xlsx"


def _date_text(value) -> str:
    return pd.Timestamp(value).strftime("%Y-%m-%d")


def _safe_text(value) -> str:
    if value is None or pd.isna(value):
        return ""
    return sanitize_excel_value(value)


def _number(value):
    parsed = pd.to_numeric(value, errors="coerce")
    return None if pd.isna(parsed) else float(parsed)


def _ensure_day_rows(worksheet, days_in_month: int) -> int:
    last_row = 4 + days_in_month
    while worksheet.max_row < last_row:
        source_row = worksheet.max_row
        target_row = source_row + 1
        for column in range(1, worksheet.max_column + 1):
            source = worksheet.cell(source_row, column)
            target = worksheet.cell(target_row, column)
            if source.has_style:
                target._style = copy(source._style)
            if isinstance(source.value, str) and source.value.startswith("="):
                target.value = Translator(source.value, origin=source.coordinate).translate_formula(target.coordinate)
        if worksheet.row_dimensions[source_row].height:
            worksheet.row_dimensions[target_row].height = worksheet.row_dimensions[source_row].height
    return last_row


def _preserve_extension_xml(template_bytes: bytes, workbook_bytes: bytes) -> bytes:
    """Restore unsupported worksheet extension nodes that openpyxl cannot retain."""
    extension_pattern = re.compile(rb"<extLst(?:\s[^>]*)?>.*?</extLst>", re.DOTALL)
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"

    def worksheet_paths(archive: ZipFile) -> dict[str, str]:
        workbook_xml = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        relationships_xml = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relationships = {
            relation.attrib["Id"]: relation.attrib["Target"]
            for relation in relationships_xml.findall(f"{{{package_rel_ns}}}Relationship")
        }
        result = {}
        for sheet in workbook_xml.findall(f"{{{main_ns}}}sheets/{{{main_ns}}}sheet"):
            target = relationships[sheet.attrib[f"{{{rel_ns}}}id"]]
            path = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
            result[sheet.attrib["name"]] = path
        return result

    source = ZipFile(BytesIO(template_bytes))
    generated = ZipFile(BytesIO(workbook_bytes))
    source_paths = worksheet_paths(source)
    generated_paths = {path: title for title, path in worksheet_paths(generated).items()}
    result = BytesIO()
    with source, generated, ZipFile(result, "w", ZIP_DEFLATED) as output:
        for item in generated.infolist():
            contents = generated.read(item.filename)
            if item.filename.startswith("xl/worksheets/") and item.filename.endswith(".xml"):
                sheet_title = generated_paths.get(item.filename)
                source_path = source_paths.get(sheet_title)
                original_xml = source.read(source_path) if source_path else b""
                original_extensions = extension_pattern.findall(original_xml)
                if original_extensions and not extension_pattern.search(contents):
                    contents = contents.replace(
                        b"</worksheet>", b"".join(original_extensions) + b"</worksheet>", 1
                    )
            output.writestr(item, contents)
    return result.getvalue()


def create_morning_briefing_template(
    month: str,
    sr_data: pd.DataFrame | None = None,
    daily_cost_data: pd.DataFrame | None = None,
    manual_data_by_date: dict | None = None,
    guest_laundry_data: pd.DataFrame | None = None,
) -> bytes:
    """Fill the user's workbook without adding or rearranging workbook content."""
    try:
        month_start = datetime.strptime(month, "%Y-%m").date()
    except (TypeError, ValueError) as exc:
        raise ValueError("Month must use YYYY-MM format.") from exc

    from calendar import monthrange

    days_in_month = monthrange(month_start.year, month_start.month)[1]
    yesterday = (datetime.now().date() - pd.Timedelta(days=1))
    briefing_date = yesterday if yesterday.strftime("%Y-%m") == month else month_start.replace(day=days_in_month)
    operation_date = (pd.Timestamp(briefing_date) + pd.Timedelta(days=1)).date()
    briefing_key = _date_text(briefing_date)
    operation_key = _date_text(operation_date)
    manual_data_by_date = manual_data_by_date or {}
    briefing_data = manual_data_by_date.get(briefing_key, {}) or {}
    operation_data = manual_data_by_date.get(operation_key, {}) or {}

    template_bytes = TEMPLATE_PATH.read_bytes()
    with warnings.catch_warnings():
        warnings.filterwarnings(
            "ignore",
            message="Conditional Formatting extension is not supported and will be removed",
            category=UserWarning,
        )
        workbook = openpyxl.load_workbook(BytesIO(template_bytes))
    expected_sheets = [
        "Briefing", "SR Total MTD", "Daily Linen Cost MTD", "Guest Laundry MTD",
        "Daily PA", "PA Matrix", "Room Cost MTD",
    ]
    if workbook.sheetnames != expected_sheets:
        raise ValueError("The shared morning briefing template has an unexpected sheet layout.")

    briefing = workbook["Briefing"]
    briefing["A2"] = f"Daily operational summary | {month} | MTD details are in the supporting sheets"
    briefing["B4"] = briefing_date
    briefing["B4"].number_format = "dd mmm yyyy"
    last_room_cost_row = 4 + days_in_month
    briefing["B6"] = None
    briefing["B7"] = f'=IFERROR(INDEX(\'Room Cost MTD\'!$B$5:$B${last_room_cost_row},MATCH($B$4,\'Room Cost MTD\'!$A$5:$A${last_room_cost_row},0)),"")'
    briefing["B7"].number_format = "0.00%"
    briefing["B8"] = f'=IFERROR(INDEX(\'Room Cost MTD\'!$C$5:$C${last_room_cost_row},MATCH($B$4,\'Room Cost MTD\'!$A$5:$A${last_room_cost_row},0)),"")'
    briefing["B8"].number_format = "#,##0"
    for row, column in ((14, "D"), (15, "E"), (16, "F"), (17, "G")):
        briefing.cell(row, 2, f'=IFERROR(INDEX(\'Room Cost MTD\'!${column}$5:${column}${last_room_cost_row},MATCH($B$4,\'Room Cost MTD\'!$A$5:$A${last_room_cost_row},0))/INDEX(\'Room Cost MTD\'!$C$5:$C${last_room_cost_row},MATCH($B$4,\'Room Cost MTD\'!$A$5:$A${last_room_cost_row},0)),"")')
    for column, source_column in ((1, "B"), (2, "C"), (3, "D")):
        briefing.cell(23, column, f'=IFERROR(INDEX(\'Guest Laundry MTD\'!${source_column}$5:${source_column}${last_room_cost_row},MATCH($B$4,\'Guest Laundry MTD\'!$A$5:$A${last_room_cost_row},0)),"")')
    briefing["A22"] = "Rooms"
    briefing["A23"] = ""
    briefing["B10"] = _safe_text(operation_data.get("Staff_Incharge"))
    briefing["A47"] = _safe_text(operation_data.get("Remarks"))

    sr_sheet = workbook["SR Total MTD"]
    sr_sheet["A2"] = f"Morning Briefing | {month} | Masukkan satu baris per voucher atau item SR; total MTD dihitung otomatis."
    for row in range(5, sr_sheet.max_row + 1):
        for column in range(1, 5):
            sr_sheet.cell(row, column).value = None
    if sr_data is not None and not sr_data.empty and {"Date", "Total"}.issubset(sr_data.columns):
        sr_dates = pd.to_datetime(sr_data["Date"], errors="coerce", format="mixed")
        month_sr = sr_data.loc[sr_dates.dt.strftime("%Y-%m") == month]
        for row, (_, entry) in zip(range(5, sr_sheet.max_row + 1), month_sr.iterrows()):
            date_value = pd.to_datetime(entry["Date"], errors="coerce", format="mixed")
            if pd.notna(date_value):
                sr_sheet.cell(row, 1, date_value.to_pydatetime())
            sr_sheet.cell(row, 2, _safe_text(entry.get("SR_Number")))
            sr_sheet.cell(row, 3, _safe_text(entry.get("Category")))
            sr_sheet.cell(row, 4, _number(entry.get("Total")))

    saved_rows = {}
    if daily_cost_data is not None and not daily_cost_data.empty and "Date" in daily_cost_data.columns:
        saved_dates = pd.to_datetime(daily_cost_data["Date"], errors="coerce", format="mixed")
        for _, entry in daily_cost_data.loc[saved_dates.dt.strftime("%Y-%m") == month].iterrows():
            parsed_date = pd.to_datetime(entry["Date"], errors="coerce", format="mixed")
            if pd.notna(parsed_date):
                saved_rows[parsed_date.date().isoformat()] = entry
    room_cost = workbook["Room Cost MTD"]
    room_cost["A2"] = f"Morning Briefing | {month} | Isi occupancy, room nights, air, listrik, dan linen; SR dan total dihitung otomatis."
    linen_sheet = workbook["Daily Linen Cost MTD"]
    linen_sheet["A2"] = f"Morning Briefing | {month} | Masukkan biaya linen harian; total MTD dihitung otomatis."
    laundry_sheet = workbook["Guest Laundry MTD"]
    laundry_sheet["A2"] = f"Morning Briefing | {month} | Masukkan biaya Drop N Go harian; biaya per room dan total MTD dihitung otomatis."
    laundry_sheet["B4"] = "Rooms"
    imported_laundry_rows = {}
    if guest_laundry_data is not None and not guest_laundry_data.empty and "Date" in guest_laundry_data.columns:
        laundry_dates = pd.to_datetime(guest_laundry_data["Date"], errors="coerce", format="mixed")
        for _, entry in guest_laundry_data.loc[laundry_dates.dt.strftime("%Y-%m") == month].iterrows():
            parsed_date = pd.to_datetime(entry["Date"], errors="coerce", format="mixed")
            if pd.notna(parsed_date):
                imported_laundry_rows[parsed_date.strftime("%Y-%m-%d")] = entry

    _ensure_day_rows(room_cost, days_in_month)
    _ensure_day_rows(linen_sheet, days_in_month)
    _ensure_day_rows(laundry_sheet, days_in_month)
    last_room_cost_row = 4 + days_in_month
    for row in range(5, room_cost.max_row + 1):
        for column in (2, 3, 5, 6, 7):
            room_cost.cell(row, column).value = None
        if row > last_room_cost_row:
            room_cost.cell(row, 1).value = None
    for row in range(5, linen_sheet.max_row + 1):
        linen_sheet.cell(row, 2).value = None
        if row > last_room_cost_row:
            linen_sheet.cell(row, 1).value = None
    for row in range(5, laundry_sheet.max_row + 1):
        for column in (2, 3):
            laundry_sheet.cell(row, column).value = None
        if row > last_room_cost_row:
            laundry_sheet.cell(row, 1).value = None

    for day in range(days_in_month):
        current_date = month_start.replace(day=day + 1)
        date_key = current_date.isoformat()
        row = day + 5
        linen_row = day + 5
        laundry_row = day + 5
        room_cost.cell(row, 1, current_date)
        room_cost.cell(row, 4, f'=IFERROR(C{row}*SUM(\'SR Total MTD\'!$D$5:$D${sr_sheet.max_row})/SUM($C$5:$C${last_room_cost_row}),0)')
        room_cost.cell(row, 8, f'=IF(COUNTA(C{row},E{row}:G{row})=0,"",SUM(D{row}:G{row}))')
        room_cost.cell(row, 9, f'=IFERROR(H{row}/C{row},"")')
        linen_sheet.cell(linen_row, 1, current_date)
        linen_sheet.cell(linen_row, 3, f'=SUM($B$5:B{linen_row})')
        laundry_sheet.cell(laundry_row, 1, current_date)
        laundry_sheet.cell(laundry_row, 4, f'=IFERROR(C{laundry_row}/B{laundry_row},"")')
        laundry_sheet.cell(laundry_row, 5, f'=SUM($C$5:C{laundry_row})')
        saved = saved_rows.get(current_date.isoformat())
        if saved is not None:
            for key, column in (("Occupancy", 2), ("Room_Nights", 3), ("Water", 5), ("Electricity", 6), ("Linen_Laundry", 7)):
                value = _number(saved.get(key))
                if value is not None:
                    room_cost.cell(row, column, value / 100 if key == "Occupancy" else value)
                    if key == "Linen_Laundry":
                        linen_sheet.cell(linen_row, 2, value)

        imported_laundry = imported_laundry_rows.get(date_key)
        if imported_laundry is not None:
            laundry_sheet.cell(laundry_row, 2, _number(imported_laundry.get("Room_Nights")))
            laundry_sheet.cell(laundry_row, 3, _number(imported_laundry.get("Guest_Laundry_Cost")))
            if date_key == briefing_key:
                room_numbers = [
                    room.strip()
                    for room in _safe_text(imported_laundry.get("Room_Numbers")).split(",")
                    if room.strip()
                ]
                briefing["A23"] = "\n".join(
                    ", ".join(room_numbers[index:index + 3])
                    for index in range(0, len(room_numbers), 3)
                )
                alignment = copy(briefing["A23"].alignment)
                alignment.wrap_text = True
                briefing["A23"].alignment = alignment
                if room_numbers:
                    briefing.row_dimensions[23].height = max(briefing.row_dimensions[23].height or 0, 15 * ((len(room_numbers) + 2) // 3))
        else:
            laundry_data = manual_data_by_date.get(date_key, {}) or {}
            laundry_entries = laundry_data.get("Entries", pd.DataFrame())
            if not laundry_entries.empty and "Entry_Type" in laundry_entries:
                matching_laundry = laundry_entries.loc[laundry_entries["Entry_Type"] == "guest_laundry"]
                if not matching_laundry.empty:
                    laundry_entry = matching_laundry.iloc[-1]
                    laundry_sheet.cell(laundry_row, 2, _number(laundry_entry.get("Qty")))
                    laundry_sheet.cell(laundry_row, 3, _number(laundry_entry.get("Total")))

    ra_records = []
    for report_date, report_data in ((briefing_date, briefing_data), (operation_date, operation_data)):
        entries = report_data.get("Entries", pd.DataFrame())
        if not entries.empty and "Entry_Type" in entries:
            ra_records.extend((report_date, entry) for _, entry in entries.loc[entries["Entry_Type"] == "ra"].iterrows())

    # Briefing 'PROJECT RA | HARI INI': A=Staff_Name, B=Description, C=Section,
    # D='{Room_Credit} Room' + Notes. Rows 28-33 hold the operation-day RA entries.
    for row in range(28, 34):
        for column in range(1, 5):
            briefing.cell(row, column).value = None
    operation_ra = [entry for report_date, entry in ra_records if report_date == operation_date]
    for row, entry in zip(range(28, 34), operation_ra):
        briefing.cell(row, 1, _safe_text(entry.get("Staff_Name")))
        briefing.cell(row, 2, _safe_text(entry.get("Description")))
        briefing.cell(row, 3, _safe_text(entry.get("Section")))
        credit = _number(entry.get("Room_Credit"))
        credit_text = f"{int(credit)} Room" if credit is not None and credit == int(credit) else (f"{credit} Room" if credit is not None else "")
        notes = _safe_text(entry.get("Notes"))
        briefing.cell(row, 4, " — ".join(filter(None, [credit_text, notes])))
    unique_staff = {_safe_text(entry.get("Staff_Name")) for _, entry in ra_records if _safe_text(entry.get("Staff_Name"))}
    briefing["B9"] = len(unique_staff) if unique_staff else None

    project_entries = operation_data.get("Entries", pd.DataFrame())
    if not project_entries.empty and "Entry_Type" in project_entries:
        project_entries = project_entries.loc[project_entries["Entry_Type"] == "project_pa"]
    # PA rows merge vertically in two-row blocks; write merge anchors (A,B,C)
    # plus column D, which is never merged.
    for row in range(37, 43):
        briefing.cell(row, 4).value = None
    for row in (37, 39, 41):
        for column in (1, 2, 3):
            briefing.cell(row, column).value = None
    for row, (_, entry) in zip((37, 39, 41), project_entries.iterrows() if not project_entries.empty else []):
        briefing.cell(row, 1, _safe_text(entry.get("Area")))
        briefing.cell(row, 2, _safe_text(entry.get("Description")))
        briefing.cell(row, 3, _safe_text(entry.get("Unit")))
        briefing.cell(row, 4, _safe_text(entry.get("Notes")))

    checklist = workbook["PA Matrix"]
    checklist["D2"] = datetime(month_start.year, month_start.month, 1)
    checklist["D2"].number_format = "mmmm yyyy"
    for row in range(4, checklist.max_row + 1):
        for column in range(4, min(35, checklist.max_column + 1)):
            checklist.cell(row, column).value = None

    output = BytesIO()
    workbook.save(output)
    return _preserve_extension_xml(template_bytes, output.getvalue())
