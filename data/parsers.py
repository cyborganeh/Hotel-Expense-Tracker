"""Excel/PDF parsing + normalization for daily room-cost inputs.

Split out of the former monolith room_cost.py. Owns the column constants,
month validation, the shared normalization pipeline, and workbook parsers for
daily room costs, audit/energy inputs, and laundry reports.
"""
import re
from datetime import datetime
from typing import Optional

import openpyxl
import pandas as pd


DAILY_COLUMNS = ["Date", "Occupancy", "Room_Nights", "Water", "Electricity", "Linen_Laundry"]
REQUIRED_DAILY_COLUMNS = ["Date", "Room_Nights", "Water", "Electricity", "Linen_Laundry"]
EXPENSE_COLUMNS = ["Water", "Electricity", "Linen_Laundry"]


def _month_key(month: str) -> str:
    value = str(month).strip()
    if not re.fullmatch(r"\d{4}-\d{2}", value):
        raise ValueError("Month must use YYYY-MM format.")
    year, month_number = (int(part) for part in value.split("-"))
    if not 1 <= month_number <= 12:
        raise ValueError("Month must use YYYY-MM format.")
    return value


def _normalise_daily_frame(
    frame: pd.DataFrame,
    month: str,
    *,
    allow_missing_expenses: bool = False,
    allow_missing_room_nights: bool = False,
) -> pd.DataFrame:
    month = _month_key(month)
    if frame.empty:
        return pd.DataFrame(columns=DAILY_COLUMNS)

    aliases = {
        "date": "Date",
        "occupancy": "Occupancy",
        "occupancypercent": "Occupancy",
        "occ": "Occupancy",
        "occpercent": "Occupancy",
        "roomnights": "Room_Nights",
        "roomsold": "Room_Nights",
        "expocc": "Room_Nights",
        "water": "Water",
        "watercost": "Water",
        "electricity": "Electricity",
        "electricitycost": "Electricity",
        "linen": "Linen_Laundry",
        "laundry": "Linen_Laundry",
        "linenlaundry": "Linen_Laundry",
        "linenlaundrycost": "Linen_Laundry",
    }
    renamed = {}
    for column in frame.columns:
        key = re.sub(r"[^a-z0-9]", "", str(column).strip().lower())
        if key in aliases:
            canonical = aliases[key]
            if canonical in renamed.values():
                raise ValueError(f"Excel contains more than one column for {canonical.replace('_', ' ')}.")
            renamed[column] = canonical
    normalized = frame.rename(columns=renamed)

    if "Occupancy" not in normalized.columns:
        normalized["Occupancy"] = pd.NA
    missing_columns = [column for column in REQUIRED_DAILY_COLUMNS if column not in normalized.columns]
    if missing_columns:
        display = ", ".join(column.replace("_", " ") for column in missing_columns)
        raise ValueError(f"Excel is missing required columns: {display}.")

    data = normalized[DAILY_COLUMNS].copy()
    parsed_dates = pd.to_datetime(data["Date"], errors="coerce", format="mixed")
    invalid_dates = parsed_dates.isna()
    if invalid_dates.any():
        rows = ", ".join(str(index + 2) for index in data.index[invalid_dates])
        raise ValueError(f"Invalid or missing Date value on Excel row(s): {rows}.")
    data["Date"] = parsed_dates.dt.strftime("%Y-%m-%d")

    wrong_month = data["Date"].str[:7] != month
    if wrong_month.any():
        dates = ", ".join(data.loc[wrong_month, "Date"].tolist())
        raise ValueError(f"Date(s) outside selected month {month}: {dates}.")

    duplicate_dates = data.loc[data["Date"].duplicated(keep=False), "Date"].unique().tolist()
    if duplicate_dates:
        raise ValueError(f"Excel has duplicate date(s): {', '.join(duplicate_dates)}.")

    for column in ["Room_Nights", *EXPENSE_COLUMNS]:
        if column == "Room_Nights" and allow_missing_room_nights:
            values = pd.to_numeric(data[column], errors="coerce")
            if (values.dropna() < 0).any() or (values.dropna() % 1 != 0).any():
                raise ValueError("Room Nights must be a positive whole number when provided.")
            if (values.dropna() == 0).any():
                raise ValueError("Room Nights must be positive; zero is not a valid value.")
            data[column] = values.astype("Int64")
            continue
        if column in EXPENSE_COLUMNS and allow_missing_expenses:
            values = pd.to_numeric(data[column], errors="coerce")
            if (values.dropna() < 0).any():
                raise ValueError(f"{column.replace('_', ' ')} cannot be negative.")
            data[column] = values.astype(float)
            continue
        values = pd.to_numeric(data[column], errors="coerce")
        if values.isna().any():
            rows = ", ".join(str(index + 2) for index in data.index[values.isna()])
            raise ValueError(f"Missing or non-numeric {column.replace('_', ' ')} on Excel row(s): {rows}.")
        if (values < 0).any():
            raise ValueError(f"{column.replace('_', ' ')} cannot be negative.")
        data[column] = values.astype(float)

    occupancy = pd.to_numeric(data["Occupancy"], errors="coerce")
    if (occupancy.dropna() < 0).any() or (occupancy.dropna() > 100).any():
        raise ValueError("Occupancy must be between 0 and 100 percent.")
    data["Occupancy"] = occupancy.astype(float)

    invalid_room_nights = (data["Room_Nights"] <= 0) | (data["Room_Nights"] % 1 != 0)
    if invalid_room_nights.any():
        raise ValueError("Room Nights must be a positive whole number for every date.")

    data["Room_Nights"] = data["Room_Nights"].astype("Int64" if allow_missing_room_nights else int)
    return data.sort_values("Date").reset_index(drop=True)


def parse_daily_room_cost_excel(source, month: str) -> pd.DataFrame:
    """Read and validate daily room-cost rows from an uploaded Excel workbook."""
    if hasattr(source, "seek"):
        source.seek(0)
    frame = pd.read_excel(source, engine="openpyxl")
    if hasattr(source, "seek"):
        source.seek(0)
    return _normalise_daily_frame(frame, month)


def parse_night_audit_report(source, month: Optional[str] = None) -> pd.DataFrame:
    """Extract date, occupancy, and expected rooms occupied from a Night Audit report."""
    workbook = openpyxl.load_workbook(source, data_only=True, read_only=True)
    try:
        for worksheet in workbook.worksheets:
            header_row = None
            columns = {}
            for row_idx, values in enumerate(worksheet.iter_rows(values_only=True), 1):
                labels = {
                    re.sub(r"[^a-z0-9]", "", str(value).strip().lower()): col_idx
                    for col_idx, value in enumerate(values)
                    if value is not None
                }
                date_col = labels.get("date")
                room_sold_col = labels.get("expocc") or labels.get("roomsold")
                occupancy_col = labels.get("occ") or labels.get("occpercent")
                if date_col is not None and room_sold_col is not None and occupancy_col is not None:
                    header_row = row_idx
                    columns = {
                        "Date": date_col,
                        "Room_Nights": room_sold_col,
                        "Occupancy": occupancy_col,
                    }
                    break
            if header_row is None:
                continue

            records = []
            for values in worksheet.iter_rows(min_row=header_row + 1, values_only=True):
                date_value = values[columns["Date"]] if columns["Date"] < len(values) else None
                if isinstance(date_value, datetime):
                    parsed_date = pd.Timestamp(date_value).strftime("%Y-%m-%d")
                else:
                    parsed = pd.to_datetime(date_value, errors="coerce", format="mixed")
                    if pd.isna(parsed):
                        continue
                    parsed_date = parsed.strftime("%Y-%m-%d")

                room_sold = values[columns["Room_Nights"]] if columns["Room_Nights"] < len(values) else None
                occupancy = values[columns["Occupancy"]] if columns["Occupancy"] < len(values) else None
                if room_sold is None or occupancy is None:
                    continue
                records.append({
                    "Date": parsed_date,
                    "Occupancy": occupancy,
                    "Room_Nights": room_sold,
                    "Water": pd.NA,
                    "Electricity": pd.NA,
                    "Linen_Laundry": pd.NA,
                })
            if records:
                detected_month = month or records[0]["Date"][:7]
                return _normalise_daily_frame(
                    pd.DataFrame(records), detected_month, allow_missing_expenses=True
                )
        raise ValueError("Could not find DATE, EXP OCC/ROOM SOLD, and OCC (%) columns in the workbook.")
    finally:
        workbook.close()


def parse_energy_cost_workbook(source, month: str) -> pd.DataFrame:
    """Extract daily electricity and water amounts from an Energy Cost workbook."""
    month = _month_key(month)
    workbook = openpyxl.load_workbook(source, data_only=True, read_only=True)
    try:
        for worksheet in workbook.worksheets:
            header_row = None
            columns = {}
            for row_idx, values in enumerate(worksheet.iter_rows(values_only=True), 1):
                labels = {
                    re.sub(r"[^a-z0-9]", "", str(value).strip().lower()): col_idx
                    for col_idx, value in enumerate(values)
                    if value is not None
                }
                date_col = labels.get("tanggal")
                if date_col is None:
                    date_col = labels.get("date")
                electricity_col = next(
                    (idx for key, idx in labels.items() if "listrik" in key or "electricity" in key),
                    None,
                )
                water_col = next(
                    (idx for key, idx in labels.items() if "air" in key or "water" in key),
                    None,
                )
                if date_col is not None and electricity_col is not None and water_col is not None:
                    header_row = row_idx
                    columns = {"Date": date_col, "Electricity": electricity_col, "Water": water_col}
                    break
            if header_row is None:
                continue

            year, month_number = (int(part) for part in month.split("-"))
            records = []
            for values in worksheet.iter_rows(min_row=header_row + 1, values_only=True):
                day_value = values[columns["Date"]] if columns["Date"] < len(values) else None
                if day_value is None:
                    continue
                if isinstance(day_value, (datetime, pd.Timestamp)):
                    parsed_date = pd.Timestamp(day_value)
                else:
                    numeric_day = pd.to_numeric(day_value, errors="coerce")
                    if pd.isna(numeric_day) or float(numeric_day) % 1:
                        continue
                    try:
                        parsed_date = pd.Timestamp(year=year, month=month_number, day=int(numeric_day))
                    except ValueError:
                        continue
                if parsed_date.strftime("%Y-%m") != month:
                    continue

                electricity = values[columns["Electricity"]] if columns["Electricity"] < len(values) else None
                water = values[columns["Water"]] if columns["Water"] < len(values) else None
                if electricity is None and water is None:
                    continue
                records.append({
                    "Date": parsed_date.strftime("%Y-%m-%d"),
                    "Occupancy": pd.NA,
                    "Room_Nights": pd.NA,
                    "Water": water,
                    "Electricity": electricity,
                    "Linen_Laundry": pd.NA,
                })
            if records:
                return _normalise_daily_frame(
                    pd.DataFrame(records),
                    month,
                    allow_missing_expenses=True,
                    allow_missing_room_nights=True,
                )
        raise ValueError("Could not find Tanggal/Date, Penggunaan Listrik/Electricity, and Penggunaan Air/Water columns.")
    finally:
        workbook.close()


def parse_daily_linen_cost_workbook(source, month: str) -> pd.DataFrame:
    """Extract daily linen cost, excluding zero-filled future rows beyond its data count."""
    month = _month_key(month)
    workbook = openpyxl.load_workbook(source, data_only=True, read_only=True)
    try:
        for worksheet in workbook.worksheets:
            header_row = None
            date_col = None
            linen_total_col = None
            for row_idx, values in enumerate(worksheet.iter_rows(values_only=True), 1):
                labels = {
                    re.sub(r"[^a-z0-9]", "", str(value).strip().lower()): col_idx
                    for col_idx, value in enumerate(values)
                    if value is not None
                }
                possible_date = labels.get("tanggal")
                if possible_date is None:
                    possible_date = labels.get("date")
                total_col = next(
                    (idx for key, idx in labels.items() if "totalbiayaharian" in key),
                    None,
                )
                if possible_date is not None and total_col is not None:
                    header_row = row_idx
                    date_col = possible_date
                    linen_total_col = total_col
                    break
            if header_row is None:
                continue

            populated_days = None
            for values in worksheet.iter_rows(values_only=True):
                labels = [re.sub(r"[^a-z0-9]", "", str(value).strip().lower()) if value is not None else "" for value in values]
                if "jumlahhariyangadadata" in labels:
                    label_col = labels.index("jumlahhariyangadadata")
                    count_values = [
                        pd.to_numeric(value, errors="coerce")
                        for value in values[label_col + 1:]
                        if value is not None
                    ]
                    count = next((value for value in count_values if pd.notna(value)), None)
                    if count is not None:
                        populated_days = int(count)
                    break

            year, month_number = (int(part) for part in month.split("-"))
            records = []
            for values in worksheet.iter_rows(min_row=header_row + 1, values_only=True):
                day_value = values[date_col] if date_col < len(values) else None
                if day_value is None or str(day_value).strip().upper() == "TOTAL":
                    continue
                if isinstance(day_value, (datetime, pd.Timestamp)):
                    parsed_date = pd.Timestamp(day_value)
                else:
                    numeric_day = pd.to_numeric(day_value, errors="coerce")
                    if pd.isna(numeric_day) or float(numeric_day) % 1:
                        continue
                    try:
                        parsed_date = pd.Timestamp(year=year, month=month_number, day=int(numeric_day))
                    except ValueError:
                        continue
                if parsed_date.strftime("%Y-%m") != month:
                    continue
                if populated_days is not None and parsed_date.day > populated_days:
                    continue

                linen_total = values[linen_total_col] if linen_total_col < len(values) else None
                if linen_total is None:
                    continue
                records.append({
                    "Date": parsed_date.strftime("%Y-%m-%d"),
                    "Occupancy": pd.NA,
                    "Room_Nights": pd.NA,
                    "Water": pd.NA,
                    "Electricity": pd.NA,
                    "Linen_Laundry": linen_total,
                })
            if records:
                return _normalise_daily_frame(
                    pd.DataFrame(records),
                    month,
                    allow_missing_expenses=True,
                    allow_missing_room_nights=True,
                )
        raise ValueError("Could not find Tanggal/Date and TOTAL BIAYA HARIAN columns in the workbook.")
    finally:
        workbook.close()


def _report_day(value, month: str) -> str | None:
    year, month_number = (int(part) for part in month.split("-"))
    if isinstance(value, (datetime, pd.Timestamp)):
        day = value.day
    else:
        numeric_day = pd.to_numeric(value, errors="coerce")
        if pd.isna(numeric_day) or float(numeric_day) % 1:
            parsed = pd.to_datetime(value, errors="coerce", format="mixed")
            if pd.isna(parsed):
                return None
            day = parsed.day
        else:
            day = int(numeric_day)
    try:
        return pd.Timestamp(year=year, month=month_number, day=int(day)).strftime("%Y-%m-%d")
    except ValueError:
        return None


def parse_bonvivo_linen_workbook(source, month: str) -> pd.DataFrame:
    """Extract daily Linen HK charges, including regular, spotting/OOO, and prewash."""
    month = _month_key(month)
    if hasattr(source, "seek"):
        source.seek(0)
    workbook = openpyxl.load_workbook(source, data_only=True, read_only=False)
    try:
        worksheet = next(
            (sheet for sheet in workbook.worksheets
             if re.sub(r"[^a-z0-9]", "", sheet.title.lower()) == "linenhk"),
            None,
        )
        if worksheet is None:
            raise ValueError("Could not find the LINEN HK sheet in the Bonvivo workbook.")

        daily_totals: dict[str, float] = {}
        header_rows = []
        for row_idx in range(1, worksheet.max_row + 1):
            label = re.sub(r"[^a-z0-9]", "", str(worksheet.cell(row_idx, 2).value or "").lower())
            if label == "tanggal":
                date_columns = {
                    col: _report_day(worksheet.cell(row_idx, col).value, month)
                    for col in range(3, 33)
                }
                if any(date_columns.values()):
                    header_rows.append((row_idx, date_columns))

        if not header_rows:
            raise ValueError("Could not find dated Linen HK detail tables in the Bonvivo workbook.")

        for header_index, (header_row, date_columns) in enumerate(header_rows):
            end_row = header_rows[header_index + 1][0] - 1 if header_index + 1 < len(header_rows) else worksheet.max_row
            for row_idx in range(header_row + 1, end_row + 1):
                item_label = str(worksheet.cell(row_idx, 2).value or "").strip().lower()
                row_label = str(worksheet.cell(row_idx, 1).value or "").strip().lower()
                if "total" in row_label or "total" in item_label:
                    break
                price = pd.to_numeric(worksheet.cell(row_idx, 34).value, errors="coerce")
                if pd.isna(price) or price < 0:
                    continue
                for col, date_value in date_columns.items():
                    if date_value is None:
                        continue
                    quantity = pd.to_numeric(worksheet.cell(row_idx, col).value, errors="coerce")
                    if pd.isna(quantity) or quantity == 0:
                        continue
                    if quantity < 0:
                        raise ValueError(f"Negative Linen HK quantity on Excel row {row_idx}.")
                    daily_totals[date_value] = daily_totals.get(date_value, 0.0) + float(quantity) * float(price)

        records = [
            {
                "Date": date_value,
                "Occupancy": pd.NA,
                "Room_Nights": pd.NA,
                "Water": pd.NA,
                "Electricity": pd.NA,
                "Linen_Laundry": amount,
            }
            for date_value, amount in sorted(daily_totals.items())
            if amount > 0
        ]
        if not records:
            raise ValueError("No Linen HK charges were found for the selected month.")
        return _normalise_daily_frame(
            pd.DataFrame(records), month,
            allow_missing_expenses=True,
            allow_missing_room_nights=True,
        )
    finally:
        workbook.close()
        if hasattr(source, "seek"):
            source.seek(0)


def parse_dropngo_guest_laundry_workbook(source, month: str) -> pd.DataFrame:
    """Extract daily guest-laundry cost and distinct serviced rooms from Drop N Go."""
    month = _month_key(month)
    if hasattr(source, "seek"):
        source.seek(0)
    workbook = openpyxl.load_workbook(source, data_only=True, read_only=False)
    try:
        worksheet = next(
            (sheet for sheet in workbook.worksheets
             if re.sub(r"[^a-z0-9]", "", sheet.title.lower()) == "guestlaundry"),
            None,
        )
        if worksheet is None:
            raise ValueError("Could not find the GUEST LAUNDRY sheet in the Drop N Go workbook.")

        header_row = None
        for row_idx in range(1, worksheet.max_row + 1):
            labels = {
                re.sub(r"[^a-z0-9]", "", str(worksheet.cell(row_idx, col).value or "").lower())
                for col in range(1, 5)
            }
            if {"item", "price", "date"}.issubset(labels):
                header_row = row_idx
                break
        if header_row is None:
            raise ValueError("Could not find Item, price, and Date headers in GUEST LAUNDRY.")

        day_by_column = {}
        merged_headers = {
            col: worksheet.cell(rng.min_row, rng.min_col).value
            for rng in worksheet.merged_cells.ranges
            if rng.min_row <= header_row + 1 <= rng.max_row
            and rng.min_col >= 4
            for col in range(rng.min_col, rng.max_col + 1)
        }
        for col in range(4, 57):
            day_value = merged_headers.get(col, worksheet.cell(header_row + 1, col).value)
            day_by_column[col] = _report_day(day_value, month)

        totals: dict[str, float] = {}
        rooms: dict[str, set[str]] = {}
        for col, date_value in day_by_column.items():
            room_value = worksheet.cell(header_row + 2, col).value
            if date_value is None or room_value is None:
                continue
            room_key = str(room_value).strip()
            if not room_key:
                continue
            for row_idx in range(header_row + 3, worksheet.max_row + 1):
                price = pd.to_numeric(worksheet.cell(row_idx, 3).value, errors="coerce")
                quantity = pd.to_numeric(worksheet.cell(row_idx, col).value, errors="coerce")
                if pd.isna(price) or pd.isna(quantity) or quantity == 0:
                    continue
                if price < 0 or quantity < 0:
                    raise ValueError(f"Negative Drop N Go price or quantity on Excel row {row_idx}.")
                totals[date_value] = totals.get(date_value, 0.0) + float(price) * float(quantity)
                rooms.setdefault(date_value, set()).add(room_key)

        records = [
            {
                "Date": date_value,
                "Room_Nights": len(rooms[date_value]),
                "Room_Numbers": ", ".join(sorted(rooms[date_value], key=lambda room: (not room.isdigit(), int(room) if room.isdigit() else room))),
                "Guest_Laundry_Cost": amount,
            }
            for date_value, amount in sorted(totals.items())
            if amount > 0 and rooms.get(date_value)
        ]
        return pd.DataFrame(records, columns=["Date", "Room_Nights", "Room_Numbers", "Guest_Laundry_Cost"])
    finally:
        workbook.close()
        if hasattr(source, "seek"):
            source.seek(0)


def parse_dropngo_linen_workbook(source, month: str) -> pd.DataFrame:
    """Extract daily Linen HK charges from Drop N Go's item-by-day worksheet."""
    month = _month_key(month)
    if hasattr(source, "seek"):
        source.seek(0)
    workbook = openpyxl.load_workbook(source, data_only=True, read_only=False)
    try:
        worksheet = next(
            (sheet for sheet in workbook.worksheets
             if re.sub(r"[^a-z0-9]", "", sheet.title.lower()) == "linenhk"),
            None,
        )
        if worksheet is None:
            raise ValueError("Could not find the LINEN HK sheet in the Drop N Go workbook.")

        header_row = None
        date_columns = {}
        for row_idx in range(1, worksheet.max_row + 1):
            item_label = re.sub(r"[^a-z0-9]", "", str(worksheet.cell(row_idx, 2).value or "").lower())
            price_label = re.sub(r"[^a-z0-9]", "", str(worksheet.cell(row_idx, 3).value or "").lower())
            if item_label == "item" and price_label == "price":
                candidate = {
                    col: _report_day(worksheet.cell(row_idx, col).value, month)
                    for col in range(4, min(35, worksheet.max_column + 1))
                }
                if any(candidate.values()):
                    header_row = row_idx
                    date_columns = candidate
                    break
        if header_row is None:
            raise ValueError("Could not find dated item/price columns in Drop N Go LINEN HK.")

        totals: dict[str, float] = {}
        for row_idx in range(header_row + 1, worksheet.max_row + 1):
            item_label = str(worksheet.cell(row_idx, 2).value or "").strip().lower()
            if "total" in item_label:
                break
            price = pd.to_numeric(worksheet.cell(row_idx, 3).value, errors="coerce")
            if pd.isna(price) or price < 0:
                continue
            for col, date_value in date_columns.items():
                if date_value is None:
                    continue
                quantity = pd.to_numeric(worksheet.cell(row_idx, col).value, errors="coerce")
                if pd.isna(quantity) or quantity == 0:
                    continue
                if quantity < 0:
                    raise ValueError(f"Negative Drop N Go Linen HK quantity on Excel row {row_idx}.")
                totals[date_value] = totals.get(date_value, 0.0) + float(quantity) * float(price)

        records = [
            {
                "Date": date_value,
                "Occupancy": pd.NA,
                "Room_Nights": pd.NA,
                "Water": pd.NA,
                "Electricity": pd.NA,
                "Linen_Laundry": amount,
            }
            for date_value, amount in sorted(totals.items())
            if amount > 0
        ]
        if not records:
            raise ValueError("No Drop N Go Linen HK charges were found for the selected month.")
        return _normalise_daily_frame(
            pd.DataFrame(records), month,
            allow_missing_expenses=True,
            allow_missing_room_nights=True,
        )
    finally:
        workbook.close()
        if hasattr(source, "seek"):
            source.seek(0)
