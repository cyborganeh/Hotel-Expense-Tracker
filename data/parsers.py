"""Excel/PDF parsing + normalization for daily room-cost inputs.

Split out of the former monolith room_cost.py. Owns the column constants,
month validation, the shared normalization pipeline, and the four upload
parsers (daily cost, night audit, energy, linen).
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
