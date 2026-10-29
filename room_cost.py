"""Daily room-cost input, MTD calculations, and local persistence."""
import json
import os
import re
import tempfile
from datetime import datetime
from pathlib import Path
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


def get_room_cost_storage_dir(storage_dir: Optional[str] = None) -> str:
    """Return the app-owned directory for persisted room-cost inputs."""
    path = storage_dir or os.environ.get("HOTEL_ROOM_COST_DIR")
    if not path:
        path = os.path.join(Path.home(), ".hotel-expense-tracker", "room_cost")
    return os.path.abspath(os.path.expanduser(path))


def load_daily_room_costs(month: str, storage_dir: Optional[str] = None) -> pd.DataFrame:
    """Load persisted daily costs for a calendar month, if present."""
    month = _month_key(month)
    path = os.path.join(get_room_cost_storage_dir(storage_dir), f"{month}.json")
    if not os.path.exists(path):
        return pd.DataFrame(columns=DAILY_COLUMNS)
    try:
        with open(path, "r", encoding="utf-8") as file_obj:
            rows = json.load(file_obj)
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"Could not read saved room-cost data for {month}.") from exc
    return _normalise_daily_frame(
        pd.DataFrame(rows), month, allow_missing_expenses=True, allow_missing_room_nights=True
    )


def save_daily_room_costs(month: str, daily_df: pd.DataFrame, storage_dir: Optional[str] = None) -> None:
    """Atomically save validated daily costs in a month-specific JSON file."""
    month = _month_key(month)
    data = _normalise_daily_frame(
        daily_df, month, allow_missing_expenses=True, allow_missing_room_nights=True
    )
    target_dir = get_room_cost_storage_dir(storage_dir)
    os.makedirs(target_dir, exist_ok=True)
    target_path = os.path.join(target_dir, f"{month}.json")
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", dir=target_dir, delete=False, suffix=".tmp"
        ) as file_obj:
            serializable = data.astype(object).where(pd.notna(data), None)
            json.dump(serializable.to_dict(orient="records"), file_obj, ensure_ascii=True, indent=2)
            temp_path = file_obj.name
        os.replace(temp_path, target_path)
    finally:
        if temp_path and os.path.exists(temp_path):
            os.unlink(temp_path)


def merge_daily_room_costs(
    saved_df: pd.DataFrame,
    imported_df: pd.DataFrame,
    replace_dates: bool = False,
    month: Optional[str] = None,
) -> pd.DataFrame:
    """Merge imported dates into saved rows, optionally replacing duplicates."""
    if month is None:
        if imported_df.empty and saved_df.empty:
            raise ValueError("Month is required when both datasets are empty.")
        date_values = pd.concat([saved_df.get("Date", pd.Series(dtype=str)), imported_df.get("Date", pd.Series(dtype=str))])
        month = str(date_values.iloc[0])[:7]
    month = _month_key(month)
    saved = _normalise_daily_frame(
        saved_df, month, allow_missing_expenses=True, allow_missing_room_nights=True
    )
    imported = _normalise_daily_frame(
        imported_df, month, allow_missing_expenses=True, allow_missing_room_nights=True
    )
    duplicates = set(saved["Date"]) & set(imported["Date"])
    if duplicates and not replace_dates:
        dates = ", ".join(sorted(duplicates))
        raise ValueError(f"Date(s) already saved; confirm replacement before importing: {dates}.")
    preserved = saved[saved["Date"].isin(duplicates)].set_index("Date")
    replacements = imported.set_index("Date")
    for column in DAILY_COLUMNS[1:]:
        if column in preserved.columns:
            blank_values = replacements[column].isna()
            shared_dates = replacements.index.intersection(preserved.index)
            replacements.loc[shared_dates[blank_values.loc[shared_dates]], column] = preserved.loc[
                shared_dates[blank_values.loc[shared_dates]], column
            ]
    replacements = replacements.reset_index()
    combined = pd.concat(
        [saved[~saved["Date"].isin(duplicates)], replacements], ignore_index=True
    )
    return combined.sort_values("Date").reset_index(drop=True)


def calculate_daily_room_costs(
    sr_df: pd.DataFrame,
    daily_df: pd.DataFrame,
    month: str,
) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    """Calculate per-day costs and the current retroactive MTD SR rate."""
    month = _month_key(month)
    daily = _normalise_daily_frame(
        daily_df, month, allow_missing_expenses=True, allow_missing_room_nights=True
    )
    if daily.empty:
        empty = pd.DataFrame()
        return empty, empty, []

    warnings = []
    missing_expenses = daily[EXPENSE_COLUMNS].isna()
    for column in EXPENSE_COLUMNS:
        missing_count = int(missing_expenses[column].sum())
        if missing_count:
            warnings.append(
                f"{column.replace('_', ' ')} is missing for {missing_count} date(s); "
                "combined room cost will remain incomplete until those amounts are entered."
            )
    missing_room_nights = int(daily["Room_Nights"].isna().sum())
    if missing_room_nights:
        warnings.append(
            f"Room sold is missing for {missing_room_nights} date(s); per-room calculations remain incomplete for those dates."
        )
    report_end = pd.Timestamp(daily["Date"].max())
    expected_dates = set(pd.date_range(f"{month}-01", report_end, freq="D").strftime("%Y-%m-%d"))
    missing_dates = sorted(expected_dates - set(daily["Date"]))
    if missing_dates:
        warnings.append(
            f"Room-night input is missing {len(missing_dates)} date(s) through {report_end:%Y-%m-%d}; "
            "the MTD SR rate uses only saved dates."
        )
    sr_by_date = pd.Series(dtype=float)
    if sr_df is not None and not sr_df.empty:
        if "Date" not in sr_df.columns or "Total" not in sr_df.columns:
            raise ValueError("SR data must contain Date and Total columns.")
        sr_dates = pd.to_datetime(sr_df["Date"], errors="coerce", format="mixed")
        sr_totals = pd.to_numeric(sr_df["Total"], errors="coerce")
        invalid_sr = sr_dates.isna() | sr_totals.isna()
        if invalid_sr.any():
            warnings.append(f"Ignored {int(invalid_sr.sum())} SR row(s) with missing or invalid dates/amounts.")
        valid = ~invalid_sr
        outside_month = valid & (sr_dates.dt.strftime("%Y-%m") != month)
        if outside_month.any():
            warnings.append(f"Ignored {int(outside_month.sum())} SR row(s) outside {month}.")
        valid &= ~outside_month
        after_report_end = valid & (sr_dates.dt.normalize() > report_end)
        if after_report_end.any():
            warnings.append(f"Ignored {int(after_report_end.sum())} SR row(s) after the latest daily input date ({report_end:%Y-%m-%d}).")
        valid &= ~after_report_end
        if (sr_totals[valid] < 0).any():
            raise ValueError("SR totals cannot be negative.")
        if valid.any():
            sr_by_date = sr_totals[valid].groupby(sr_dates[valid].dt.strftime("%Y-%m-%d")).sum()

    result = daily.copy()
    result["SR_Received_On_Date"] = result["Date"].map(sr_by_date).fillna(0.0)
    mtd_sr_spend = float(sr_by_date.sum())
    mtd_room_nights = int(result["Room_Nights"].sum())
    sr_rate = mtd_sr_spend / mtd_room_nights

    for column in EXPENSE_COLUMNS:
        result[f"{column}_Per_Room"] = result[column] / result["Room_Nights"]
    result["MTD_SR_Spend"] = mtd_sr_spend
    result["MTD_Room_Nights"] = mtd_room_nights
    result["SR_Per_Room"] = sr_rate
    expense_complete = result[EXPENSE_COLUMNS].notna().all(axis=1)
    room_cost_complete = expense_complete & result["Room_Nights"].notna()
    result["Total_Cost_Per_Room"] = (
        result[[f"{column}_Per_Room" for column in EXPENSE_COLUMNS]].sum(axis=1, min_count=len(EXPENSE_COLUMNS))
        + sr_rate
    )
    result.loc[~room_cost_complete, "Total_Cost_Per_Room"] = float("nan")
    result["Total_Room_Cost"] = (
        result[EXPENSE_COLUMNS].sum(axis=1, min_count=len(EXPENSE_COLUMNS))
        + result["SR_Per_Room"] * result["Room_Nights"]
    )
    result.loc[~room_cost_complete, "Total_Room_Cost"] = float("nan")

    monthly_total_room_cost = (
        float(result["Total_Room_Cost"].sum()) if room_cost_complete.all() else float("nan")
    )

    monthly = pd.DataFrame([{
        "Month": month,
        "Room_Nights": mtd_room_nights,
        "Occupancy_Average": float(result["Occupancy"].mean()) if result["Occupancy"].notna().any() else float("nan"),
        "Water_Total": float(result["Water"].sum(min_count=1)) if result["Water"].notna().any() else float("nan"),
        "Electricity_Total": float(result["Electricity"].sum(min_count=1)) if result["Electricity"].notna().any() else float("nan"),
        "Linen_Laundry_Total": float(result["Linen_Laundry"].sum(min_count=1)) if result["Linen_Laundry"].notna().any() else float("nan"),
        "SR_Total": mtd_sr_spend,
        "Total_Room_Cost": monthly_total_room_cost,
        "Water_Per_Room": float(result.loc[result["Room_Nights"].notna(), "Water"].sum(min_count=1)) / mtd_room_nights if result["Room_Nights"].notna().all() and result["Water"].notna().all() else float("nan"),
        "Electricity_Per_Room": float(result.loc[result["Room_Nights"].notna(), "Electricity"].sum(min_count=1)) / mtd_room_nights if result["Room_Nights"].notna().all() and result["Electricity"].notna().all() else float("nan"),
        "Linen_Laundry_Per_Room": float(result.loc[result["Room_Nights"].notna(), "Linen_Laundry"].sum(min_count=1)) / mtd_room_nights if result["Room_Nights"].notna().all() and result["Linen_Laundry"].notna().all() else float("nan"),
        "SR_Per_Room": sr_rate,
        "Total_Cost_Per_Room": monthly_total_room_cost / mtd_room_nights,
    }])
    return result, monthly, warnings