"""Daily room-cost calculations: MTD occupancy, merging, and per-room cost.

Split out of the former monolith room_cost.py. Pure pandas logic, no I/O.
"""
from typing import Optional

import pandas as pd

from .parsers import DAILY_COLUMNS, EXPENSE_COLUMNS, _month_key, _normalise_daily_frame


def calculate_mtd_occupancy_and_room_nights(
    daily_df: pd.DataFrame,
    month: str,
) -> tuple[float, int]:
    """Return average available occupancy and cumulative room nights for a month."""
    daily = _normalise_daily_frame(
        daily_df, month, allow_missing_expenses=True, allow_missing_room_nights=True
    )
    occupancy = pd.to_numeric(daily["Occupancy"], errors="coerce")
    room_nights = pd.to_numeric(daily["Room_Nights"], errors="coerce")
    occupancy_mtd = float(occupancy.mean()) if occupancy.notna().any() else float("nan")
    room_nights_mtd = int(room_nights.sum()) if room_nights.notna().any() else 0
    return occupancy_mtd, room_nights_mtd


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
    sr_category_columns = {
        "Guest Supplies": "Guest_Supplies_Per_Room",
        "Cleaning Supplies": "Cleaning_Supplies_Per_Room",
        "Paper Supplies": "Paper_Supplies_Per_Room",
        "Print & Stationery": "Print_Stationery_Per_Room",
        "Other": "Other_SR_Per_Room",
    }
    category_totals = {category: 0.0 for category in sr_category_columns}
    if sr_df is not None and not sr_df.empty and "Category" in sr_df.columns and valid.any():
        normalized_categories = sr_df.loc[valid, "Category"].fillna("").astype(str).str.strip().str.lower()
        category_aliases = {
            "guest supplies": "Guest Supplies",
            "cleaning supplies": "Cleaning Supplies",
            "paper supplies": "Paper Supplies",
            "print": "Print & Stationery",
            "print & stationery": "Print & Stationery",
        }
        categorized = pd.DataFrame({
            "Category": normalized_categories.map(category_aliases).fillna("Other"),
            "Total": sr_totals.loc[valid],
        })
        category_totals.update(categorized.groupby("Category")["Total"].sum().to_dict())
    else:
        category_totals["Other"] = mtd_sr_spend
    for category, column in sr_category_columns.items():
        result[column] = category_totals[category] / mtd_room_nights
    expense_complete = result[EXPENSE_COLUMNS].notna().all(axis=1)
    room_cost_complete = expense_complete & result["Room_Nights"].notna()
    result["Total_Cost_Per_Room"] = (
        result[[*sr_category_columns.values(), *(f"{column}_Per_Room" for column in EXPENSE_COLUMNS)]].sum(
            axis=1, min_count=len(sr_category_columns) + len(EXPENSE_COLUMNS)
        )
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
