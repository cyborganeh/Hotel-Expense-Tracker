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


def calculate_per_room_metrics(
    cat_df: pd.DataFrame,
    item_df: pd.DataFrame,
    room_nights: int,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Add per-room and per-100-room ratios to category and item summaries.

    Normalizes spend and item usage by room sold so months with different
    occupancy stay comparable. Returns NaN ratios when room_nights <= 0.
    """
    categories = cat_df.copy() if cat_df is not None else pd.DataFrame()
    items = item_df.copy() if item_df is not None else pd.DataFrame()
    category_columns = ["Actual_Per_Room", "Budget_Per_Room", "Variance_Per_Room"]
    item_columns = ["Qty_Per_100_Rooms", "Cost_Per_Room"]

    if room_nights <= 0:
        for frame, columns in ((categories, category_columns), (items, item_columns)):
            for column in columns:
                frame[column] = float("nan")
        return categories, items

    if not categories.empty:
        categories["Actual_Per_Room"] = categories["Actual"] / room_nights
        categories["Budget_Per_Room"] = categories["Budget"] / room_nights
        categories["Variance_Per_Room"] = categories["Variance_IDR"] / room_nights
    if not items.empty:
        items["Qty_Per_100_Rooms"] = items["Total_Qty"] / room_nights * 100.0
        items["Cost_Per_Room"] = items["Total_Amount"] / room_nights
    return categories, items

def calculate_per_room_trend(
    monthly_inputs: list[tuple[str, pd.DataFrame, int]],
) -> pd.DataFrame:
    """Build a per-room trend across months.

    monthly_inputs: chronological list of (month_label, category_summary_df,
    room_nights). Returns one row per month with spend, budget, and per-room
    ratios. Ratios are NaN when room_nights <= 0.
    """
    rows = []
    for label, cat_df, room_nights in monthly_inputs:
        has_data = cat_df is not None and not cat_df.empty
        actual = float(cat_df["Actual"].sum()) if has_data else 0.0
        budget = float(cat_df["Budget"].sum()) if has_data else 0.0
        variance = actual - budget
        valid = room_nights > 0
        rows.append({
            "Month": label,
            "Room_Nights": int(room_nights),
            "Actual": actual,
            "Budget": budget,
            "Variance_IDR": variance,
            "Actual_Per_Room": actual / room_nights if valid else float("nan"),
            "Budget_Per_Room": budget / room_nights if valid else float("nan"),
            "Variance_Per_Room": variance / room_nights if valid else float("nan"),
        })
    return pd.DataFrame(rows)

def detect_usage_spikes(
    item_frames: list[tuple[str, pd.DataFrame]],
    room_nights_by_month: dict[str, int],
    threshold_pct: float = 25.0,
    lookback: int = 3,
) -> pd.DataFrame:
    """Flag items whose usage per 100 rooms rose above their trailing average.

    item_frames must be chronological (oldest first). The latest month is
    compared against the mean of up to `lookback` preceding months. Items with
    no baseline usage are skipped because there is nothing to compare against.
    """
    columns = [
        "Category", "Item_Name", "Unit", "Latest_Month",
        "Latest_Qty_Per_100", "Baseline_Qty_Per_100", "Change_Pct", "Status",
    ]
    records = []
    for label, item_df in item_frames:
        room_nights = room_nights_by_month.get(label, 0)
        if item_df is None or item_df.empty or room_nights <= 0:
            continue
        frame = item_df[["Category", "Item_Name", "Unit", "Total_Qty"]].copy()
        frame["Month"] = label
        frame["Qty_Per_100_Rooms"] = frame["Total_Qty"] / room_nights * 100.0
        records.append(frame)
    if len(records) < 2:
        return pd.DataFrame(columns=columns)

    long_df = pd.concat(records, ignore_index=True)
    latest_month = records[-1]["Month"].iloc[0]
    baseline_months = [frame["Month"].iloc[0] for frame in records[-(lookback + 1):-1]]

    latest = long_df[long_df["Month"] == latest_month]
    baseline = (
        long_df[long_df["Month"].isin(baseline_months)]
        .groupby(["Category", "Item_Name", "Unit"], as_index=False)["Qty_Per_100_Rooms"]
        .mean()
        .rename(columns={"Qty_Per_100_Rooms": "Baseline_Qty_Per_100"})
    )
    merged = latest.merge(baseline, on=["Category", "Item_Name", "Unit"], how="inner")
    merged = merged[merged["Baseline_Qty_Per_100"] > 0].copy()
    merged["Change_Pct"] = (
        (merged["Qty_Per_100_Rooms"] - merged["Baseline_Qty_Per_100"])
        / merged["Baseline_Qty_Per_100"] * 100.0
    )
    merged = merged[merged["Change_Pct"] >= threshold_pct]
    merged = merged.rename(columns={
        "Month": "Latest_Month",
        "Qty_Per_100_Rooms": "Latest_Qty_Per_100",
    })
    merged["Status"] = "Usage spike"
    merged = merged.sort_values("Change_Pct", ascending=False)
    return merged[columns].reset_index(drop=True)

def calculate_cpor(total_spend: float, daily_df: pd.DataFrame) -> dict:
    """Combined cost per occupied room: ledger spend + utilities + linen.

    Ledger spend already contains SR consumption, so utilities and linen are
    the only additions. Components stay NaN when their daily inputs are missing.
    """
    result = {
        "Room_Nights": 0,
        "Ledger_Spend": float(total_spend),
        "Utilities": float("nan"),
        "Linen": float("nan"),
        "Total_Cost": float("nan"),
        "Cost_Per_Room": float("nan"),
        "Ledger_Per_Room": float("nan"),
        "Utilities_Per_Room": float("nan"),
        "Linen_Per_Room": float("nan"),
    }
    if daily_df is None or daily_df.empty or "Room_Nights" not in daily_df.columns:
        return result
    room_nights = int(pd.to_numeric(daily_df["Room_Nights"], errors="coerce").sum())
    if room_nights <= 0:
        return result

    water = pd.to_numeric(daily_df.get("Water"), errors="coerce")
    electricity = pd.to_numeric(daily_df.get("Electricity"), errors="coerce")
    linen = pd.to_numeric(daily_df.get("Linen_Laundry"), errors="coerce")
    utilities = water.sum(min_count=1) + electricity.sum(min_count=1)
    linen_total = linen.sum(min_count=1)
    total = float(total_spend) + utilities + linen_total
    result.update({
        "Room_Nights": room_nights,
        "Utilities": float(utilities),
        "Linen": float(linen_total),
        "Total_Cost": float(total),
        "Cost_Per_Room": float(total) / room_nights,
        "Ledger_Per_Room": float(total_spend) / room_nights,
        "Utilities_Per_Room": float(utilities) / room_nights,
        "Linen_Per_Room": float(linen_total) / room_nights,
    })
    return result

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
