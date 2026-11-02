"""SQLite persistence for daily room-cost inputs, manual briefing, and migrations.

Split out of the former monolith room_cost.py. Owns storage-dir resolution,
schema creation, legacy-JSON/Excel import into SQLite, load/save daily costs,
and manual briefing note + entry persistence.
"""
import json
import os
import re
import hashlib
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Optional

import openpyxl
import pandas as pd

from .parsers import DAILY_COLUMNS, _month_key, _normalise_daily_frame


def get_room_cost_storage_dir(storage_dir: Optional[str] = None) -> str:
    """Return the app-owned directory for persisted room-cost inputs."""
    path = storage_dir or os.environ.get("HOTEL_ROOM_COST_DIR")
    if not path:
        path = os.path.join(Path.home(), ".hotel-expense-tracker", "room_cost")
    return os.path.abspath(os.path.expanduser(path))


def get_room_cost_database_path(storage_dir: Optional[str] = None) -> str:
    """Return the SQLite database path used for operational data storage."""
    storage_path = get_room_cost_storage_dir(storage_dir)
    os.makedirs(storage_path, exist_ok=True)
    return os.path.join(storage_path, "hotel_data.sqlite3")


def _connect_room_cost_database(storage_dir: Optional[str] = None) -> sqlite3.Connection:
    connection = sqlite3.connect(get_room_cost_database_path(storage_dir))
    connection.execute("PRAGMA foreign_keys = ON")
    connection.execute("""
        CREATE TABLE IF NOT EXISTS daily_room_costs (
            date TEXT PRIMARY KEY,
            month TEXT NOT NULL,
            occupancy REAL,
            room_nights INTEGER,
            water REAL,
            electricity REAL,
            linen_laundry REAL
        )
    """)
    connection.execute("""
        CREATE TABLE IF NOT EXISTS imported_files (
            sha256 TEXT PRIMARY KEY,
            source_path TEXT NOT NULL,
            source_name TEXT NOT NULL,
            imported_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            record_count INTEGER NOT NULL DEFAULT 0
        )
    """)
    connection.execute("""
        CREATE TABLE IF NOT EXISTS imported_rows (
            sha256 TEXT NOT NULL,
            sheet_name TEXT NOT NULL,
            row_number INTEGER NOT NULL,
            row_data TEXT NOT NULL,
            PRIMARY KEY (sha256, sheet_name, row_number),
            FOREIGN KEY (sha256) REFERENCES imported_files(sha256) ON DELETE CASCADE
        )
    """)
    connection.execute("""
        CREATE TABLE IF NOT EXISTS briefing_notes (
            date TEXT PRIMARY KEY,
            staff_incharge TEXT NOT NULL DEFAULT '',
            remarks TEXT NOT NULL DEFAULT ''
        )
    """)
    connection.execute("""
        CREATE TABLE IF NOT EXISTS manual_daily_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            entry_type TEXT NOT NULL,
            staff_name TEXT NOT NULL DEFAULT '',
            section TEXT NOT NULL DEFAULT '',
            room_credit REAL,
            area TEXT NOT NULL DEFAULT '',
            description TEXT NOT NULL DEFAULT '',
            qty REAL,
            unit TEXT NOT NULL DEFAULT '',
            unit_cost REAL,
            total REAL,
            status TEXT NOT NULL DEFAULT '',
            pic TEXT NOT NULL DEFAULT '',
            notes TEXT NOT NULL DEFAULT ''
        )
    """)
    connection.commit()
    return connection


def _upsert_room_cost_rows(
    connection: sqlite3.Connection,
    data: pd.DataFrame,
    *,
    ignore_existing: bool = False,
) -> int:
    if data.empty:
        return 0
    command = "INSERT OR IGNORE" if ignore_existing else "INSERT OR REPLACE"
    sql = f"""{command} INTO daily_room_costs
        (date, month, occupancy, room_nights, water, electricity, linen_laundry)
        VALUES (?, ?, ?, ?, ?, ?, ?)"""
    records = []
    for row in data.to_dict(orient="records"):
        records.append((
            row["Date"], row["Date"][:7],
            None if pd.isna(row["Occupancy"]) else float(row["Occupancy"]),
            None if pd.isna(row["Room_Nights"]) else int(row["Room_Nights"]),
            None if pd.isna(row["Water"]) else float(row["Water"]),
            None if pd.isna(row["Electricity"]) else float(row["Electricity"]),
            None if pd.isna(row["Linen_Laundry"]) else float(row["Linen_Laundry"]),
        ))
    before = connection.total_changes
    connection.executemany(sql, records)
    return connection.total_changes - before


def _legacy_daily_room_costs(month: str, storage_dir: Optional[str] = None) -> pd.DataFrame:
    path = os.path.join(get_room_cost_storage_dir(storage_dir), f"{_month_key(month)}.json")
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


def _import_workbook_rows(connection: sqlite3.Connection, source_path: str) -> dict:
    path = os.path.abspath(os.path.expanduser(source_path))
    digest = hashlib.sha256()
    with open(path, "rb") as source_file:
        for chunk in iter(lambda: source_file.read(1024 * 1024), b""):
            digest.update(chunk)
    file_hash = digest.hexdigest()
    if connection.execute("SELECT 1 FROM imported_files WHERE sha256 = ?", (file_hash,)).fetchone():
        return {"file": path, "imported": False, "rows": 0}

    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    records = []
    try:
        for worksheet in workbook.worksheets:
            for row_number, values in enumerate(worksheet.iter_rows(values_only=True), 1):
                if not any(value is not None for value in values):
                    continue
                serializable = []
                for value in values:
                    if value is None or (not isinstance(value, (str, bytes)) and pd.isna(value)):
                        serializable.append(None)
                    elif isinstance(value, (datetime, pd.Timestamp)):
                        serializable.append(pd.Timestamp(value).isoformat())
                    elif hasattr(value, "item"):
                        serializable.append(value.item())
                    else:
                        serializable.append(value)
                records.append((file_hash, worksheet.title, row_number, json.dumps(serializable, ensure_ascii=True)))
    finally:
        workbook.close()

    connection.execute(
        "INSERT INTO imported_files (sha256, source_path, source_name, record_count) VALUES (?, ?, ?, ?)",
        (file_hash, path, os.path.basename(path), len(records)),
    )
    connection.executemany(
        "INSERT INTO imported_rows (sha256, sheet_name, row_number, row_data) VALUES (?, ?, ?, ?)",
        records,
    )
    return {"file": path, "imported": True, "rows": len(records)}


def migrate_existing_data_to_database(
    month: str,
    source_files: Optional[list[str]] = None,
    storage_dir: Optional[str] = None,
) -> dict:
    """Copy existing daily JSON and selected workbook rows into the local SQLite store."""
    month = _month_key(month)
    legacy_data = _legacy_daily_room_costs(month, storage_dir)
    connection = _connect_room_cost_database(storage_dir)
    try:
        room_cost_rows = _upsert_room_cost_rows(connection, legacy_data, ignore_existing=True)
        workbook_results = []
        for source_path in source_files or []:
            if not os.path.isfile(os.path.expanduser(source_path)):
                raise FileNotFoundError(f"Source workbook not found: {source_path}")
            workbook_results.append(_import_workbook_rows(connection, source_path))
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()
    return {
        "database": get_room_cost_database_path(storage_dir),
        "room_cost_rows_imported": room_cost_rows,
        "workbooks": workbook_results,
    }


def load_manual_briefing_data(report_date, storage_dir: Optional[str] = None) -> dict:
    """Load manual briefing fields and related rows for one date."""
    date_value = pd.Timestamp(report_date).strftime("%Y-%m-%d")
    connection = _connect_room_cost_database(storage_dir)
    try:
        note = connection.execute(
            "SELECT staff_incharge, remarks FROM briefing_notes WHERE date = ?", (date_value,)
        ).fetchone()
        entries = connection.execute(
            "SELECT entry_type, staff_name, section, room_credit, area, description, qty, unit, "
            "unit_cost, total, status, pic, notes FROM manual_daily_entries WHERE date = ? ORDER BY id",
            (date_value,),
        ).fetchall()
    finally:
        connection.close()

    entry_columns = [
        "Entry_Type", "Staff_Name", "Section", "Room_Credit", "Area", "Description",
        "Qty", "Unit", "Unit_Cost", "Total", "Status", "PIC", "Notes",
    ]
    entries_frame = pd.DataFrame(entries, columns=entry_columns)
    return {
        "Staff_Incharge": note[0] if note else "",
        "Remarks": note[1] if note else "",
        "Entries": entries_frame,
    }


def save_manual_briefing_data(
    report_date,
    *,
    staff_incharge: str = "",
    remarks: str = "",
    entries: Optional[pd.DataFrame] = None,
    storage_dir: Optional[str] = None,
) -> None:
    """Replace manual briefing entries for a date in one SQLite transaction."""
    date_value = pd.Timestamp(report_date).strftime("%Y-%m-%d")
    allowed_types = {"ra", "project_pa", "daily_pa", "guest_laundry"}
    records = []
    if entries is not None and not entries.empty:
        for row in entries.to_dict(orient="records"):
            entry_type = str(row.get("Entry_Type", "")).strip()
            if entry_type not in allowed_types:
                raise ValueError(f"Unsupported manual entry type: {entry_type}")
            text = lambda key: "" if pd.isna(row.get(key)) else str(row.get(key)).strip()
            number = lambda key: None if pd.isna(row.get(key)) or row.get(key) == "" else float(row[key])
            qty = number("Qty")
            unit_cost = number("Unit_Cost")
            total = number("Total")
            if total is None and qty is not None and unit_cost is not None:
                total = qty * unit_cost
            records.append((
                date_value, entry_type, text("Staff_Name"), text("Section"),
                number("Room_Credit"), text("Area"), text("Description"), qty,
                text("Unit"), unit_cost, total, text("Status"), text("PIC"), text("Notes"),
            ))

    connection = _connect_room_cost_database(storage_dir)
    try:
        with connection:
            connection.execute(
                "INSERT INTO briefing_notes (date, staff_incharge, remarks) VALUES (?, ?, ?) "
                "ON CONFLICT(date) DO UPDATE SET staff_incharge = excluded.staff_incharge, remarks = excluded.remarks",
                (date_value, str(staff_incharge or "").strip(), str(remarks or "").strip()),
            )
            connection.execute("DELETE FROM manual_daily_entries WHERE date = ?", (date_value,))
            connection.executemany(
                "INSERT INTO manual_daily_entries "
                "(date, entry_type, staff_name, section, room_credit, area, description, qty, unit, unit_cost, total, status, pic, notes) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                records,
            )
    finally:
        connection.close()


def load_guest_laundry_data(month: str, storage_dir: Optional[str] = None) -> pd.DataFrame:
    """Load stored guest-laundry room counts and costs for one month."""
    month = _month_key(month)
    connection = _connect_room_cost_database(storage_dir)
    try:
        rows = connection.execute(
            "SELECT date, SUM(qty), SUM(total), GROUP_CONCAT(NULLIF(notes, ''), ', ') "
            "FROM manual_daily_entries "
            "WHERE entry_type = 'guest_laundry' AND date LIKE ? GROUP BY date ORDER BY date",
            (f"{month}-%",),
        ).fetchall()
    finally:
        connection.close()
    return pd.DataFrame(rows, columns=["Date", "Room_Nights", "Guest_Laundry_Cost", "Room_Numbers"])


def save_guest_laundry_data(
    month: str,
    data: pd.DataFrame,
    *,
    replace_dates: bool = False,
    storage_dir: Optional[str] = None,
) -> None:
    """Replace imported guest-laundry rows for dates, leaving other entry types intact."""
    month = _month_key(month)
    required = {"Date", "Room_Nights", "Guest_Laundry_Cost"}
    if not required.issubset(data.columns):
        raise ValueError("Drop N Go data must include Date, Room_Nights, and Guest_Laundry_Cost.")
    if data.empty:
        raise ValueError("No Drop N Go rows to save.")

    normalized = data[["Date", "Room_Nights", "Guest_Laundry_Cost"]].copy()
    normalized["Room_Numbers"] = data.get("Room_Numbers", "")
    parsed_dates = pd.to_datetime(normalized["Date"], errors="coerce", format="mixed")
    if parsed_dates.isna().any():
        raise ValueError("Drop N Go contains an invalid date.")
    normalized["Date"] = parsed_dates.dt.strftime("%Y-%m-%d")
    if (normalized["Date"].str[:7] != month).any():
        raise ValueError(f"Drop N Go contains dates outside selected month {month}.")
    if normalized["Date"].duplicated().any():
        raise ValueError("Drop N Go contains duplicate dates.")

    room_nights = pd.to_numeric(normalized["Room_Nights"], errors="coerce")
    costs = pd.to_numeric(normalized["Guest_Laundry_Cost"], errors="coerce")
    if room_nights.isna().any() or (room_nights <= 0).any() or (room_nights % 1 != 0).any():
        raise ValueError("Drop N Go room counts must be positive whole numbers.")
    if costs.isna().any() or (costs < 0).any():
        raise ValueError("Drop N Go costs must be non-negative numbers.")
    normalized["Room_Nights"] = room_nights.astype(int)
    normalized["Guest_Laundry_Cost"] = costs.astype(float)

    connection = _connect_room_cost_database(storage_dir)
    try:
        dates = normalized["Date"].tolist()
        placeholders = ",".join("?" for _ in dates)
        existing = {
            row[0] for row in connection.execute(
                f"SELECT DISTINCT date FROM manual_daily_entries "
                f"WHERE entry_type = 'guest_laundry' AND date IN ({placeholders})",
                dates,
            ).fetchall()
        }
        duplicates = existing.intersection(dates)
        if duplicates and not replace_dates:
            raise ValueError(
                "Drop N Go data already exists for: " + ", ".join(sorted(duplicates))
                + ". Confirm replacement before importing."
            )
        with connection:
            for row in normalized.itertuples(index=False):
                room_numbers = getattr(row, "Room_Numbers", "")
                if pd.isna(room_numbers):
                    room_numbers = ""
                connection.execute(
                    "DELETE FROM manual_daily_entries WHERE date = ? AND entry_type = 'guest_laundry'",
                    (row.Date,),
                )
                connection.execute(
                    "INSERT INTO manual_daily_entries "
                    "(date, entry_type, qty, total, notes) VALUES (?, 'guest_laundry', ?, ?, ?)",
                    (row.Date, int(row.Room_Nights), float(row.Guest_Laundry_Cost), str(room_numbers)),
                )
    finally:
        connection.close()


def load_daily_room_costs(month: str, storage_dir: Optional[str] = None) -> pd.DataFrame:
    """Load daily costs from SQLite, importing legacy month JSON if needed."""
    month = _month_key(month)
    legacy_data = _legacy_daily_room_costs(month, storage_dir)
    connection = _connect_room_cost_database(storage_dir)
    try:
        _upsert_room_cost_rows(connection, legacy_data, ignore_existing=True)
        rows = connection.execute(
            "SELECT date, occupancy, room_nights, water, electricity, linen_laundry "
            "FROM daily_room_costs WHERE month = ? ORDER BY date",
            (month,),
        ).fetchall()
        connection.commit()
    finally:
        connection.close()
    frame = pd.DataFrame(rows, columns=DAILY_COLUMNS)
    return _normalise_daily_frame(
        frame, month, allow_missing_expenses=True, allow_missing_room_nights=True
    )


def save_daily_room_costs(month: str, daily_df: pd.DataFrame, storage_dir: Optional[str] = None) -> None:
    """Persist validated daily costs in SQLite, retaining legacy JSON as backup."""
    month = _month_key(month)
    data = _normalise_daily_frame(
        daily_df, month, allow_missing_expenses=True, allow_missing_room_nights=True
    )
    connection = _connect_room_cost_database(storage_dir)
    try:
        _upsert_room_cost_rows(connection, data)
        connection.commit()
    finally:
        connection.close()
