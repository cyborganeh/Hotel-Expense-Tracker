"""room_cost - compatibility shim for the split room-cost package.

Keeps the public surface used by views and tests unchanged while the actual
implementation now lives in data/parsers.py (Excel parsing/normalization),
data/storage.py (SQLite persistence + migration), and data/calculator.py
(MTD/merge/per-room calculations).
"""
from data.calculator import (
    calculate_daily_room_costs,
    calculate_mtd_occupancy_and_room_nights,
    merge_daily_room_costs,
)
from data.parsers import (
    DAILY_COLUMNS,
    EXPENSE_COLUMNS,
    REQUIRED_DAILY_COLUMNS,
    _month_key,
    _normalise_daily_frame,
    parse_bonvivo_linen_workbook,
    parse_daily_linen_cost_workbook,
    parse_daily_room_cost_excel,
    parse_dropngo_guest_laundry_workbook,
    parse_dropngo_linen_workbook,
    parse_energy_cost_workbook,
    parse_night_audit_report,
)
from data.storage import (
    get_room_cost_database_path,
    get_room_cost_storage_dir,
    load_daily_room_costs,
    load_guest_laundry_data,
    load_manual_briefing_data,
    migrate_existing_data_to_database,
    save_daily_room_costs,
    save_guest_laundry_data,
    save_manual_briefing_data,
)
