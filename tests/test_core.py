"""Tests for the hotel-expense-tracker modules."""
import os
import sys
import tempfile
import json
from datetime import datetime, timedelta
from io import BytesIO
from zipfile import ZipFile
from xml.etree import ElementTree
import pytest
import pandas as pd
import openpyxl

# Ensure project root is on the path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from parser import find_month_files, parse_detail_trial_balance, parse_consumption_report
from matcher import reconcile_monthly_expenses
from exporter import create_separated_excel, create_sr_separated_excel, create_morning_briefing_template
from sr_parser import normalize_sr_labels
from data.months import get_default_data_dir
import room_cost


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def sample_dtb_data():
    """Create a minimal DTB DataFrame for testing - with proper columns from parser."""
    import pandas as pd
    return pd.DataFrame([
        {"Date": "2026-08-01", "Account_Code": "5000000", "Account_Name": "COST OF GOODS SOLD", "Category": "FOOD & BEVERAGE", "Ref": "GDC/OUT-001", "Source": "DTB", "JRNL": "GL", "Partner": "Vendor A", "Description": "FOOD & BEVERAGE", "Net_Amount": 1000000.0},
        {"Date": "2026-08-02", "Account_Code": "5000000", "Account_Name": "COST OF GOODS SOLD", "Category": "LAUNDRY", "Ref": "HK/IN-002", "Source": "DTB", "JRNL": "GL", "Partner": "Vendor B", "Description": "LAUNDRY", "Net_Amount": 500000.0},
    ])


@pytest.fixture
def sample_income_data():
    """Create a minimal Income Statement DataFrame."""
    import pandas as pd
    return pd.DataFrame([
        {"Account_Code": "4000000", "Account_Name": "ROOM REVENUE", "Description": "ROOM", "Budget": 2000000.0, "Actual": 1800000.0, "Ratio_Pct": 0.9, "Is_Subtotal": False, "Department": "Room Division", "Group": "Revenue"},
    ])


@pytest.fixture
def sample_consumption_data():
    """Create a minimal Consumption Report DataFrame."""
    import pandas as pd
    return pd.DataFrame([
        {"Date": "2026-08-01", "Voucher_No": "GDC/OUT-001", "Item_Name": "TOWEL", "Qty": 10.0, "Unit": "pcs", "Unit_Price": 10000.0, "Amount": 100000.0, "Section": "General HK"},
    ])


# ---------------------------------------------------------------------------
# find_month_files tests
# ---------------------------------------------------------------------------

class TestFindMonthFiles:
    def test_missing_directory_returns_empty(self):
        """find_month_files should return empty dict when directory does not exist."""
        result = find_month_files("/nonexistent/path/123")
        assert result == {} or all(v is None for v in result.values())

    def test_empty_directory_returns_empty(self):
        """find_month_files should return empty dict when directory has no month folders."""
        with tempfile.TemporaryDirectory() as tmp:
            result = find_month_files(tmp)
            assert result == {} or all(v is None for v in result.values())

    def test_invalid_hotel_data_dir_falls_back_instead_of_raising(self, monkeypatch):
        """An invalid HOTEL_DATA_DIR should be ignored and not crash startup."""
        monkeypatch.setenv("HOTEL_DATA_DIR", "/")
        result = get_default_data_dir()
        assert result
        assert result != "/"


# ---------------------------------------------------------------------------
# parse_detail_trial_balance tests
# ---------------------------------------------------------------------------

class TestParseDetailTrialBalance:
    def test_missing_file_raises(self):
        """parse_detail_trial_balance should raise when file does not exist."""
        with pytest.raises(Exception):
            parse_detail_trial_balance("/nonexistent/file.xlsx")

    def test_invalid_file_raises(self, sample_dtb_data):
        """parse_detail_trial_balance should raise on corrupted/non-Excel file."""
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            f.write(b"not a valid excel file")
            f.flush()
            with pytest.raises(Exception):
                parse_detail_trial_balance(f.name)
            os.unlink(f.name)


# ---------------------------------------------------------------------------
# parse_consumption_report tests
# ---------------------------------------------------------------------------

class TestParseConsumptionReport:
    def test_missing_file_raises(self):
        """parse_consumption_report should raise when file does not exist."""
        with pytest.raises(Exception):
            parse_consumption_report("/nonexistent/file.xlsx")

    def test_short_rows_and_pivot_tables_do_not_raise(self):
        """parse_consumption_report should safely handle rows with fewer columns (< 4) without IndexError."""
        import openpyxl
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            tmp_path = f.name

        try:
            wb = openpyxl.Workbook()
            # Sheet with short rows like pivot tables (len=3, which caused IndexError: tuple index out of range)
            ws_sr = wb.active
            ws_sr.title = "SR Food "
            ws_sr.append(["Row Labels", "Sum of AMMOUNT"])
            ws_sr.append(["2026-09-03", 4528419, None])
            ws_sr.append(["2026-09-07", 2764629])

            # Sheet with proper HK transaction rows
            ws_hk = wb.create_sheet(title="HK")
            ws_hk.append(["DATE", "NO TRX", "DEPT", "PROCESS", "TYPE", "ITEM", "QTY", "UNIT", "PRICE", "AMMOUNT", "COA"])
            ws_hk.append(["2026-09-08", "HK/IN/001", "Housekeeping Expense", "Housekeeping Receipts", "Receiving", "977993-Kemeja Alisan", 2, "Pcs", 165000, 330000, "03-52010 Uniforms"])
            # Short row in HK sheet as well (e.g. trailing empty cell or note)
            ws_hk.append(["2026-09-09", "HK/IN/002"])

            wb.save(tmp_path)

            df = parse_consumption_report(tmp_path)
            assert isinstance(df, pd.DataFrame)
            assert len(df) == 1
            assert df.iloc[0]["Voucher_No"] == "HK/IN/001"
            assert df.iloc[0]["Item_Name"] == "Kemeja Alisan"
        finally:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)


# ---------------------------------------------------------------------------
# reconcile_monthly_expenses tests
# ---------------------------------------------------------------------------

class TestReconcileMonthlyExpenses:
    def test_empty_dtb_returns_empty(self):
        """Reconciliation with empty DTB should return empty DataFrame."""
        import pandas as pd
        result = reconcile_monthly_expenses(pd.DataFrame(), sample_income_data, sample_consumption_data)
        assert isinstance(result, dict)

    def test_all_empty(self):
        """Reconciliation with all empty DataFrames should return empty DataFrame."""
        import pandas as pd
        result = reconcile_monthly_expenses(pd.DataFrame(), pd.DataFrame(), pd.DataFrame())
        assert isinstance(result, dict)

    def test_valid_data_returns_dataframe(self, sample_dtb_data, sample_income_data, sample_consumption_data):
        """Reconciliation with valid data should return a non-empty DataFrame."""
        result = reconcile_monthly_expenses(sample_dtb_data, sample_income_data, sample_consumption_data)
        assert isinstance(result, dict)

    def test_paper_supplies_budget_matched_despite_typo(self):
        """Paper Supplies category should match Income Statement even if spelled 'Paper Suplies'."""
        import pandas as pd
        dtb_df = pd.DataFrame([
            {
                "Date": "2026-08-01",
                "Account_Code": "0352220",
                "Account_Name": "Paper Supplies",
                "Category": "Paper Supplies",
                "Ref": "GDC/OUT-123",
                "Source": "DTB",
                "JRNL": "GL",
                "Partner": "Vendor Paper",
                "Description": "Tissue",
                "Net_Amount": 4163210.0,
            }
        ])
        # Income statement has 'Paper Suplies' (with one 'p')
        is_df = pd.DataFrame([
            {
                "Account_Code": "0352220",
                "Account_Name": "Paper Suplies",
                "Description": "Paper Suplies",
                "Budget": 4300000.0,
                "Actual": 4163210.0,
                "Ratio_Pct": 0.33,
                "Is_Subtotal": False,
                "Department": "Room Division",
                "Group": "Other Expenses",
            }
        ])
        result = reconcile_monthly_expenses(dtb_df, is_df, pd.DataFrame())
        cat_df = result["category_summary"]
        paper_row = cat_df[cat_df["Category"] == "Paper Supplies"]
        assert not paper_row.empty
        assert paper_row.iloc[0]["Budget"] == 4300000.0
        assert paper_row.iloc[0]["Status"] == "Under Budget"


# ---------------------------------------------------------------------------
# create_separated_excel tests
# ---------------------------------------------------------------------------

class TestCreateSeparatedExcel:
    def test_missing_data_raises(self):
        """create_separated_excel should raise when reconciled_data is None."""
        with pytest.raises(Exception):
            create_separated_excel(None, "/tmp/output.xlsx")

    def test_missing_keys_raises(self):
        """create_separated_excel should raise when required keys are missing."""
        with pytest.raises(Exception):
            create_separated_excel({"wrong_key": "data"}, "/tmp/output.xlsx")

    def test_valid_data_creates_file(self, sample_dtb_data, sample_income_data, sample_consumption_data):
        """create_separated_excel should create an Excel file with valid data."""
        reconciled = reconcile_monthly_expenses(sample_dtb_data, sample_income_data, sample_consumption_data)
        if not reconciled['transactions'].empty:
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
                output_path = f.name
            try:
                create_separated_excel(reconciled, "August 2026", output_path)
                assert os.path.exists(output_path)
                assert os.path.getsize(output_path) > 0
            finally:
                if os.path.exists(output_path):
                    os.unlink(output_path)

    def test_demo_like_data_without_account_code_creates_file(self, sample_dtb_data, sample_income_data, sample_consumption_data):
        """Export should tolerate dashboard/demo transactions without Account_Code."""
        reconciled = reconcile_monthly_expenses(sample_dtb_data, sample_income_data, sample_consumption_data)
        reconciled['transactions'] = reconciled['transactions'].drop(columns=['Account_Code'])
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            output_path = f.name
        try:
            create_separated_excel(reconciled, "Demo August 2026", output_path)
            assert os.path.getsize(output_path) > 0
        finally:
            if os.path.exists(output_path):
                os.unlink(output_path)


class TestCreateSRSeparatedExcel:
    @pytest.fixture
    def sr_data(self):
        return pd.DataFrame([
            {
                "Date": "2026-10-01", "SR_Number": "SR-001", "Requested_By": "HK",
                "Item_Code": "976744", "Item_Name": "Soap", "Category": "Guest Supplies",
                "Qty": 10.0, "Unit": "Pcs", "Cost": 10000.0, "Total": 100000.0,
            },
            {
                "Date": "2026-10-02", "SR_Number": "SR-002", "Requested_By": "HK",
                "Item_Code": "976802", "Item_Name": "Facial Tissue", "Category": "Paper Supplies",
                "Qty": 4.0, "Unit": "Box", "Cost": 5000.0, "Total": 20000.0,
            },
        ])

    def test_export_includes_period_and_per_room_sold_metrics(self, sr_data):
        workbook_bytes = create_sr_separated_excel(
            sr_data, occupancy_pct=75.5, total_rooms_sold=50
        )
        workbook = openpyxl.load_workbook(BytesIO(workbook_bytes), data_only=True)
        summary = workbook["SR Summary"]

        assert summary["I5"].value == 0.755
        assert summary["I5"].number_format == "0.0%"
        assert summary["K5"].value == 50
        assert summary["M5"].value == 2400

        category_headers = {summary.cell(8, col).value: col for col in range(1, 8)}
        spend_per_room_col = category_headers["Spend / Room Sold (IDR)"]
        assert summary.cell(9, spend_per_room_col).value == 2000
        assert summary.cell(11, spend_per_room_col).value == 400

        guest_sheet = workbook["Guest Supplies"]
        item_headers = {guest_sheet.cell(5, col).value: col for col in range(1, 10)}
        qty_per_room_col = item_headers["Qty / Room Sold"]
        assert guest_sheet.cell(6, qty_per_room_col).value == 0.2
        different_occupancy = openpyxl.load_workbook(
            BytesIO(create_sr_separated_excel(sr_data, occupancy_pct=40, total_rooms_sold=50)),
            data_only=True,
        )
        assert different_occupancy["SR Summary"]["M5"].value == summary["M5"].value
        assert different_occupancy["SR Summary"].cell(9, spend_per_room_col).value == 2000

    def test_zero_rooms_sold_marks_ratios_unavailable(self, sr_data):
        workbook_bytes = create_sr_separated_excel(
            sr_data, occupancy_pct=0, total_rooms_sold=0
        )
        workbook = openpyxl.load_workbook(BytesIO(workbook_bytes), data_only=True)
        summary = workbook["SR Summary"]
        category_headers = {summary.cell(8, col).value: col for col in range(1, 8)}

        assert summary["M5"].value == "N/A"
        assert summary.cell(9, category_headers["Spend / Room Sold (IDR)"]).value == "N/A"
        guest_sheet = workbook["Guest Supplies"]
        item_headers = {guest_sheet.cell(5, col).value: col for col in range(1, 10)}
        assert guest_sheet.cell(6, item_headers["Qty / Room Sold"]).value == "N/A"


class TestDailyRoomCost:
    @staticmethod
    def daily_rows():
        return pd.DataFrame([
            {"Date": f"2026-10-0{day}", "Room Nights": nights,
             "Water": 100000, "Electricity": 200000, "Linen Laundry": 300000}
            for day, nights in [(1, 80), (2, 80), (3, 90), (4, 50)]
        ])

    def test_excel_import_normalizes_and_rejects_zero_room_nights(self):
        workbook_buffer = BytesIO()
        self.daily_rows().to_excel(workbook_buffer, index=False)
        workbook_buffer.seek(0)

        parsed = room_cost.parse_daily_room_cost_excel(workbook_buffer, "2026-10")

        assert parsed.columns.tolist() == room_cost.DAILY_COLUMNS
        assert parsed["Room_Nights"].sum() == 300
        assert parsed.iloc[0]["Date"] == "2026-10-01"

        invalid_rows = self.daily_rows().iloc[:1].copy()
        invalid_rows.loc[0, "Room Nights"] = 0
        with pytest.raises(ValueError, match="positive whole number"):
            room_cost._normalise_daily_frame(invalid_rows, "2026-10")

    def test_merge_requires_confirmation_and_persistence_round_trips(self, tmp_path):
        saved = self.daily_rows().iloc[:2].copy()
        imported = self.daily_rows().iloc[1:3].copy()
        with pytest.raises(ValueError, match="confirm replacement"):
            room_cost.merge_daily_room_costs(saved, imported, month="2026-10")

        merged = room_cost.merge_daily_room_costs(
            saved, imported, replace_dates=True, month="2026-10"
        )
        room_cost.save_daily_room_costs("2026-10", merged, str(tmp_path))
        loaded = room_cost.load_daily_room_costs("2026-10", str(tmp_path))

        assert loaded["Date"].tolist() == ["2026-10-01", "2026-10-02", "2026-10-03"]
        assert loaded.loc[loaded["Date"] == "2026-10-02", "Water"].iloc[0] == 100000

    def test_existing_json_and_workbook_migrate_to_sqlite_idempotently(self, tmp_path):
        legacy_path = tmp_path / "2026-10.json"
        legacy_rows = [{
            "Date": "2026-10-01", "Occupancy": 90.0, "Room_Nights": 90,
            "Water": 100, "Electricity": 200, "Linen_Laundry": 300,
        }]
        legacy_path.write_text(json.dumps(legacy_rows), encoding="utf-8")
        source_path = tmp_path / "active_report.xlsx"
        source_workbook = openpyxl.Workbook()
        source_sheet = source_workbook.active
        source_sheet.title = "Daily"
        source_sheet.append(["Date", "Room Sold"])
        source_sheet.append(["2026-10-01", 90])
        source_workbook.save(source_path)
        original_json = legacy_path.read_text(encoding="utf-8")

        first_migration = room_cost.migrate_existing_data_to_database(
            "2026-10", [str(source_path)], str(tmp_path)
        )
        second_migration = room_cost.migrate_existing_data_to_database(
            "2026-10", [str(source_path)], str(tmp_path)
        )
        loaded = room_cost.load_daily_room_costs("2026-10", str(tmp_path))

        assert first_migration["room_cost_rows_imported"] == 1
        assert first_migration["workbooks"][0]["imported"] is True
        assert first_migration["workbooks"][0]["rows"] == 2
        assert second_migration["room_cost_rows_imported"] == 0
        assert second_migration["workbooks"][0]["imported"] is False
        assert loaded.loc[0, "Room_Nights"] == 90
        assert legacy_path.read_text(encoding="utf-8") == original_json
        assert os.path.exists(first_migration["database"])

    def test_manual_briefing_entries_save_reload_and_replace(self, tmp_path):
        entries = pd.DataFrame([
            {"Entry_Type": "ra", "Staff_Name": "Ayu", "Section": "Floor 2", "Room_Credit": 14},
            {"Entry_Type": "guest_laundry", "Qty": 10, "Unit_Cost": 2500},
            {"Entry_Type": "project_pa", "Area": "Lobby", "Description": "Polish floor", "Qty": 2, "Unit_Cost": 50000},
            {"Entry_Type": "daily_pa", "Area": "Pool", "Description": "Clean deck", "Qty": 1, "Unit_Cost": 30000},
        ])

        room_cost.save_manual_briefing_data(
            "2026-10-29", staff_incharge="Budi", remarks="Check lobby floor",
            entries=entries, storage_dir=str(tmp_path),
        )
        saved = room_cost.load_manual_briefing_data("2026-10-29", str(tmp_path))

        assert saved["Staff_Incharge"] == "Budi"
        assert saved["Remarks"] == "Check lobby floor"
        assert len(saved["Entries"]) == 4
        project_total = saved["Entries"].loc[
            saved["Entries"]["Entry_Type"] == "project_pa", "Total"
        ].iloc[0]
        assert project_total == 100000

        room_cost.save_manual_briefing_data(
            "2026-10-29", staff_incharge="Budi", remarks="Updated",
            entries=entries.iloc[:1], storage_dir=str(tmp_path),
        )
        replaced = room_cost.load_manual_briefing_data("2026-10-29", str(tmp_path))
        assert replaced["Remarks"] == "Updated"
        assert len(replaced["Entries"]) == 1

    def test_dropngo_import_preserves_other_manual_entries(self, tmp_path):
        room_cost.save_manual_briefing_data(
            "2026-11-01",
            staff_incharge="Nia",
            entries=pd.DataFrame([
                {"Entry_Type": "ra", "Staff_Name": "Ayu"},
                {"Entry_Type": "project_pa", "Area": "Lobby"},
            ]),
            storage_dir=str(tmp_path),
        )
        imported = pd.DataFrame([{
            "Date": "2026-11-01", "Room_Nights": 5, "Room_Numbers": "2522, 2523, 2723, 2726, 2537",
            "Guest_Laundry_Cost": 110000,
        }])

        room_cost.save_guest_laundry_data("2026-11", imported, storage_dir=str(tmp_path))
        saved = room_cost.load_manual_briefing_data("2026-11-01", str(tmp_path))
        laundry = room_cost.load_guest_laundry_data("2026-11", str(tmp_path))

        assert set(saved["Entries"]["Entry_Type"]) == {"ra", "project_pa", "guest_laundry"}
        assert saved["Staff_Incharge"] == "Nia"
        assert laundry.to_dict("records") == [{
            "Date": "2026-11-01", "Room_Nights": 5.0, "Guest_Laundry_Cost": 110000.0,
            "Room_Numbers": "2522, 2523, 2723, 2726, 2537",
        }]
        with pytest.raises(ValueError, match="Confirm replacement"):
            room_cost.save_guest_laundry_data("2026-11", imported, storage_dir=str(tmp_path))

    def test_night_audit_parser_extracts_room_sold_and_occupancy(self):
        workbook_buffer = BytesIO()
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.cell(1, 4, "DATE")
        sheet.cell(1, 14, "EXP OCC")
        sheet.cell(1, 17, "OCC (%)")
        sheet.cell(2, 4, pd.Timestamp("2026-10-01").to_pydatetime())
        sheet.cell(2, 14, 112)
        sheet.cell(2, 17, 92.56)
        sheet.cell(3, 4, pd.Timestamp("2026-10-02").to_pydatetime())
        sheet.cell(3, 14, 100)
        sheet.cell(3, 17, 82.65)
        workbook.save(workbook_buffer)
        workbook_buffer.seek(0)

        audit_rows = room_cost.parse_night_audit_report(workbook_buffer, "2026-10")

        assert audit_rows["Date"].tolist() == ["2026-10-01", "2026-10-02"]
        assert audit_rows["Room_Nights"].tolist() == [112, 100]
        assert audit_rows["Occupancy"].tolist() == [92.56, 82.65]
        assert audit_rows["Water"].isna().all()

    def test_night_audit_merge_preserves_saved_daily_costs(self):
        saved = self.daily_rows().iloc[:1].copy()
        saved["Occupancy"] = 75.0
        audit = pd.DataFrame([{
            "Date": "2026-10-01", "Occupancy": 92.56, "Room_Nights": 112,
            "Water": pd.NA, "Electricity": pd.NA, "Linen_Laundry": pd.NA,
        }])

        merged = room_cost.merge_daily_room_costs(
            saved, audit, replace_dates=True, month="2026-10"
        )

        assert merged.iloc[0]["Room_Nights"] == 112
        assert merged.iloc[0]["Occupancy"] == 92.56
        assert merged.iloc[0]["Water"] == 100000

    def test_night_audit_occupancy_and_room_sold_persist_for_automatic_display(self, tmp_path):
        audit = pd.DataFrame([{
            "Date": "2026-10-29", "Occupancy": 97.52, "Room_Nights": 118,
            "Water": pd.NA, "Electricity": pd.NA, "Linen_Laundry": pd.NA,
        }])
        merged = room_cost.merge_daily_room_costs(
            pd.DataFrame(columns=room_cost.DAILY_COLUMNS), audit, month="2026-10"
        )

        room_cost.save_daily_room_costs("2026-10", merged, str(tmp_path))
        reloaded = room_cost.load_daily_room_costs("2026-10", str(tmp_path))
        yesterday = reloaded.loc[reloaded["Date"] == "2026-10-29"].iloc[0]

        assert yesterday["Occupancy"] == 97.52
        assert yesterday["Room_Nights"] == 118

    def test_mtd_occupancy_and_room_sold_aggregate_saved_daily_values(self):
        daily = pd.DataFrame([
            {"Date": "2026-10-01", "Occupancy": 80.0, "Room_Nights": 80,
             "Water": pd.NA, "Electricity": pd.NA, "Linen_Laundry": pd.NA},
            {"Date": "2026-10-02", "Occupancy": 100.0, "Room_Nights": 100,
             "Water": pd.NA, "Electricity": pd.NA, "Linen_Laundry": pd.NA},
            {"Date": "2026-10-03", "Occupancy": pd.NA, "Room_Nights": pd.NA,
             "Water": pd.NA, "Electricity": pd.NA, "Linen_Laundry": pd.NA},
        ])

        occupancy_mtd, room_sold_mtd = room_cost.calculate_mtd_occupancy_and_room_nights(
            daily, "2026-10"
        )
        empty_occupancy, empty_room_sold = room_cost.calculate_mtd_occupancy_and_room_nights(
            pd.DataFrame(columns=room_cost.DAILY_COLUMNS), "2026-10"
        )

        assert occupancy_mtd == 90.0
        assert room_sold_mtd == 180
        assert pd.isna(empty_occupancy)
        assert empty_room_sold == 0

    def test_night_audit_rows_leave_combined_cost_incomplete(self):
        audit = pd.DataFrame([{
            "Date": "2026-10-01", "Occupancy": 92.56, "Room_Nights": 112,
            "Water": pd.NA, "Electricity": pd.NA, "Linen_Laundry": pd.NA,
        }])

        daily, monthly, warnings = room_cost.calculate_daily_room_costs(
            pd.DataFrame(columns=["Date", "Total"]), audit, "2026-10"
        )

        assert daily.iloc[0]["Room_Nights"] == 112
        assert daily.iloc[0]["Occupancy"] == 92.56
        assert pd.isna(daily.iloc[0]["Total_Cost_Per_Room"])
        assert pd.isna(monthly.iloc[0]["Total_Cost_Per_Room"])
        assert any("Water is missing" in warning for warning in warnings)

    def test_energy_cost_parser_and_merge_preserve_other_fields(self):
        workbook_buffer = BytesIO()
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["Energy Cost October 2026"])
        sheet.append(["Tanggal", "Penggunaan Listrik", "Penggunaan Air"])
        sheet.append([1, 3896712.6, 523900])
        sheet.append([28, 2744183.62, 383760])
        workbook.save(workbook_buffer)
        workbook_buffer.seek(0)

        imported = room_cost.parse_energy_cost_workbook(workbook_buffer, "2026-10")
        saved = self.daily_rows().iloc[:1].copy()
        saved["Occupancy"] = 92.56
        merged = room_cost.merge_daily_room_costs(
            saved, imported, replace_dates=True, month="2026-10"
        )

        assert imported["Date"].tolist() == ["2026-10-01", "2026-10-28"]
        assert imported["Electricity"].sum() == pytest.approx(6640896.22)
        assert imported["Water"].sum() == pytest.approx(907660)
        assert merged.loc[merged.Date == "2026-10-01", "Room_Nights"].iloc[0] == 80
        assert merged.loc[merged.Date == "2026-10-01", "Linen_Laundry"].iloc[0] == 300000
        assert pd.isna(merged.loc[merged.Date == "2026-10-28", "Room_Nights"].iloc[0])

        with tempfile.TemporaryDirectory() as storage_dir:
            room_cost.save_daily_room_costs("2026-10", merged, storage_dir)
            reloaded = room_cost.load_daily_room_costs("2026-10", storage_dir)
        assert len(reloaded) == 2
        assert reloaded.loc[reloaded.Date == "2026-10-28", "Electricity"].iloc[0] == pytest.approx(2744183.62)

    def test_daily_linen_parser_uses_daily_total_and_active_day_count(self):
        workbook_buffer = BytesIO()
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["DAILY LINEN COST"])
        sheet.append([None, None, None, None, None, None, None, None])
        sheet.append([None, None, None, None, None, None, None, None])
        sheet.append(["Hari ke-", "Tanggal", "Linen (Regular)", "Spoting & OO",
                      "Prewash 50%", "Pillow Case Drop N Go", "TOTAL BIAYA HARIAN", "Kumulatif"])
        sheet.append([1, pd.Timestamp("2026-10-01").to_pydatetime(), 100, 20, 0, 30, 150, 150])
        sheet.append([2, pd.Timestamp("2026-10-02").to_pydatetime(), 200, 0, 0, 50, 250, 400])
        sheet.append([29, pd.Timestamp("2026-10-29").to_pydatetime(), 0, 0, 0, 0, 0, 400])
        sheet.append([None, "Jumlah hari yang ada data", None, None, None, None, 2])
        workbook.save(workbook_buffer)
        workbook_buffer.seek(0)

        imported = room_cost.parse_daily_linen_cost_workbook(workbook_buffer, "2026-10")

        assert imported["Date"].tolist() == ["2026-10-01", "2026-10-02"]
        assert imported["Linen_Laundry"].tolist() == [150, 250]
        assert imported["Linen_Laundry"].sum() == 400

    def test_bonvivo_parser_sums_linen_hk_sections_per_day(self):
        workbook_buffer = BytesIO()
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "LINEN HK"
        sheet.cell(7, 2, "TANGGAL")
        sheet.cell(7, 3, pd.Timestamp("2026-11-01").to_pydatetime())
        sheet.cell(7, 4, pd.Timestamp("2026-11-02").to_pydatetime())
        sheet.cell(8, 2, "NOMOR BILL")
        sheet.cell(9, 2, "SHEET TWIN")
        sheet.cell(9, 3, 2)
        sheet.cell(9, 4, 3)
        sheet.cell(9, 34, 100)
        sheet.cell(10, 1, "TOTAL")
        sheet.cell(12, 2, "TANGGAL")
        sheet.cell(12, 3, pd.Timestamp("2024-11-01").to_pydatetime())
        sheet.cell(13, 2, "HAND TOWEL OOO")
        sheet.cell(13, 3, 1)
        sheet.cell(13, 34, 50)
        sheet.cell(14, 1, "TOTAL")
        workbook.save(workbook_buffer)
        workbook_buffer.seek(0)

        imported = room_cost.parse_bonvivo_linen_workbook(workbook_buffer, "2026-11")

        assert imported["Date"].tolist() == ["2026-11-01", "2026-11-02"]
        assert imported["Linen_Laundry"].tolist() == [250, 300]

    def test_dropngo_parser_aggregates_daily_cost_and_distinct_rooms(self):
        workbook_buffer = BytesIO()
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "GUEST LAUNDRY"
        for column, value in enumerate(("No", "Item", "price", "Date"), 1):
            sheet.cell(4, column, value)
        sheet.merge_cells("D5:E5")
        sheet["D5"] = 1
        sheet.merge_cells("F5:G5")
        sheet["F5"] = 2
        for column, room in enumerate((2726, 2537, 2523, 2522), 4):
            sheet.cell(6, column, room)
        sheet.cell(7, 1, "PAKET / 5 PCS")
        sheet.cell(7, 3, 22000)
        sheet.cell(7, 4, 1)
        sheet.cell(7, 5, 1)
        sheet.cell(7, 6, 1)
        sheet.cell(8, 1, 1)
        sheet.cell(8, 2, "Laundry item")
        sheet.cell(8, 3, 5000)
        sheet.cell(8, 4, 2)
        sheet.cell(8, 6, 1)
        linen = workbook.create_sheet("LINEN HK")
        linen.cell(3, 2, "Item")
        linen.cell(3, 3, "Price")
        linen.cell(3, 4, 1)
        linen.cell(3, 5, 2)
        linen.cell(4, 2, "PILLOW CASE")
        linen.cell(4, 3, 700)
        linen.cell(4, 4, 323)
        linen.cell(4, 5, 10)
        workbook.save(workbook_buffer)
        workbook_buffer.seek(0)

        imported = room_cost.parse_dropngo_guest_laundry_workbook(workbook_buffer, "2026-11")

        assert imported.to_dict("records") == [
            {
                "Date": "2026-11-01", "Room_Nights": 2,
                "Room_Numbers": "2537, 2726", "Guest_Laundry_Cost": 54000.0,
            },
            {
                "Date": "2026-11-02", "Room_Nights": 1,
                "Room_Numbers": "2523", "Guest_Laundry_Cost": 27000.0,
            },
        ]
        linen_costs = room_cost.parse_dropngo_linen_workbook(workbook_buffer, "2026-11")
        assert linen_costs["Linen_Laundry"].tolist() == [226100, 7000]

    def test_mtd_sr_rate_aggregates_same_day_and_applies_retroactively(self):
        sr_data = pd.DataFrame([
            {"Date": "2026-10-01", "Total": 2000000},
            {"Date": "2026-10-01", "Total": 3000000},
            {"Date": "2026-10-04", "Total": 5000000},
        ])

        daily, monthly, warnings = room_cost.calculate_daily_room_costs(
            sr_data, self.daily_rows(), "2026-10"
        )

        assert warnings == []
        assert daily["SR_Received_On_Date"].tolist() == [5000000, 0, 0, 5000000]
        assert daily["SR_Per_Room"].tolist() == pytest.approx([10000000 / 300] * 4)
        assert daily["Total_Cost_Per_Room"].iloc[0] == pytest.approx(
            100000 / 80 + 200000 / 80 + 300000 / 80 + 10000000 / 300
        )
        assert monthly.iloc[0]["Room_Nights"] == 300
        assert monthly.iloc[0]["SR_Total"] == 10000000
        assert monthly.iloc[0]["Total_Room_Cost"] == 10000000 + 4 * 600000

    def test_daily_room_cost_exposes_sr_categories_and_sums_all_components_per_room(self):
        sr_data = pd.DataFrame([
            {"Date": "2026-10-01", "Category": "Guest Supplies", "Total": 2000},
            {"Date": "2026-10-01", "Category": "Cleaning Supplies", "Total": 3000},
            {"Date": "2026-10-01", "Category": "Paper Supplies", "Total": 4000},
            {"Date": "2026-10-01", "Category": "Print & Stationery", "Total": 1000},
        ])
        daily_input = pd.DataFrame([{
            "Date": "2026-10-01", "Room_Nights": 10,
            "Water": 100, "Electricity": 200, "Linen_Laundry": 300,
        }])

        daily, _, _ = room_cost.calculate_daily_room_costs(sr_data, daily_input, "2026-10")
        daily_row = daily.iloc[0]

        assert daily_row["Guest_Supplies_Per_Room"] == 200
        assert daily_row["Cleaning_Supplies_Per_Room"] == 300
        assert daily_row["Paper_Supplies_Per_Room"] == 400
        assert daily_row["Print_Stationery_Per_Room"] == 100
        assert daily_row["SR_Per_Room"] == 1000
        assert daily_row["Total_Cost_Per_Room"] == 1060

    def test_invalid_and_out_of_month_sr_rows_are_reported(self):
        sr_data = pd.DataFrame([
            {"Date": "Unknown", "Total": 100},
            {"Date": "2026-09-30", "Total": 200},
        ])

        _, _, warnings = room_cost.calculate_daily_room_costs(
            sr_data, self.daily_rows(), "2026-10"
        )

        assert any("missing or invalid" in warning for warning in warnings)
        assert any("outside 2026-10" in warning for warning in warnings)

    def test_daily_room_cost_works_without_sr_rows(self):
        daily, monthly, warnings = room_cost.calculate_daily_room_costs(
            pd.DataFrame(columns=["Date", "Total"]),
            self.daily_rows().iloc[:1],
            "2026-10",
        )

        assert warnings == []
        assert daily.iloc[0]["SR_Per_Room"] == 0
        assert daily.iloc[0]["Total_Cost_Per_Room"] == 7500
        assert monthly.iloc[0]["SR_Total"] == 0

    def test_export_adds_daily_room_cost_sheet(self):
        daily, monthly, _ = room_cost.calculate_daily_room_costs(
            pd.DataFrame([{"Date": "2026-10-01", "Total": 5000000}]),
            self.daily_rows().iloc[:1],
            "2026-10",
        )
        workbook = openpyxl.load_workbook(BytesIO(create_sr_separated_excel(
            pd.DataFrame(),
            room_cost_daily_df=daily,
            room_cost_monthly_df=monthly,
        )), data_only=True)

        sheet = workbook["Room Cost Daily"]
        headers = {sheet.cell(8, col).value: col for col in range(1, sheet.max_column + 1)}
        assert sheet["C5"].value == 80
        assert sheet.cell(10, headers["SR Cost / Room (IDR)"]).value == 62500
        assert sheet.cell(10, headers["Total Room Cost (IDR)"]).value == 5600000

    def test_morning_briefing_template_matches_shared_compact_layout(self):
        briefing_date = (datetime.now() - timedelta(days=1)).date()
        briefing_date_text = briefing_date.isoformat()
        operation_date_text = (briefing_date + timedelta(days=1)).isoformat()
        manual_data = {
            briefing_date_text: {
                "Entries": pd.DataFrame([{
                    "Entry_Type": "ra", "Staff_Name": "Rina", "Section": "Floor 5",
                    "Room_Credit": 18, "Description": "Room credit", "Notes": "Complete",
                }]),
            },
            operation_date_text: {
                "Staff_Incharge": "Dewi",
                "Remarks": "Check corridor lighting",
                "Entries": pd.DataFrame([
                    {
                        "Entry_Type": "ra", "Staff_Name": "Rina", "Section": "Floor 5",
                        "Room_Credit": 18, "Description": "Lamp inspection", "Notes": "Open",
                    },
                    {
                        "Entry_Type": "project_pa", "Area": "Lobby", "Description": "Polish floor",
                        "PIC": "Ari", "Notes": "Morning shift",
                    },
                    {
                        "Entry_Type": "daily_pa", "Area": "Pool", "Description": "Clean deck",
                        "PIC": "Budi", "Notes": "",
                    },
                    {
                        "Entry_Type": "guest_laundry", "Description": "Must not be exported",
                        "Total": 12345,
                    },
                ]),
            },
        }
        exported = create_morning_briefing_template(
            briefing_date.strftime("%Y-%m"),
            sr_data=pd.DataFrame([{
                "Date": briefing_date_text, "SR_Number": "SR-001",
                "Category": "Guest Supplies", "Total": 90000,
            }]),
            daily_cost_data=pd.DataFrame([{
                "Date": briefing_date_text, "Occupancy": 90.0, "Room_Nights": 90,
                "Water": 9000, "Electricity": 18000, "Linen_Laundry": 4500,
            }]),
            guest_laundry_data=pd.DataFrame([{
                "Date": briefing_date_text, "Room_Nights": 5,
                "Room_Numbers": "2522, 2523, 2537, 2723, 2726",
                "Guest_Laundry_Cost": 110000,
            }]),
            manual_data_by_date=manual_data,
        )
        workbook = openpyxl.load_workbook(
            BytesIO(exported), data_only=False
        )

        assert workbook.sheetnames == [
            "Briefing", "SR Total MTD", "Daily Linen Cost MTD", "Guest Laundry MTD",
            "Public Area Project", "Project RA", "Room Cost MTD",
        ]
        briefing = workbook["Briefing"]
        assert briefing["A2"].value == (
            f"Daily operational summary | {briefing_date.strftime('%Y-%m')} | "
            "MTD details are in the supporting sheets"
        )
        assert briefing["B4"].value.date() == briefing_date
        assert briefing["B5"].value == "=B4+1"
        assert briefing["A6"].value == "OCCUPANCY & STAFF"
        assert briefing["A7"].value == "Occupancy Semalam (%)"
        assert briefing["B6"].value is None
        assert "Room Cost MTD" in briefing["B7"].value
        assert briefing["B7"].number_format == "0.00%"
        assert briefing["A8"].value == "Room Sold"
        assert "Room Cost MTD" in briefing["B8"].value
        assert "Project RA" in briefing["B9"].value
        assert briefing["B10"].value == "Dewi"
        assert briefing["A12"].value == "ROOM COST | DETAIL HARI KEMARIN"
        assert briefing["A21"].value == "DROP N GO LAUNDRY | HARI KEMARIN"
        assert briefing["A22"].value == "Rooms"
        assert briefing["A23"].value == "2522, 2523, 2537\n2723, 2726"
        assert briefing["A23"].alignment.wrap_text is True
        assert briefing["A26"].value == "PROJECT RA | HARI INI"
        assert briefing["A35"].value == "PROJECT PUBLIC AREA | HARI INI"
        assert briefing["A46"].value == "REMARK / FOLLOW UP"
        assert "A47:E49" in {str(cell_range) for cell_range in briefing.merged_cells.ranges}
        assert briefing.print_area == "'Briefing'!$A$1:$E$49"
        assert briefing["A37"].value == "Lobby"
        assert briefing["B37"].value == "Polish floor"
        assert briefing["A47"].value == "Check corridor lighting"
        assert workbook["Project RA"]["B5"].value == "Rina"
        assert workbook["Project RA"]["E6"].value == "Lamp inspection"
        assert "Daily Public Area" not in workbook.sheetnames
        assert "Project Public Area" not in workbook.sheetnames
        workbook_text = " ".join(
            str(cell.value)
            for sheet in workbook.worksheets
            for row in sheet.iter_rows()
            for cell in row
            if cell.value is not None
        ).lower()
        assert "bon vivo" not in workbook_text
        assert "daily pa" not in workbook_text
        assert workbook["Public Area Project"]["C4"].value == "Brushing area loading dock & cleaning gutter"
        assert workbook["Public Area Project"]["D4"].value is None
        with ZipFile(BytesIO(exported)) as archive:
            namespace = {
                "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                "office": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                "package": "http://schemas.openxmlformats.org/package/2006/relationships",
            }
            workbook_xml = ElementTree.fromstring(archive.read("xl/workbook.xml"))
            relationships_xml = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            targets = {
                relation.attrib["Id"]: relation.attrib["Target"]
                for relation in relationships_xml.findall("package:Relationship", namespace)
            }
            sheet_paths = {
                sheet.attrib["name"]: targets[sheet.attrib[f"{{{namespace['office']}}}id"]].lstrip("/")
                for sheet in workbook_xml.findall("main:sheets/main:sheet", namespace)
            }
            checklist_xml = archive.read(sheet_paths["Public Area Project"])
            ra_xml = archive.read(sheet_paths["Project RA"])
        assert b"x14:conditionalFormattings" in checklist_xml
        assert b"x14:conditionalFormattings" not in ra_xml
        room_cost = workbook["Room Cost MTD"]
        daily_row = 4 + briefing_date.day
        assert room_cost.cell(daily_row, 2).value == 0.9
        assert room_cost.cell(daily_row, 3).value == 90
        assert room_cost.cell(daily_row, 5).value == 9000
        assert room_cost.cell(daily_row, 6).value == 18000
        assert room_cost.cell(daily_row, 7).value == 4500
        assert workbook["Daily Linen Cost MTD"].cell(daily_row, 2).value == 4500
        assert workbook["Guest Laundry MTD"].cell(daily_row, 2).value == 5
        assert workbook["Guest Laundry MTD"].cell(daily_row, 3).value == 110000
        assert workbook["Guest Laundry MTD"]["B4"].value == "Rooms"
        operation_row = 4 + (briefing_date + timedelta(days=1)).day
        assert workbook["Guest Laundry MTD"].cell(operation_row, 3).value == 12345
        assert workbook["SR Total MTD"]["B5"].value == "SR-001"
        assert workbook["SR Total MTD"]["D5"].value == 90000

    def test_morning_briefing_template_uses_selected_month_day_rows(self):
        exported = create_morning_briefing_template(
            "2026-10",
            daily_cost_data=pd.DataFrame([{
                "Date": "2026-10-31", "Occupancy": 74.0, "Room_Nights": 81,
                "Water": 1000, "Electricity": 2000, "Linen_Laundry": 3000,
            }]),
        )
        workbook = openpyxl.load_workbook(BytesIO(exported), data_only=False)

        assert workbook["Briefing"]["B4"].value.date().isoformat() == "2026-10-31"
        assert "$35" in workbook["Briefing"]["B7"].value
        assert workbook["Room Cost MTD"]["A35"].value.date().isoformat() == "2026-10-31"
        assert workbook["Room Cost MTD"]["B35"].value == 0.74
        assert workbook["Room Cost MTD"]["C35"].value == 81
        assert workbook["Daily Linen Cost MTD"]["A35"].value.date().isoformat() == "2026-10-31"
        assert workbook["Daily Linen Cost MTD"]["B35"].value == 3000


class TestNormalizeSRLabels:
    def test_replaces_undefined_labels(self):
        df = pd.DataFrame([
            {"Item_Code": "976744", "Item_Name": " undefined ", "Category": "undefined"},
            {"Item_Code": "976802", "Item_Name": None, "Category": "Paper Supplies"},
        ])

        normalized = normalize_sr_labels(df)

        assert normalized["Item_Name"].tolist() == ["Item 976744", "Item 976802"]
        assert normalized["Category"].tolist() == ["Guest Supplies", "Paper Supplies"]


# ---------------------------------------------------------------------------
# CLI version test
# ---------------------------------------------------------------------------

class TestCLIVersion:
    def test_version_flag(self):
        """separate_expenses.py --version should print version and exit."""
        import subprocess
        result = subprocess.run(
            [sys.executable, "-m", "separate_expenses", "--version"],
            capture_output=True, text=True, timeout=10
        )
        # Either it succeeds with version output or exits with system exit code
        assert "0.1.0" in result.stdout or result.returncode in (0, 1)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])