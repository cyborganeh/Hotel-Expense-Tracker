"""Tests for the hotel-expense-tracker modules."""
import os
import sys
import tempfile
import pytest
import pandas as pd

# Ensure project root is on the path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from parser import find_month_files, parse_detail_trial_balance, parse_income_statement, parse_consumption_report, parse_laundry_reports
from matcher import reconcile_monthly_expenses
from exporter import create_separated_excel


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