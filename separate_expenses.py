#!/usr/bin/env python3
"""
separate_expenses.py - Command-line interface for Hotel Santika Depok Expense Separator.
Usage:
    python separate_expenses.py --input "/path/to/8.AGUSTUS" --output "Separated_Expenses_Agustus_2026.xlsx"
    python separate_expenses.py --all
"""

import os
import sys
import argparse
import parser
import matcher
import exporter


def process_month(input_dir: str, output_path: str = None, month_label: str = None):
    if not os.path.isdir(input_dir):
        print(f"Error: Directory '{input_dir}' not found.", file=sys.stderr)
        return False

    base_name = os.path.basename(os.path.abspath(input_dir))
    if not month_label:
        # Infer a canonical 'Month YYYY' label (e.g. '8.AGUSTUS' -> 'August 2026')
        month_label = parser.infer_month_label(base_name, default=base_name)
    else:
        # Normalize explicitly supplied labels too (e.g. 'Agustus 2026' -> 'August 2026')
        month_label = parser.infer_month_label(month_label, default=month_label)

    if not output_path:
        output_path = f"Separated_Expenses_{month_label.replace(' ', '_')}.xlsx"

    print(f"\n=======================================================")
    print(f"Processing Month: {month_label}")
    print(f"Input Directory : {input_dir}")
    print(f"Output File     : {output_path}")
    print(f"=======================================================")

    files = parser.find_month_files(input_dir)
    print(f"• Detail Trial Balance : {files['dtb'] or 'NOT FOUND'}")
    print(f"• Income Statement MTD : {files['is_mtd'] or 'NOT FOUND'}")
    print(f"• Consumption Report   : {files['consumption'] or 'NOT FOUND'}")

    if not files['dtb']:
        print("Error: Could not locate Detail Trial Balance file.", file=sys.stderr)
        return False

    # Parse files
    print("\n[1/3] Ingesting files...")
    dtb_df = parser.parse_detail_trial_balance(files['dtb'])
    print(f"      Loaded {len(dtb_df)} transactions from Detail Trial Balance.")

    is_df = parser.parse_income_statement(files['is_mtd']) if files['is_mtd'] else None
    if is_df is not None:
        print(f"      Loaded {len(is_df)} budget & actual lines from Income Statement.")

    cons_df = parser.parse_consumption_report(files['consumption']) if files['consumption'] else None
    if cons_df is not None:
        print(f"      Loaded {len(cons_df)} item issuing lines from Consumption Report.")

    # Reconcile & Enrich
    print("\n[2/3] Reconciling and matching transactions...")
    reconciled = matcher.reconcile_monthly_expenses(
        dtb_df=dtb_df,
        is_df=is_df if is_df is not None else dtb_df.iloc[:0],
        cons_df=cons_df if cons_df is not None else dtb_df.iloc[:0]
    )
    
    m = reconciled['metrics']
    print(f"      Reconciled {m['transaction_count']} transactions across {m['category_count']} categories.")
    print(f"      Total Spent : Rp {m['total_spent']:,.0f}")
    print(f"      Total Budget: Rp {m['total_budget']:,.0f}")
    print(f"      Net Variance: Rp {m['variance_idr']:,.0f} ({m['variance_pct']:.1f}%)")

    # Export
    print("\n[3/3] Generating separated Excel workbook...")
    exporter.create_separated_excel(reconciled, month_label, output_path)
    print(f"SUCCESS: Saved multi-tab separated workbook to:\n   {os.path.abspath(output_path)}\n")
    return True


def get_default_review_dir() -> str:
    env_path = os.environ.get("HOTEL_DATA_DIR")
    if env_path:
        return env_path

    home_dir = os.path.expanduser("~")
    candidates = [
        os.path.join(home_dir, "Documents", "Business Review"),
        os.path.join(home_dir, "Business Review"),
        os.path.join("C:\\", "Users", os.path.basename(home_dir), "Documents", "Business Review"),
        os.path.join("C:\\", "Business Review"),
    ]

    for candidate in candidates:
        if os.path.exists(candidate):
            return candidate

    return os.path.join(home_dir, "Documents", "Business Review")


VERSION = "0.1.0"

def main():
    arg_parser = argparse.ArgumentParser(description="Hotel Santika Depok - Expense Separator CLI")
    arg_parser.add_argument("--input", "-i", type=str, help="Path to month folder (e.g. C:/Business Review/8.AGUSTUS)")
    arg_parser.add_argument("--output", "-o", type=str, help="Output Excel filename (.xlsx)")
    arg_parser.add_argument("--month", "-m", type=str, help="Month label (e.g. 'Agustus 2026')")
    arg_parser.add_argument("--all", action="store_true", help="Process all available months in the local Business Review folder")
    arg_parser.add_argument("--version", "-v", action="version", version=f"%(prog)s {VERSION}")

    args = arg_parser.parse_args()

    if args.all:
        base_review = get_default_review_dir()
        months = ["8.AGUSTUS", "7.JULY", "6. JUN"]
        for m in months:
            m_path = os.path.join(base_review, m)
            if os.path.exists(m_path):
                process_month(m_path)
    elif args.input:
        process_month(args.input, args.output, args.month)
    else:
        default_dir = os.path.join(get_default_review_dir(), "8.AGUSTUS")
        if os.path.exists(default_dir):
            process_month(default_dir)
        else:
            arg_parser.print_help()


if __name__ == "__main__":
    main()
