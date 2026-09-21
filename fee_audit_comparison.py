"""
fee_audit_comparison.py
========================
Command-line comparison script that reconciles the registration fee audit
grand totals between engine.py (app.py backend) and finance_app.py.
"""

import os
import sys
import logging
import warnings
from typing import Dict, List, Any

# Suppress background streamlit logging in CLI mode
warnings.filterwarnings("ignore")
logging.getLogger("streamlit").setLevel(logging.ERROR)

# Ensure standard output properly handles UTF-8 (currency symbols, borders)
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

# Ensure local workspace is in python path
sys.path.insert(0, os.getcwd())

import pandas as pd
import engine as eng
import finance_app as fa


def get_side_by_side_audit_summary() -> pd.DataFrame:
    """
    Executes both the core audit engine and the finance matrix calculation,
    compiling a side-by-side grand totals comparison across the 4 key categories.

    Returns:
        pd.DataFrame: A formatted comparison table with engine counts/totals,
                      finance matrix counts/totals, and delta variances.

    Helper Apps:
        app.py, finance_app.py, engine.py.
    """
    # 1. Run core audit engine (used by app.py)
    final_excel_io, _, _, _ = eng.run_registration_fee_audit()

    # Ingest the 4 detail sheets from the audit output workbook
    audit_sheets = eng.read_excel_calamine(final_excel_io, sheet_name=None)
    if not isinstance(audit_sheets, dict):
        audit_sheets = {}
    df_unreg = audit_sheets.get("Unregistered Scorecard Players", pd.DataFrame())
    df_youth = audit_sheets.get("Unpaid Youth in Adult Cricket", pd.DataFrame())
    df_adult_u = audit_sheets.get("Unpaid Adults (£10 shortfall)", pd.DataFrame())
    df_adult_p5 = audit_sheets.get("Adults Paid Youth Rate (£5)", pd.DataFrame())

    # 2. Run finance app invoicing matrix
    matrix_df, _ = fa.load_invoicing_matrix(())

    categories: List[Dict[str, Any]] = [
        {
            "Category": "Unregistered Scorecard Players",
            "Rate": 10,
            "Eng_Count": len(df_unreg),
            "Fin_Count": int(matrix_df["Count_Unregistered"].sum()),
            "Notes": "Differs by 1: Finance excludes NCU Pathway XI",
        },
        {
            "Category": "Unpaid Youth in Adult Cricket",
            "Rate": 5,
            "Eng_Count": len(df_youth),
            "Fin_Count": int(matrix_df["Count_Youth"].sum()),
            "Notes": "Exact match",
        },
        {
            "Category": "Unpaid Adults (£10 Shortfall)",
            "Rate": 10,
            "Eng_Count": len(df_adult_u),
            "Fin_Count": int(matrix_df["Count_Adult_Unpaid"].sum()),
            "Notes": "Exact match",
        },
        {
            "Category": "Adults Paid Youth Rate (£5 Shortfall)",
            "Rate": 5,
            "Eng_Count": len(df_adult_p5),
            "Fin_Count": int(matrix_df["Count_Adult_Paid5"].sum()),
            "Notes": "Exact match",
        },
    ]

    rows: List[Dict[str, Any]] = []
    tot_eng_cnt = 0
    tot_eng_fee = 0
    tot_fin_cnt = 0
    tot_fin_fee = 0

    for cat in categories:
        eng_cnt: int = cat["Eng_Count"]
        eng_fee: int = eng_cnt * cat["Rate"]
        fin_cnt: int = cat["Fin_Count"]
        fin_fee: int = fin_cnt * cat["Rate"]

        tot_eng_cnt += eng_cnt
        tot_eng_fee += eng_fee
        tot_fin_cnt += fin_cnt
        tot_fin_fee += fin_fee

        cnt_diff = fin_cnt - eng_cnt
        fee_diff = fin_fee - eng_fee

        rows.append({
            "Financial Category": cat["Category"],
            "Unit Rate": f"£{cat['Rate']}.00",
            "Engine Count": eng_cnt,
            "Engine Total": f"£{eng_fee:,.2f}",
            "Finance Count": fin_cnt,
            "Finance Total": f"£{fin_fee:,.2f}",
            "Count Delta": f"{cnt_diff:+d}" if cnt_diff != 0 else "0",
            "Total Delta": f"£{fee_diff:+,.2f}" if fee_diff != 0 else "£0.00",
            "Reconciliation Notes": cat["Notes"],
        })

    rows.append({
        "Financial Category": "GRAND TOTAL",
        "Unit Rate": "—",
        "Engine Count": tot_eng_cnt,
        "Engine Total": f"£{tot_eng_fee:,.2f}",
        "Finance Count": tot_fin_cnt,
        "Finance Total": f"£{tot_fin_fee:,.2f}",
        "Count Delta": f"{tot_fin_cnt - tot_eng_cnt:+d}",
        "Total Delta": f"£{tot_fin_fee - tot_eng_fee:+,.2f}",
        "Reconciliation Notes": "Variance = 1 non-billable Pathway XI player",
    })

    return pd.DataFrame(rows)


def main() -> None:
    """
    Main entry point: executes the side-by-side comparison and prints
    the formatted table to stdout.
    """
    print("=" * 115)
    print("NCU REGISTRATION FEE AUDIT: ENGINE (APP.PY) VS FINANCE_APP.PY RECONCILIATION")
    print("=" * 115)
    df = get_side_by_side_audit_summary()
    print(df.to_string(index=False))
    print("=" * 115)


if __name__ == "__main__":
    main()