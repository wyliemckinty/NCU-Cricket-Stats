"""
NCU Cricket Hub - Finance & Invoicing Command Center
===================================================
A standalone financial auditing and invoicing management tool for the Northern Cricket Union (NCU).

Calculates, isolates, and exports club-by-club invoicing schedules across the 4 mandatory infraction categories:
1. Unregistered Scorecard Players (£10 fine per player)
2. Unpaid Youth in Adult Cricket (£5 fee shortfall per player)
3. Unpaid Adults (£10 Shortfall) (£10 fee shortfall per player)
4. Adults Paid Youth Rate (£5 Shortfall) (£5 catch-up charge per player)

Features:
- Live KPI summary metrics.
- Interactive club invoicing matrix with sorting and search.
- Per-club drill-down player audit inspector.
- Excel ledger downloader generating 'NCU_Season_2026_Invoicing_Schedule.xlsx' with OpenPyXL,
  #1F4E78 Navy Blue styling, freeze panes, auto-fit columns, and explicit Excel =SUM() audit formulas.
- Individual Word invoice document exporter generating '[Club_Name]_Season_2026_Invoice.docx'
  into 'club_invoices/' with executive #1F4E78 branding, itemized category table, and audit trail footnote.
"""

import io
import os
import re
import zipfile
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

try:
    import docx
    from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Inches, Pt, RGBColor
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

import engine as eng
import starring_rules as sr

# Configure Streamlit page
st.set_page_config(
    page_title="NCU Finance & Invoicing Command Center",
    page_icon="💳",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Dark Navy Branding CSS
st.markdown("""
<style>
    /* Metric styling */
    [data-testid="stMetric"] {
        background-color: #F8FAFC !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 8px !important;
        padding: 12px 16px !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04) !important;
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.88rem !important;
        font-weight: 600 !important;
        color: #475569 !important;
    }
    [data-testid="stMetricValue"] {
        font-size: 1.5rem !important;
        color: #1F4E78 !important;
        font-weight: 700 !important;
        white-space: normal !important;
        word-break: break-word !important;
        text-overflow: unset !important;
        overflow: visible !important;
        line-height: 1.3 !important;
    }
    /* Buttons */
    .stButton button {
        border-radius: 6px !important;
        font-weight: 600 !important;
    }
    /* Status Badges & Tag Protections: Prevents text clipping and truncation */
    span[data-testid="stBadge"],
    div[data-testid="stBadge"],
    .status-badge {
        white-space: nowrap !important;
        overflow: visible !important;
        text-overflow: unset !important;
        font-weight: 600 !important;
        display: inline-flex !important;
        align-items: center !important;
    }

    /* Center alignment for DOM/HTML table headers */
    th, th[role="columnheader"] {
        text-align: center !important;
    }
    th.col-left, th[data-col-align="left"], td.col-left, td[data-col-align="left"] {
        text-align: left !important;
    }
</style>
""", unsafe_allow_html=True)

def center_col_label(label: str, target_width: int = 14) -> str:
    """
    Pads a column header label symmetrically so that the text visually
    aligns to the center of the column and mirrors centered data cells.
    """
    if len(label) >= target_width:
        return f" {label} "
    return label.center(target_width)

def get_standard_column_config() -> Dict[str, Any]:
    """
    Constructs a centralized, standardized column configuration dictionary
    for st.dataframe across the finance application.
    Centered columns have center-balanced header labels mirroring centered cells,
    while left-aligned columns remain strictly left-aligned.
    """
    return {
        "Club": st.column_config.TextColumn("Club Name", width="medium", alignment="left"),
        "Club Name": st.column_config.TextColumn("Club Name", width="medium", alignment="left"),
        "Player": st.column_config.TextColumn("Player", alignment="left", width="medium"),
        "Full_Name": st.column_config.TextColumn("Player Name", width="medium", alignment="left"),
        "From Club": st.column_config.TextColumn("From Club", alignment="left", width="medium"),
        "Unregistered Scorecard Players": st.column_config.NumberColumn(
            center_col_label("Unreg (£10)", 16),
            format="£%d",
            width="small",
            alignment="center",
            help="£10 per player: played in matches but completely lacks a Sport80 registration entry"
        ),
        "Unpaid Youth in Adult Cricket": st.column_config.NumberColumn(
            center_col_label("Unpaid Youth (£5)", 20),
            format="£%d",
            width="small",
            alignment="center",
            help="£5 per player: junior players (<18) who played senior adult matches with £0 paid"
        ),
        "Unpaid Adults (£10 Shortfall)": st.column_config.NumberColumn(
            center_col_label("Unpaid Adults (£10)", 20),
            format="£%d",
            width="small",
            alignment="center",
            help="£10 per player: adult players (>=18) who played senior adult matches with £0 paid"
        ),
        "Adults Paid Youth Rate (£5 Shortfall)": st.column_config.NumberColumn(
            center_col_label("Youth Rate Diff (£5)", 22),
            format="£%d",
            width="small",
            alignment="center",
            help="£5 per player: adult players (>=18) who mistakenly paid the reduced £5 youth rate"
        ),
        "Transfer Fees (£25)": st.column_config.NumberColumn(
            center_col_label("Transfers (£25)", 18),
            format="£%d",
            width="small",
            alignment="center",
            help="£25 per player: seasonal transfer registered on or after 1st April (NCU Rule A13)"
        ),
        "Total Invoice Due": st.column_config.NumberColumn(
            center_col_label("Total Due", 14),
            format="£%d",
            width="small",
            alignment="center",
            help="Total fee owed across all infraction categories"
        ),
        "Transfer Number": st.column_config.Column(center_col_label("Transfer #", 14), alignment="center", width="small"),
        "Transfer Date": st.column_config.Column(center_col_label("Transfer Date", 16), alignment="center", width="medium"),
        "Fee Due (£)": st.column_config.NumberColumn(center_col_label("Fee Due (£)", 14), format="£%.2f", alignment="center", width="small"),
        "Fee Infraction": st.column_config.CheckboxColumn("£25 Late Fee Infraction (>= 1 Apr)", width="medium"),
        "Total_Paid": st.column_config.NumberColumn(center_col_label("Fee Cleared", 14), format="£%.2f", alignment="center", width="small"),
        "Compliance Status": st.column_config.TextColumn(center_col_label("Compliance Status", 20), width="medium", alignment="center"),
    }


def clean_club_name(raw_club: Any) -> str:
    """
    Normalizes a club or team string into a clean, canonical club name.

    Inputs:
        raw_club: Raw club or team string from scorecard or registration records.

    Outputs:
        str: Canonical club base name (e.g. 'Instonians', 'Holywood 1881', 'CSNI').

    Helper Apps:
        finance_app.py, engine.py.
    """
    if not raw_club or pd.isna(raw_club):
        return "Unknown"
    c_str = str(raw_club).strip()
    if c_str == "" or c_str.lower() == "nan":
        return "Unknown"
    # Defensively strip parenthesized annotations (e.g. '(Overseas Professional)')
    c_str = re.sub(r'\(.*?\)', '', c_str).strip()
    base = eng.extract_base_club_name(c_str)
    if base.lower() in ["drumaness", "drumaness superkings", "drumaness super kings"]:
        return "Drumaness Superkings"
    return base


def sanitize_club_filename(club_name: str) -> str:
    """
    Sanitizes a club name for safe use in filenames while preserving readable formatting.

    Inputs:
        club_name: Raw club name string.

    Outputs:
        str: Filesystem-safe club name string.

    Helper Apps:
        finance_app.py.
    """
    safe_name = re.sub(r'[\\/*?:"<>|]', "", str(club_name)).strip()
    return safe_name if safe_name else "Club"


def set_cell_background(cell: Any, hex_color: str) -> None:
    """
    Sets the background fill color of a table cell in python-docx using w:shd.

    Inputs:
        cell: python-docx table cell object.
        hex_color: Hexadecimal color string (e.g., '1F4E78', 'D9E1F2', 'F2F2F2').

    Outputs:
        None.

    Helper Apps:
        finance_app.py.
    """
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith("shd"):
            tcPr.remove(child)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def set_cell_margins(cell: Any, top: int = 120, bottom: int = 120, left: int = 160, right: int = 160) -> None:
    """
    Sets inner cell padding (margins) for a table cell in dxa (1 pt = 20 dxa).

    Inputs:
        cell: python-docx table cell object.
        top: Top padding in dxa.
        bottom: Bottom padding in dxa.
        left: Left padding in dxa.
        right: Right padding in dxa.

    Outputs:
        None.

    Helper Apps:
        finance_app.py.
    """
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


@st.cache_data(show_spinner="Running registration fee audit & calculating club invoices...")
def load_invoicing_matrix(file_signatures: tuple) -> Tuple[pd.DataFrame, Dict[str, pd.DataFrame]]:
    """
    Ingests registration, revenue, and scorecard records, executing the full audit
    and aggregating club invoices across the 4 mandatory infraction categories.

    Inputs:
        file_signatures: Tuple of (filepath, modification_time) pairs for cache invalidation.

    Outputs:
        Tuple[pd.DataFrame, Dict[str, pd.DataFrame]]:
            - matrix_df: Club-by-club financial invoice summary.
            - detail_dfs: Dictionary of player-level detail DataFrames for each category.

    Helper Apps:
        finance_app.py.
    """
    try:
        final_excel_io, _, _, df_master = eng.run_registration_fee_audit()
    except Exception:
        return pd.DataFrame(), {}

    # Ingest the detail sheets from the audit engine output
    audit_sheets = eng.read_excel_calamine(final_excel_io, sheet_name=None)
    if not isinstance(audit_sheets, dict):
        audit_sheets = {}

    df_unreg = audit_sheets.get("Unregistered Scorecard Players", pd.DataFrame())
    df_youth = audit_sheets.get("Unpaid Youth in Adult Cricket", pd.DataFrame())
    df_adult_unpaid = audit_sheets.get("Unpaid Adults (£10 shortfall)", pd.DataFrame())
    df_adult_paid5 = audit_sheets.get("Adults Paid Youth Rate (£5)", pd.DataFrame())

    # Ingest compliant player sheets
    df_comp_adults = audit_sheets.get("Compliant Adults (£10+)", pd.DataFrame()).copy()
    if not df_comp_adults.empty:
        df_comp_adults["Player Category"] = "Adult"
        df_comp_adults["Compliance Status"] = "Verified & Cleared"

    df_comp_youth = audit_sheets.get("Compliant Youths (£5)", pd.DataFrame()).copy()
    if not df_comp_youth.empty:
        df_comp_youth["Player Category"] = "Youth"
        df_comp_youth["Compliance Status"] = "Verified & Cleared"

    # Combine compliant profiles
    df_compliant = pd.concat([df_comp_adults, df_comp_youth], ignore_index=True)

    if df_master is not None and not df_master.empty:
        try:
            extra_youth = df_master[
                df_master["Is_Youth"] & df_master["Played_Adult_Matches"] & (df_master["Total_Paid"] > 5)
            ].copy()
            if not extra_youth.empty:
                extra_youth = extra_youth.rename(columns={
                    "Date Registered Formatted": "Date Registered",
                    "Types_Paid": "Payment Details"
                })
                extra_youth["Player Category"] = "Youth (Adult Rate)"
                extra_youth["Compliance Status"] = "Verified & Cleared"
                if "Full_Name" in df_compliant.columns and not df_compliant.empty:
                    existing_names = set(df_compliant["Full_Name"].dropna())
                    extra_youth = extra_youth[~extra_youth["Full_Name"].isin(existing_names)]
                if not extra_youth.empty:
                    common_cols = [c for c in df_compliant.columns if c in extra_youth.columns]
                    if common_cols:
                        extra_youth = extra_youth[common_cols]
                    df_compliant = pd.concat([df_compliant, extra_youth], ignore_index=True)
        except Exception:
            pass

    if not df_compliant.empty and "Individual Membership Primary Club" in df_compliant.columns:
        df_compliant["Club_Clean"] = df_compliant["Individual Membership Primary Club"].apply(clean_club_name)
    elif not df_compliant.empty and "Club" in df_compliant.columns:
        df_compliant["Club_Clean"] = df_compliant["Club"].apply(clean_club_name)
    else:
        df_compliant["Club_Clean"] = pd.Series(dtype=str)

    # Standardize club names in detail DataFrames
    df_unreg["Club_Clean"] = df_unreg["Inferred Club"].apply(clean_club_name)
    df_youth["Club_Clean"] = df_youth["Individual Membership Primary Club"].apply(clean_club_name)
    df_adult_unpaid["Club_Clean"] = df_adult_unpaid["Individual Membership Primary Club"].apply(clean_club_name)
    df_adult_paid5["Club_Clean"] = df_adult_paid5["Individual Membership Primary Club"].apply(clean_club_name)

    # 1. Unregistered Scorecard Players (£10 fine per player)
    unreg_counts = df_unreg[df_unreg["Club_Clean"] != "Unknown"]["Club_Clean"].value_counts()

    # 2. Unpaid Youth in Adult Cricket (£5 shortfall per player)
    youth_counts = df_youth[df_youth["Club_Clean"] != "Unknown"]["Club_Clean"].value_counts()

    # 3. Unpaid Adults (£10 Shortfall) (£10 shortfall per player)
    adult_u_counts = df_adult_unpaid[df_adult_unpaid["Club_Clean"] != "Unknown"]["Club_Clean"].value_counts()

    # 4. Adults Paid Youth Rate (£5 Shortfall) (£5 shortfall per player)
    adult_p5_counts = df_adult_paid5[df_adult_paid5["Club_Clean"] != "Unknown"]["Club_Clean"].value_counts()

    # 5. Rule A13 Seasonal Player Transfers (£25 fee on or after 1st April)
    reg_file = eng.DEFAULT_FILES.get("Men's", {}).get("reg", "1. NCU_Registered_Players.xlsx")
    df_reg = eng.get_excel_df(reg_file) if reg_file and os.path.exists(reg_file) else pd.DataFrame()
    transfer_results = sr.evaluate_player_transfers(df_reg, season_year=2026, clean_club_fn=clean_club_name)
    transfer_club_summary = transfer_results.get("club_summary", {})
    df_transfers = transfer_results.get("valid_transfers", pd.DataFrame())
    df_blocked_transfers = transfer_results.get("blocked_transfers", pd.DataFrame())

    # Union of all official NCU clubs plus any clubs with at least one record
    all_clubs = sorted(
        set(eng.NCU_ALL_CLUBS)
        .union(unreg_counts.index)
        .union(youth_counts.index)
        .union(adult_u_counts.index)
        .union(adult_p5_counts.index)
        .union(transfer_club_summary.keys())
    )

    rows = []
    for club in all_clubs:
        if club in ["Unknown", "NCU Pathway XI"]:
            continue

        c_unreg = int(unreg_counts.get(club, 0))
        c_youth = int(youth_counts.get(club, 0))
        c_adult_u = int(adult_u_counts.get(club, 0))
        c_adult_p5 = int(adult_p5_counts.get(club, 0))

        t_summary = transfer_club_summary.get(club, {})
        c_transfers = int(t_summary.get("fee_transfers_count", 0))
        fee_transfers = float(t_summary.get("total_transfer_fees", 0.0))

        fee_unreg = c_unreg * 10
        fee_youth = c_youth * 5
        fee_adult_u = c_adult_u * 10
        fee_adult_p5 = c_adult_p5 * 5
        total_invoice = fee_unreg + fee_youth + fee_adult_u + fee_adult_p5 + fee_transfers

        rows.append({
            "Club": club,
            "Unregistered Scorecard Players": fee_unreg,
            "Unpaid Youth in Adult Cricket": fee_youth,
            "Unpaid Adults (£10 Shortfall)": fee_adult_u,
            "Adults Paid Youth Rate (£5 Shortfall)": fee_adult_p5,
            "Transfer Fees (£25)": fee_transfers,
            "Total Invoice Due": total_invoice,
            "Count_Unregistered": c_unreg,
            "Count_Youth": c_youth,
            "Count_Adult_Unpaid": c_adult_u,
            "Count_Adult_Paid5": c_adult_p5,
            "Count_Transfers": c_transfers,
        })

    matrix_df = pd.DataFrame(rows)
    if not matrix_df.empty:
        matrix_df = matrix_df.sort_values(by="Club", ascending=True).reset_index(drop=True)

    detail_dfs = {
        "unreg": df_unreg,
        "youth": df_youth,
        "adult_unpaid": df_adult_unpaid,
        "adult_paid5": df_adult_paid5,
        "compliant": df_compliant,
        "transfers": df_transfers,
        "blocked_transfers": df_blocked_transfers,
    }

    return matrix_df, detail_dfs


def generate_invoicing_workbook(
    matrix_df: pd.DataFrame,
    detail_dfs: Dict[str, pd.DataFrame]
) -> io.BytesIO:
    """
    Generates an OpenPyXL Excel workbook containing the Invoicing Schedule and detailed audit tabs.
    Strictly follows mandatory #1F4E78 Navy Blue styling, freeze panes, auto-fit columns,
    and explicit Excel =SUM() formulas on the bottom row for Sharon's audit trail.

    Inputs:
        matrix_df: Club-by-club financial invoice summary.
        detail_dfs: Dictionary of player-level detail DataFrames for each category.

    Outputs:
        io.BytesIO: In-memory Excel workbook (.xlsx).

    Helper Apps:
        finance_app.py.
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Styling Tokens per Gemini.md
    navy_header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_bold_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    total_font = Font(name="Calibri", size=11, bold=True, color="000000")
    currency_fmt = "£#,##0.00"

    border_thin = Side(style="thin", color="D3D3D3")
    border_double = Side(style="double", color="000000")
    border_total = Border(top=border_thin, bottom=border_double)

    # 1. Main Sheet: Invoicing Schedule
    ws_main = wb.create_sheet(title="Invoicing Schedule")
    ws_main.freeze_panes = "A2"

    has_transfer_col = "Transfer Fees (£25)" in matrix_df.columns

    main_headers = [
        "Club Name",
        "Unregistered Scorecard Players (£10)",
        "Unpaid Youth in Adult Cricket (£5)",
        "Unpaid Adults (£10 Shortfall)",
        "Adults Paid Youth Rate (£5 Shortfall)",
    ]
    if has_transfer_col:
        main_headers.append("Transfer Fees (£25)")
    main_headers.append("Total Invoice Due")

    ws_main.append(main_headers)
    num_cols = len(main_headers)

    # Style Header Row
    for col_idx in range(1, num_cols + 1):
        cell = ws_main.cell(row=1, column=col_idx)
        cell.fill = navy_header_fill
        cell.font = white_bold_font
        cell.alignment = Alignment(horizontal="center" if col_idx > 1 else "left", vertical="center")

    # Add Data Rows
    num_rows = len(matrix_df)
    for _, r in matrix_df.iterrows():
        row_vals = [
            r["Club"],
            r["Unregistered Scorecard Players"],
            r["Unpaid Youth in Adult Cricket"],
            r["Unpaid Adults (£10 Shortfall)"],
            r["Adults Paid Youth Rate (£5 Shortfall)"],
        ]
        if has_transfer_col:
            row_vals.append(r["Transfer Fees (£25)"])
        row_vals.append(r["Total Invoice Due"])
        ws_main.append(row_vals)

    # Apply currency format to data cells
    for row_idx in range(2, num_rows + 2):
        for col_idx in range(2, num_cols + 1):
            cell = ws_main.cell(row=row_idx, column=col_idx)
            cell.number_format = currency_fmt
            cell.alignment = Alignment(horizontal="right", vertical="center")

    # Insert Explicit Excel =SUM() Formulas at Bottom Row
    total_row_idx = num_rows + 2
    ws_main.cell(row=total_row_idx, column=1, value="TOTAL")
    ws_main.cell(row=total_row_idx, column=1).font = total_font
    ws_main.cell(row=total_row_idx, column=1).fill = total_fill
    ws_main.cell(row=total_row_idx, column=1).border = border_total

    for col_idx in range(2, num_cols + 1):
        col_letter = get_column_letter(col_idx)
        cell = ws_main.cell(
            row=total_row_idx,
            column=col_idx,
            value=f"=SUM({col_letter}2:{col_letter}{total_row_idx - 1})"
        )
        cell.font = total_font
        cell.fill = total_fill
        cell.border = border_total
        cell.number_format = currency_fmt
        cell.alignment = Alignment(horizontal="right", vertical="center")

    # Vectorized column width calculation
    for col_idx in range(1, num_cols + 1):
        col_letter = get_column_letter(col_idx)
        max_len = max(
            len(str(ws_main.cell(row=r, column=col_idx).value or ""))
            for r in range(1, total_row_idx + 1)
        )
        ws_main.column_dimensions[col_letter].width = max(max_len + 3, 14)

    # 2. Detail Sheets for Sharon's Audit Trail
    detail_configs = [
        ("Unregistered Players", detail_dfs.get("unreg", pd.DataFrame()), [
            "Match_Player_Display", "Inferred Club", "Registration Status", "Payment Status", "Total_Matches", "Competitions", "Matches Played"
        ]),
        ("Unpaid Youth", detail_dfs.get("youth", pd.DataFrame()), [
            "Full_Name", "Date of Birth", "Date Registered", "Age at Registration", "Individual Membership Primary Club", "Total_Matches", "Payment Details", "Matches Played"
        ]),
        ("Unpaid Adults (£10)", detail_dfs.get("adult_unpaid", pd.DataFrame()), [
            "Full_Name", "Date of Birth", "Date Registered", "Age at Registration", "Individual Membership Primary Club", "Total_Matches", "Payment Details", "Matches Played"
        ]),
        ("Adults Paid Youth Rate (£5)", detail_dfs.get("adult_paid5", pd.DataFrame()), [
            "Full_Name", "Date of Birth", "Date Registered", "Age at Registration", "Individual Membership Primary Club", "Total_Matches", "Payment Details", "Matches Played"
        ]),
        ("Rule A13 Transfers", detail_dfs.get("transfers", pd.DataFrame()), [
            "Player", "From Club", "To Club", "Transfer Number", "Transfer Date", "Fee Due (£)", "Fee Infraction"
        ]),
    ]

    for sheet_title, df_d, cols in detail_configs:
        if df_d.empty:
            continue
        ws_d = wb.create_sheet(title=sheet_title)
        ws_d.freeze_panes = "A2"

        avail_cols = [c for c in cols if c in df_d.columns]
        ws_d.append(avail_cols)

        # Style header
        for col_idx in range(1, len(avail_cols) + 1):
            cell = ws_d.cell(row=1, column=col_idx)
            cell.fill = navy_header_fill
            cell.font = white_bold_font
            cell.alignment = Alignment(horizontal="left", vertical="center")

        # Append data
        for _, drow in df_d[avail_cols].iterrows():
            row_vals = []
            for val in drow.values:
                if pd.isna(val) or str(val).lower() == "nan":
                    row_vals.append("")
                else:
                    row_vals.append(str(val))
            ws_d.append(row_vals)

        # Auto-fit columns
        for col_idx in range(1, len(avail_cols) + 1):
            col_letter = get_column_letter(col_idx)
            max_len = max(
                len(str(ws_d.cell(row=r, column=col_idx).value or ""))
                for r in range(1, min(ws_d.max_row + 1, 200))
            )
            ws_d.column_dimensions[col_letter].width = max(min(max_len + 2, 45), 12)

    out_bio = io.BytesIO()
    wb.save(out_bio)
    out_bio.seek(0)
    return out_bio


def format_match_appearances(
    raw_matches: Any,
    total_matches: Optional[Any] = None
) -> List[Tuple[str, bool]]:
    """
    Parses and trims a player's match appearance history to a maximum of 2 preview entries,
    followed by an italicized summary count footer if additional appearances exist.

    Inputs:
        raw_matches: Raw match appearance string from the audit record.
        total_matches: Optional total match count integer or string.

    Outputs:
        List[Tuple[str, bool]]: List of (text_line, is_italic) tuples for Word document cell rendering.

    Helper Apps:
        finance_app.py.
    """
    if not raw_matches or pd.isna(raw_matches):
        return [("-", False)]

    raw_str = str(raw_matches).strip()
    if not raw_str or raw_str.lower() in ["nan", "none", "-"]:
        return [("-", False)]

    # Split match entries cleanly
    if "\n" in raw_str:
        entries = [p.strip() for p in raw_str.split("\n") if p.strip()]
    elif ";" in raw_str:
        entries = [p.strip() for p in raw_str.split(";") if p.strip()]
    else:
        entries = [p.strip() for p in re.split(r"(?<=\d{4}),\s*", raw_str) if p.strip()]
        if len(entries) == 1 and ", " in raw_str and raw_str.count(" v ") > 1:
            entries = [p.strip() for p in re.split(r",\s*(?=[^,]+ v )", raw_str) if p.strip()]

    if not entries:
        return [("-", False)]

    # Determine total appearance count
    parsed_tot: Optional[int] = None
    if total_matches is not None and not pd.isna(total_matches):
        try:
            parsed_tot = int(total_matches)
        except (ValueError, TypeError):
            parsed_tot = None

    tot_count = max(parsed_tot or len(entries), len(entries))

    # Slice max 2 match previews
    preview_entries = entries[:2]
    formatted_lines: List[Tuple[str, bool]] = []

    for entry in preview_entries:
        bullet_text = f"• {entry}" if tot_count > 1 else entry
        formatted_lines.append((bullet_text, False))

    # Append summary count footer if player participated in more than 2 matches
    if tot_count > 2:
        remaining_count = tot_count - len(preview_entries)
        plural_str = "appearance" if remaining_count == 1 else "appearances"
        formatted_lines.append((f"+ {remaining_count} other match {plural_str}", True))

    return formatted_lines


def generate_club_word_invoice(
    club_name: str,
    club_row: pd.Series,
    detail_dfs: Dict[str, pd.DataFrame],
    output_path: str
) -> str:
    """
    Generates an individual invoice statement document (.docx) for a specific club,
    featuring an executive #1F4E78 navy header, club metadata summary block,
    itemized 4-category fee breakdown table, player audit appendix, and verification footnote.

    Inputs:
        club_name: Name of the cricket club.
        club_row: Series containing fee totals and counts for the club.
        detail_dfs: Dictionary of detail DataFrames containing player breakdown.
        output_path: Destination file path for the .docx file.

    Outputs:
        str: Absolute path to the saved document.

    Helper Apps:
        finance_app.py.
    """
    doc = docx.Document()

    # Page Margins: 0.75 in
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    # 1. Executive Header Block (#1F4E78 Navy Theme)
    header_table = doc.add_table(rows=1, cols=1)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_table.autofit = False
    header_cell = header_table.cell(0, 0)
    header_cell.width = Inches(7.0)
    set_cell_background(header_cell, "1F4E78")
    set_cell_margins(header_cell, top=180, bottom=180, left=240, right=240)

    p_org = header_cell.paragraphs[0]
    p_org.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_org = p_org.add_run("NORTHERN CRICKET UNION")
    r_org.font.name = "Calibri"
    r_org.font.size = Pt(11)
    r_org.font.bold = True
    r_org.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    p_title = header_cell.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(3)
    p_title.paragraph_format.space_after = Pt(3)
    r_title = p_title.add_run("NCU Season Invoice Statement - 2026")
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(20)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    p_sub = header_cell.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Official Registration Non-Compliance & Audit Schedule")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(10)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(0xD9, 0xE1, 0xF2)

    doc.add_paragraph()  # Spacing

    # 2. Club Summary Block
    summary_table = doc.add_table(rows=1, cols=2)
    summary_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    summary_table.autofit = False

    cell_meta = summary_table.cell(0, 0)
    cell_meta.width = Inches(4.3)
    set_cell_background(cell_meta, "F8FAFC")
    set_cell_margins(cell_meta, top=140, bottom=140, left=180, right=180)

    club_slug = re.sub(r"[^A-Za-z0-9]", "", club_name).upper()[:8]
    total_due = float(club_row.get("Total Invoice Due", 0))

    p_c1 = cell_meta.paragraphs[0]
    r_c1_lbl = p_c1.add_run("ISSUED TO:\n")
    r_c1_lbl.font.size = Pt(9)
    r_c1_lbl.font.bold = True
    r_c1_lbl.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)
    r_c1_val = p_c1.add_run(f"{club_name} Cricket Club\n")
    r_c1_val.font.size = Pt(14)
    r_c1_val.font.bold = True
    r_c1_val.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)

    r_c1_meta = p_c1.add_run(
        f"Invoice Ref: NCU-2026-INV-{club_slug}\n"
        f"Billing Season: 2026 Season\n"
        f"Date of Issue: {datetime.now().strftime('%d %B %Y')}"
    )
    r_c1_meta.font.size = Pt(9.5)
    r_c1_meta.font.color.rgb = RGBColor(0x33, 0x41, 0x55)

    cell_bal = summary_table.cell(0, 1)
    cell_bal.width = Inches(2.7)
    set_cell_background(cell_bal, "EBF1F5")
    set_cell_margins(cell_bal, top=140, bottom=140, left=180, right=180)

    p_bal = cell_bal.paragraphs[0]
    p_bal.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_bal_lbl = p_bal.add_run("TOTAL OUTSTANDING BALANCE\n")
    r_bal_lbl.font.size = Pt(8.5)
    r_bal_lbl.font.bold = True
    r_bal_lbl.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)

    r_bal_amt = p_bal.add_run(f"£{total_due:,.2f}\n")
    r_bal_amt.font.size = Pt(22)
    r_bal_amt.font.bold = True
    r_bal_amt.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)

    if total_due > 0:
        r_bal_status = p_bal.add_run("STATUS: PAYMENT DUE\nTerms: Remit within 30 days")
        r_bal_status.font.size = Pt(8.5)
        r_bal_status.font.bold = True
        r_bal_status.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    else:
        r_bal_status = p_bal.add_run("STATUS: NIL BALANCE\nAccount Fully Compliant")
        r_bal_status.font.size = Pt(8.5)
        r_bal_status.font.bold = True
        r_bal_status.font.color.rgb = RGBColor(0x16, 0x65, 0x34)

    doc.add_paragraph()  # Spacing

    # 3. Itemized Fee Breakdown Table (4 Mandatory Categories)
    p_table_head = doc.add_paragraph()
    r_th = p_table_head.add_run("Itemized Infraction Fee Schedule")
    r_th.font.name = "Calibri"
    r_th.font.size = Pt(13)
    r_th.font.bold = True
    r_th.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)
    p_table_head.paragraph_format.space_after = Pt(4)

    items_table = doc.add_table(rows=1, cols=4)
    items_table.style = "Table Grid"
    items_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    items_table.autofit = False

    col_widths = [Inches(3.4), Inches(1.2), Inches(1.1), Inches(1.3)]
    headers = [
        "Mandatory Financial Category",
        "Unit Rate",
        "Flagged Count",
        "Subtotal (£)"
    ]

    for col_idx, (h_text, w) in enumerate(zip(headers, col_widths)):
        c = items_table.cell(0, col_idx)
        c.width = w
        set_cell_background(c, "1F4E78")
        set_cell_margins(c, top=100, bottom=100, left=140, right=140)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if col_idx == 0 else WD_ALIGN_PARAGRAPH.RIGHT
        r = p.add_run(h_text)
        r.font.name = "Calibri"
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # Categories data
    c_unreg = int(club_row.get("Count_Unregistered", 0))
    fee_unreg = float(club_row.get("Unregistered Scorecard Players", 0))

    c_youth = int(club_row.get("Count_Youth", 0))
    fee_youth = float(club_row.get("Unpaid Youth in Adult Cricket", 0))

    c_adult_u = int(club_row.get("Count_Adult_Unpaid", 0))
    fee_adult_u = float(club_row.get("Unpaid Adults (£10 Shortfall)", 0))

    c_adult_p5 = int(club_row.get("Count_Adult_Paid5", 0))
    fee_adult_p5 = float(club_row.get("Adults Paid Youth Rate (£5 Shortfall)", 0))

    c_transfers = int(club_row.get("Count_Transfers", 0))
    fee_transfers = float(club_row.get("Transfer Fees (£25)", 0))

    categories_data = [
        ("Unregistered Scorecard Players", "£10.00 / player", c_unreg, fee_unreg),
        ("Unpaid Youth in Adult Cricket", "£5.00 / player", c_youth, fee_youth),
        ("Unpaid Adults (£10 Shortfall)", "£10.00 / player", c_adult_u, fee_adult_u),
        ("Adults Paid Youth Rate (£5 Shortfall)", "£5.00 / player", c_adult_p5, fee_adult_p5),
    ]
    if c_transfers > 0 or fee_transfers > 0:
        categories_data.append(("Rule A13 Transfer Fees", "£25.00 / transfer", c_transfers, fee_transfers))

    active_categories = [cat for cat in categories_data if cat[2] > 0]

    if active_categories:
        for row_num, (cat_title, unit_rate, count, subtotal) in enumerate(active_categories, start=1):
            row = items_table.add_row()
            bg_color = "FFFFFF" if row_num % 2 == 1 else "F9FAFB"

            row_vals = [
                (cat_title, WD_ALIGN_PARAGRAPH.LEFT, False),
                (unit_rate, WD_ALIGN_PARAGRAPH.RIGHT, False),
                (str(count), WD_ALIGN_PARAGRAPH.RIGHT, False),
                (f"£{subtotal:,.2f}", WD_ALIGN_PARAGRAPH.RIGHT, False),
            ]
            for col_idx, (val_text, align, is_bold) in enumerate(row_vals):
                c = row.cells[col_idx]
                c.width = col_widths[col_idx]
                set_cell_background(c, bg_color)
                set_cell_margins(c, top=80, bottom=80, left=140, right=140)
                p = c.paragraphs[0]
                p.alignment = align
                r = p.add_run(val_text)
                r.font.name = "Calibri"
                r.font.size = Pt(9.5)
                r.font.bold = is_bold
                r.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)
    else:
        # Zero Balance confirmation row
        row = items_table.add_row()
        row_vals = [
            ("No registration infractions or fee shortfalls recorded (Fully Compliant)", WD_ALIGN_PARAGRAPH.LEFT, False),
            ("-", WD_ALIGN_PARAGRAPH.RIGHT, False),
            ("0", WD_ALIGN_PARAGRAPH.RIGHT, False),
            ("£0.00", WD_ALIGN_PARAGRAPH.RIGHT, False),
        ]
        for col_idx, (val_text, align, is_bold) in enumerate(row_vals):
            c = row.cells[col_idx]
            c.width = col_widths[col_idx]
            set_cell_background(c, "F8FAFC")
            set_cell_margins(c, top=80, bottom=80, left=140, right=140)
            p = c.paragraphs[0]
            p.alignment = align
            r = p.add_run(val_text)
            r.font.name = "Calibri"
            r.font.size = Pt(9.5)
            r.font.italic = (col_idx == 0)
            r.font.color.rgb = RGBColor(0x16, 0x65, 0x34) if col_idx == 0 else RGBColor(0x64, 0x74, 0x8B)

    # Total Row
    total_row = items_table.add_row()
    tot_count = c_unreg + c_youth + c_adult_u + c_adult_p5
    total_vals = [
        ("TOTAL INVOICE DUE", WD_ALIGN_PARAGRAPH.LEFT),
        ("", WD_ALIGN_PARAGRAPH.RIGHT),
        (str(tot_count), WD_ALIGN_PARAGRAPH.RIGHT),
        (f"£{total_due:,.2f}", WD_ALIGN_PARAGRAPH.RIGHT),
    ]
    for col_idx, (t_text, align) in enumerate(total_vals):
        c = total_row.cells[col_idx]
        c.width = col_widths[col_idx]
        set_cell_background(c, "D9E1F2")
        set_cell_margins(c, top=100, bottom=100, left=140, right=140)
        p = c.paragraphs[0]
        p.alignment = align
        r = p.add_run(t_text)
        r.font.name = "Calibri"
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)

    doc.add_paragraph()  # Spacing

    # 4. Detailed Player Audit Appendix
    club_unreg = detail_dfs.get("unreg", pd.DataFrame())
    club_unreg = club_unreg[club_unreg["Club_Clean"] == club_name] if "Club_Clean" in club_unreg.columns else pd.DataFrame()

    club_youth = detail_dfs.get("youth", pd.DataFrame())
    club_youth = club_youth[club_youth["Club_Clean"] == club_name] if "Club_Clean" in club_youth.columns else pd.DataFrame()

    club_adult_u = detail_dfs.get("adult_unpaid", pd.DataFrame())
    club_adult_u = club_adult_u[club_adult_u["Club_Clean"] == club_name] if "Club_Clean" in club_adult_u.columns else pd.DataFrame()

    club_adult_p5 = detail_dfs.get("adult_paid5", pd.DataFrame())
    club_adult_p5 = club_adult_p5[club_adult_p5["Club_Clean"] == club_name] if "Club_Clean" in club_adult_p5.columns else pd.DataFrame()

    has_details = not (club_unreg.empty and club_youth.empty and club_adult_u.empty and club_adult_p5.empty)

    if has_details:
        p_app_head = doc.add_paragraph()
        r_app = p_app_head.add_run("Audit Trail: Flagged Player Appearances")
        r_app.font.name = "Calibri"
        r_app.font.size = Pt(12)
        r_app.font.bold = True
        r_app.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)
        p_app_head.paragraph_format.space_after = Pt(4)

        detail_sections = [
            ("Unregistered Scorecard Players (£10 fine)", club_unreg, "Match_Player_Display", "Matches Played"),
            ("Unpaid Youth in Adult Cricket (£5 shortfall)", club_youth, "Full_Name", "Matches Played"),
            ("Unpaid Adults (£10 shortfall)", club_adult_u, "Full_Name", "Matches Played"),
            ("Adults Paid Youth Rate (£5 shortfall)", club_adult_p5, "Full_Name", "Matches Played"),
        ]

        for sec_title, df_sub, name_col, match_col in detail_sections:
            if df_sub.empty:
                continue

            p_sec = doc.add_paragraph()
            r_sec = p_sec.add_run(f"• {sec_title} — {len(df_sub)} Player(s):")
            r_sec.font.name = "Calibri"
            r_sec.font.size = Pt(10)
            r_sec.font.bold = True
            r_sec.font.color.rgb = RGBColor(0x33, 0x41, 0x55)
            p_sec.paragraph_format.space_before = Pt(4)
            p_sec.paragraph_format.space_after = Pt(2)

            sub_table = doc.add_table(rows=1, cols=3)
            sub_table.style = "Table Grid"
            sub_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            sub_widths = [Inches(0.6), Inches(2.6), Inches(3.8)]
            sub_headers = ["#", "Player Name", "Match Appearances / Notes"]

            for col_idx, (sh_text, sw) in enumerate(zip(sub_headers, sub_widths)):
                sc = sub_table.cell(0, col_idx)
                sc.width = sw
                set_cell_background(sc, "E2E8F0")
                set_cell_margins(sc, top=50, bottom=50, left=80, right=80)
                sp = sc.paragraphs[0]
                sr = sp.add_run(sh_text)
                sr.font.size = Pt(8.5)
                sr.font.bold = True
                sr.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)

            for p_idx, (_, prow) in enumerate(df_sub.iterrows(), start=1):
                p_row = sub_table.add_row()
                p_name = str(prow.get(name_col, "Unknown"))
                p_matches = str(prow.get(match_col, prow.get("Total_Matches", "-")))

                # Col 0: Index
                c0 = p_row.cells[0]
                c0.width = sub_widths[0]
                set_cell_margins(c0, top=40, bottom=40, left=80, right=80)
                p0 = c0.paragraphs[0]
                r0 = p0.add_run(str(p_idx))
                r0.font.name = "Calibri"
                r0.font.size = Pt(8.5)
                r0.font.color.rgb = RGBColor(0x33, 0x41, 0x55)

                # Col 1: Player Name
                c1 = p_row.cells[1]
                c1.width = sub_widths[1]
                set_cell_margins(c1, top=40, bottom=40, left=80, right=80)
                p1 = c1.paragraphs[0]
                r1 = p1.add_run(p_name)
                r1.font.name = "Calibri"
                r1.font.size = Pt(8.5)
                r1.font.color.rgb = RGBColor(0x33, 0x41, 0x55)

                # Col 2: Match Appearances / Notes (Trimmed to max 2 previews + italic summary footer)
                c2 = p_row.cells[2]
                c2.width = sub_widths[2]
                set_cell_margins(c2, top=40, bottom=40, left=80, right=80)
                tot_m = prow.get("Total_Matches")
                lines = format_match_appearances(p_matches, total_matches=tot_m)
                for l_idx, (line_text, is_italic) in enumerate(lines):
                    p2 = c2.paragraphs[0] if l_idx == 0 else c2.add_paragraph()
                    p2.paragraph_format.space_before = Pt(0)
                    p2.paragraph_format.space_after = Pt(1)
                    r2 = p2.add_run(line_text)
                    r2.font.name = "Calibri"
                    r2.font.size = Pt(8.5)
                    r2.font.italic = is_italic
                    if is_italic:
                        r2.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)
                    else:
                        r2.font.color.rgb = RGBColor(0x33, 0x41, 0x55)

            doc.add_paragraph()  # Spacing

    # 5. Remittance & Bank Instructions
    p_remit = doc.add_paragraph()
    p_remit.paragraph_format.space_before = Pt(6)
    p_remit.paragraph_format.space_after = Pt(2)
    if total_due > 0:
        r_remit_head = p_remit.add_run("Payment & Remittance Instructions:\n")
        r_remit_head.font.name = "Calibri"
        r_remit_head.font.size = Pt(9.5)
        r_remit_head.font.bold = True
        r_remit_head.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)

        r_remit_txt = p_remit.add_run(
            f"Please remit total balance due of £{total_due:,.2f} quoting invoice reference 'NCU-2026-INV-{club_slug}' "
            "via electronic bank transfer to the Northern Cricket Union accounts. For inquiries or payment receipts, "
            "email: finance@northerncricketunion.org."
        )
        r_remit_txt.font.name = "Calibri"
        r_remit_txt.font.size = Pt(8.5)
        r_remit_txt.font.color.rgb = RGBColor(0x47, 0x55, 0x69)
    else:
        r_remit_head = p_remit.add_run("Account Status & Confirmation:\n")
        r_remit_head.font.name = "Calibri"
        r_remit_head.font.size = Pt(9.5)
        r_remit_head.font.bold = True
        r_remit_head.font.color.rgb = RGBColor(0x16, 0x65, 0x34)

        r_remit_txt = p_remit.add_run(
            f"No payment is required. This statement confirms that {club_name} Cricket Club has zero outstanding registration "
            "fees or scorecard compliance shortfalls for the 2026 season. For inquiries, email: finance@northerncricketunion.org."
        )
        r_remit_txt.font.name = "Calibri"
        r_remit_txt.font.size = Pt(8.5)
        r_remit_txt.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    # 6. Official Standard Footnote
    p_foot = doc.add_paragraph()
    p_foot.paragraph_format.space_before = Pt(12)
    r_foot = p_foot.add_run(
        "Audit Trail & Verification Notice: All player registration, fee tier, and match appearance records are "
        "reconciled directly against official NV Play electronic scorecards and Sport80 membership registers. "
        "Clubs wishing to inspect the complete ball-by-ball scorecards, player rosters, or query any itemized charge "
        "can access the full audit ledger via the NCU Cricket Hub or contact the NCU Finance Directorate."
    )
    r_foot.font.name = "Calibri"
    r_foot.font.size = Pt(8)
    r_foot.font.italic = True
    r_foot.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)

    # Ensure parent directory exists
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc.save(output_path)
    return os.path.abspath(output_path)


def generate_all_club_word_invoices(
    matrix_df: pd.DataFrame,
    detail_dfs: Dict[str, pd.DataFrame],
    output_dir: str = "club_invoices",
    only_invoiced: bool = True
) -> List[str]:
    """
    Loops through the club matrix and automatically generates individual Word invoices (.docx)
    for each club into the designated output folder.

    Inputs:
        matrix_df: Master DataFrame of club invoicing totals and counts.
        detail_dfs: Dictionary of detail DataFrames containing player breakdown.
        output_dir: Local destination directory (defaults to 'club_invoices').
        only_invoiced: If True, only generates invoices for clubs with balance > £0.

    Outputs:
        List[str]: List of filepaths of all generated invoice documents.

    Helper Apps:
        finance_app.py.
    """
    os.makedirs(output_dir, exist_ok=True)
    generated_paths = []

    target_df = matrix_df[matrix_df["Total Invoice Due"] > 0] if only_invoiced else matrix_df

    for _, row in target_df.iterrows():
        club = str(row["Club"]).strip()
        if not club or club in ["Unknown", "TOTAL"]:
            continue
        safe_club = sanitize_club_filename(club)
        filename = f"{safe_club}_Season_2026_Invoice.docx"
        out_path = os.path.join(output_dir, filename)
        saved_path = generate_club_word_invoice(club, row, detail_dfs, out_path)
        generated_paths.append(saved_path)

    return generated_paths


def create_invoices_zip(file_paths: List[str]) -> io.BytesIO:
    """
    Compresses a list of generated document files into an in-memory ZIP archive.

    Inputs:
        file_paths: List of filepaths to compress.

    Outputs:
        io.BytesIO: In-memory zip archive buffer.

    Helper Apps:
        finance_app.py.
    """
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for fpath in file_paths:
            if os.path.exists(fpath):
                zf.write(fpath, arcname=os.path.basename(fpath))
    zip_buffer.seek(0)
    return zip_buffer


def get_name_sort_key(row: Any) -> Tuple[str, str]:
    """
    Computes a sorting key tuple (surname, christian_name) for a player record.
    Delegates to engine.get_name_sort_key().

    Inputs:
        row: Series, dictionary, or string containing player details.

    Outputs:
        Tuple[str, str]: (surname.lower(), christian_name.lower()) tuple for stable sorting.

    Helper Apps:
        finance_app.py.
    """
    return eng.get_name_sort_key(row)


def extract_compliant_rosters(
    detail_dfs: Dict[str, pd.DataFrame]
) -> Dict[str, pd.DataFrame]:
    """
    Extracts and groups verified compliant player profiles on a club-by-club basis.
    When a player has paid fees across multiple clubs (e.g. in-season transfers, dual registrations,
    or youth-to-adult transitions across clubs), splits their profile cleanly across the respective clubs,
    attributing each club's specific fee amount and payment detail line.
    Sorts all player records within each club alphabetically by surname, then Christian name.

    Inputs:
        detail_dfs: Dictionary of audit detail DataFrames containing 'compliant'.

    Outputs:
        Dict[str, pd.DataFrame]: Mapping of canonical club name to DataFrame of compliant players.

    Helper Apps:
        finance_app.py.
    """
    df_comp = detail_dfs.get("compliant", pd.DataFrame())
    if df_comp.empty:
        return {}

    # Standardize club grouping key if not already populated
    if "Club_Clean" not in df_comp.columns:
        if "Individual Membership Primary Club" in df_comp.columns:
            df_comp["Club_Clean"] = df_comp["Individual Membership Primary Club"].apply(clean_club_name)
        elif "Club" in df_comp.columns:
            df_comp["Club_Clean"] = df_comp["Club"].apply(clean_club_name)
        else:
            df_comp["Club_Clean"] = "Unknown"

    club_records: Dict[str, List[Dict[str, Any]]] = {}

    for _, row in df_comp.iterrows():
        details = str(row.get("Payment Details", "") or row.get("Types_Paid", "") or "").strip()
        primary_club = str(row.get("Club_Clean", "")).strip() or clean_club_name(row.get("Individual Membership Primary Club", ""))

        # Check if details contain multiple payment entries
        if ";" in details:
            parts = [p.strip() for p in details.split(";") if p.strip()]
            club_breakdown: Dict[str, Dict[str, Any]] = {}

            for p in parts:
                if " - " in p:
                    c_raw = p.split(" - ", 1)[0].strip()
                    c_clean = clean_club_name(c_raw)
                else:
                    c_clean = primary_club if primary_club and primary_club != "Unknown" else "Unknown"

                # Extract fee amount from payment line (e.g. '(£10)' or '(£5)')
                amt_match = re.search(r"£(\d+(?:\.\d+)?)", p)
                if amt_match:
                    item_amt = float(amt_match.group(1))
                else:
                    try:
                        item_amt = float(row.get("Total_Paid", 10.0)) / max(len(parts), 1)
                    except (ValueError, TypeError):
                        item_amt = 10.0

                if c_clean not in club_breakdown:
                    club_breakdown[c_clean] = {"amt": 0.0, "lines": []}
                club_breakdown[c_clean]["amt"] += item_amt
                club_breakdown[c_clean]["lines"].append(p)

            # If valid clubs were parsed
            valid_clubs = {c: v for c, v in club_breakdown.items() if c not in ["Unknown", "TOTAL"]}
            if valid_clubs:
                for c_name, c_data in valid_clubs.items():
                    r_dict = row.to_dict()
                    r_dict["Club_Clean"] = c_name
                    r_dict["Individual Membership Primary Club"] = c_name
                    r_dict["Total_Paid"] = c_data["amt"]
                    r_dict["Payment Details"] = "; ".join(c_data["lines"])
                    if c_name not in club_records:
                        club_records[c_name] = []
                    club_records[c_name].append(r_dict)
                continue

        # Single club / fallback handling
        target_club = primary_club if primary_club and primary_club not in ["Unknown", "TOTAL"] else "Unknown"
        if target_club != "Unknown":
            if target_club not in club_records:
                club_records[target_club] = []
            club_records[target_club].append(row.to_dict())

    grouped: Dict[str, pd.DataFrame] = {}
    for club_str, rows in club_records.items():
        if not rows or club_str in ["Unknown", "TOTAL"]:
            continue
        c_df = pd.DataFrame(rows)
        # Sort by surname, then Christian name
        sort_keys = c_df.apply(get_name_sort_key, axis=1)
        grouped[club_str] = c_df.iloc[sort_keys.argsort()].reset_index(drop=True)

    return grouped


def generate_club_compliant_roster_excel(
    club_name: str,
    compliant_df: pd.DataFrame
) -> io.BytesIO:
    """
    Builds an OpenPyXL Excel workbook containing the verified compliant player roster
    for a specific club ('[Club_Name]_Season_2026_Compliant_Roster.xlsx').
    Strictly applies #1F4E78 Navy Blue styling to header lines, freezes row 2,
    enables cell sorting (auto_filter), and uses vectorization to auto-fit columns.

    Inputs:
        club_name: Name of the cricket club.
        compliant_df: DataFrame containing the club's compliant player records.

    Outputs:
        io.BytesIO: In-memory Excel workbook (.xlsx).

    Helper Apps:
        finance_app.py.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Compliant Roster"

    # Mandatory styling per Gemini.md
    navy_header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_bold_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    currency_fmt = "£#,##0.00"

    # Freeze panes at row 2
    ws.freeze_panes = "A2"

    headers = [
        "Player Name",
        "Category",
        "Date of Birth",
        "Age at Registration",
        "Date Registered",
        "Fee Cleared",
        "Payment Details",
        "Matches Played",
        "Compliance Status"
    ]
    ws.append(headers)

    # Style Header Row
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = navy_header_fill
        cell.font = white_bold_font
        cell.alignment = Alignment(horizontal="center" if col_idx > 1 else "left", vertical="center")

    # Add Data Rows sorted by surname then Christian name
    if not compliant_df.empty:
        sort_keys = compliant_df.apply(get_name_sort_key, axis=1)
        sorted_df = compliant_df.iloc[sort_keys.argsort()].reset_index(drop=True)
        for _, r in sorted_df.iterrows():
            name = str(r.get("Full_Name", "") or "").strip()
            category = str(r.get("Player Category", "") or ("Youth" if r.get("Is_Youth", False) else "Adult")).strip()
            dob = str(r.get("Date of Birth", "") or "").strip()
            age = r.get("Age at Registration", "")
            age_val = "" if pd.isna(age) else str(age)
            date_reg = str(r.get("Date Registered", "") or "").strip()
            fee = r.get("Total_Paid", 0)
            try:
                fee_val = float(fee)
            except (ValueError, TypeError):
                fee_val = 0.0
            pay_details = str(r.get("Payment Details", "") or "").strip()
            matches = str(r.get("Matches Played", "") or r.get("Total_Matches", "") or "").strip()
            status = str(r.get("Compliance Status", "") or "Verified & Cleared").strip()

            ws.append([
                name,
                category,
                dob,
                age_val,
                date_reg,
                fee_val,
                pay_details,
                matches,
                status
            ])

        # Apply formatting to data rows
        num_rows = len(compliant_df)
        for row_idx in range(2, num_rows + 2):
            # Fee Cleared column (col 6)
            fee_cell = ws.cell(row=row_idx, column=6)
            fee_cell.number_format = currency_fmt
            fee_cell.alignment = Alignment(horizontal="right", vertical="center")
            # Compliance Status column (col 9)
            status_cell = ws.cell(row=row_idx, column=9)
            status_cell.alignment = Alignment(horizontal="center", vertical="center")
    else:
        ws.append(["No verified compliant players recorded for this club."])

    # Enable cell sorting (auto-filter)
    ws.auto_filter.ref = ws.dimensions

    # Vectorized column width calculation
    for col_idx in range(1, len(headers) + 1):
        col_letter = get_column_letter(col_idx)
        max_len = max(
            len(str(ws.cell(row=r, column=col_idx).value or ""))
            for r in range(1, max(ws.max_row + 1, 2))
        )
        ws.column_dimensions[col_letter].width = max(max_len + 2, 12)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


def generate_club_compliance_statement_docx(
    club_name: str,
    compliant_df: pd.DataFrame,
    output_path: Optional[str] = None
) -> io.BytesIO:
    """
    Generates an individual Word compliance statement document (.docx) for a specific club
    named '[Club_Name]_Season_2026_Compliance_Statement.docx'.
    Features an executive #1F4E78 Navy header, green status badge ('🟢 Account Compliant'),
    and an itemized table listing all approved compliant players.

    Inputs:
        club_name: Name of the cricket club.
        compliant_df: DataFrame containing the club's compliant player records.
        output_path: Optional file path to save the document to disk.

    Outputs:
        io.BytesIO: In-memory Word document stream (.docx).

    Helper Apps:
        finance_app.py.
    """
    doc = docx.Document()

    # Page Margins: 0.75 in
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    # 1. Executive Header Block (#1F4E78 Navy Theme)
    header_table = doc.add_table(rows=1, cols=1)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_table.autofit = False
    header_cell = header_table.cell(0, 0)
    header_cell.width = Inches(7.0)
    set_cell_background(header_cell, "1F4E78")
    set_cell_margins(header_cell, top=180, bottom=180, left=240, right=240)

    p_org = header_cell.paragraphs[0]
    p_org.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_org = p_org.add_run("NORTHERN CRICKET UNION")
    r_org.font.name = "Calibri"
    r_org.font.size = Pt(11)
    r_org.font.bold = True
    r_org.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    p_title = header_cell.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(3)
    p_title.paragraph_format.space_after = Pt(3)
    r_title = p_title.add_run("NCU Season 2026 - Roster Compliance Clearance Statement")
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(18)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    p_sub = header_cell.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Official Player Registration & Financial Clearance Certificate")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(10)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(0xD9, 0xE1, 0xF2)

    doc.add_paragraph()  # Spacing

    # 2. Club Summary Block & Green Status Badge
    summary_table = doc.add_table(rows=1, cols=2)
    summary_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    summary_table.autofit = False

    cell_meta = summary_table.cell(0, 0)
    cell_meta.width = Inches(4.2)
    set_cell_background(cell_meta, "F8FAFC")
    set_cell_margins(cell_meta, top=140, bottom=140, left=180, right=180)

    club_slug = re.sub(r"[^A-Za-z0-9]", "", club_name).upper()[:8]
    p_c1 = cell_meta.paragraphs[0]
    r_c1_lbl = p_c1.add_run("ISSUED TO:\n")
    r_c1_lbl.font.size = Pt(9)
    r_c1_lbl.font.bold = True
    r_c1_lbl.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)

    r_c1_val = p_c1.add_run(f"{club_name} Cricket Club\n")
    r_c1_val.font.size = Pt(13)
    r_c1_val.font.bold = True
    r_c1_val.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)

    r_c1_meta = p_c1.add_run(
        f"Clearance Ref: NCU-2026-CLR-{club_slug}\n"
        f"Billing Season: 2026 Season\n"
        f"Date of Issue: {datetime.now().strftime('%d %B %Y')}"
    )
    r_c1_meta.font.size = Pt(9.5)
    r_c1_meta.font.color.rgb = RGBColor(0x33, 0x41, 0x55)

    # Status Badge cell
    cell_badge = summary_table.cell(0, 1)
    cell_badge.width = Inches(2.8)
    set_cell_background(cell_badge, "F0FDF4")
    set_cell_margins(cell_badge, top=140, bottom=140, left=180, right=180)

    p_badge = cell_badge.paragraphs[0]
    p_badge.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_b_badge = p_badge.add_run("🟢 Account Compliant\n")
    r_b_badge.font.size = Pt(14)
    r_b_badge.font.bold = True
    r_b_badge.font.color.rgb = RGBColor(0x16, 0x65, 0x34)

    r_b_status = p_badge.add_run("STATUS: FULLY CLEARED\n")
    r_b_status.font.size = Pt(9)
    r_b_status.font.bold = True
    r_b_status.font.color.rgb = RGBColor(0x16, 0x65, 0x34)

    num_cleared = len(compliant_df)
    r_b_count = p_badge.add_run(f"Approved Compliant Players: {num_cleared:,}\n")
    r_b_count.font.size = Pt(10)
    r_b_count.font.bold = True
    r_b_count.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)

    r_b_note = p_badge.add_run("Registration & Fees 100% Reconciled")
    r_b_note.font.size = Pt(8.5)
    r_b_note.font.italic = True
    r_b_note.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    doc.add_paragraph()  # Spacing

    # 3. Itemized Table of Approved Players
    p_th = doc.add_paragraph()
    r_th = p_th.add_run("Itemized Roster of Approved & Cleared Players")
    r_th.font.name = "Calibri"
    r_th.font.size = Pt(13)
    r_th.font.bold = True
    r_th.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)
    p_th.paragraph_format.space_after = Pt(2)

    p_th_sub = doc.add_paragraph()
    r_th_sub = p_th_sub.add_run(
        "The following players are verified in Sport80 with approved registration and fee clearance for the 2026 season."
    )
    r_th_sub.font.size = Pt(9.5)
    r_th_sub.font.italic = True
    r_th_sub.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)
    p_th_sub.paragraph_format.space_after = Pt(6)

    table_cols = [
        ("#", Inches(0.4), WD_ALIGN_PARAGRAPH.CENTER),
        ("Player Name", Inches(2.2), WD_ALIGN_PARAGRAPH.LEFT),
        ("Category", Inches(0.9), WD_ALIGN_PARAGRAPH.CENTER),
        ("Date of Birth", Inches(1.0), WD_ALIGN_PARAGRAPH.CENTER),
        ("Fee Cleared", Inches(0.9), WD_ALIGN_PARAGRAPH.RIGHT),
        ("Payment Details", Inches(1.6), WD_ALIGN_PARAGRAPH.LEFT),
    ]

    items_table = doc.add_table(rows=1, cols=len(table_cols))
    items_table.style = "Table Grid"
    items_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    items_table.autofit = False

    # Header Row
    hdr_cells = items_table.rows[0].cells
    for i, (title, width, align) in enumerate(table_cols):
        cell = hdr_cells[i]
        cell.width = width
        set_cell_background(cell, "1F4E78")
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = align
        r = p.add_run(title)
        r.font.name = "Calibri"
        r.font.size = Pt(9)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # Data Rows
    if compliant_df.empty:
        r_empty = items_table.add_row()
        for i, (_, width, _) in enumerate(table_cols):
            cell = r_empty.cells[i]
            cell.width = width
            set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
            if i == 1:
                p = cell.paragraphs[0]
                r = p.add_run("No verified compliant player records found for this club.")
                r.font.size = Pt(9)
                r.font.italic = True
    else:
        sort_keys = compliant_df.apply(get_name_sort_key, axis=1)
        sorted_df = compliant_df.iloc[sort_keys.argsort()].reset_index(drop=True)
        for idx, (_, r) in enumerate(sorted_df.iterrows(), start=1):
            row = items_table.add_row()
            bg_color = "FFFFFF" if idx % 2 != 0 else "F8FAFC"

            name = str(r.get("Full_Name", "") or "").strip()
            category = str(r.get("Player Category", "") or ("Youth" if r.get("Is_Youth", False) else "Adult")).strip()
            dob = str(r.get("Date of Birth", "") or "-").strip()
            fee = r.get("Total_Paid", 0)
            try:
                fee_str = f"£{float(fee):,.2f}"
            except (ValueError, TypeError):
                fee_str = "£0.00"
            pay_details = str(r.get("Payment Details", "") or "-").strip()

            row_data = [
                (str(idx), WD_ALIGN_PARAGRAPH.CENTER),
                (name, WD_ALIGN_PARAGRAPH.LEFT),
                (category, WD_ALIGN_PARAGRAPH.CENTER),
                (dob, WD_ALIGN_PARAGRAPH.CENTER),
                (fee_str, WD_ALIGN_PARAGRAPH.RIGHT),
                (pay_details, WD_ALIGN_PARAGRAPH.LEFT),
            ]

            for i, (val, align) in enumerate(row_data):
                cell = row.cells[i]
                cell.width = table_cols[i][1]
                set_cell_background(cell, bg_color)
                set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
                p = cell.paragraphs[0]
                p.alignment = align
                run = p.add_run(val)
                run.font.name = "Calibri"
                run.font.size = Pt(8.5)

    doc.add_paragraph()  # Spacing

    # 4. Audit Trail & Verification Footnote
    foot_table = doc.add_table(rows=1, cols=1)
    foot_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    foot_table.autofit = False
    foot_cell = foot_table.cell(0, 0)
    foot_cell.width = Inches(7.0)
    set_cell_background(foot_cell, "F1F5F9")
    set_cell_margins(foot_cell, top=140, bottom=140, left=180, right=180)

    p_f = foot_cell.paragraphs[0]
    r_f1 = p_f.add_run("Northern Cricket Union • Official Registration & Compliance Clearance\n")
    r_f1.font.size = Pt(9)
    r_f1.font.bold = True
    r_f1.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)

    r_f2 = p_f.add_run(
        "This official clearance statement is issued by the Northern Cricket Union (NCU) based on verified Sport80 "
        "registration and revenue records. Certified players are fully cleared for participation in NCU competitions. "
        "Questions regarding roster clearance should be directed to the NCU Administration & Finance Office."
    )
    r_f2.font.size = Pt(8.5)
    r_f2.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    if output_path:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        doc.save(output_path)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


def generate_all_roster_clearance_archives(
    all_clubs: List[str],
    compliant_rosters: Dict[str, pd.DataFrame]
) -> io.BytesIO:
    """
    Loops through all clubs, generates both the OpenPyXL compliant roster Excel workbook
    ('[Club_Name]_Season_2026_Compliant_Roster.xlsx') and the python-docx compliance statement
    ('[Club_Name]_Season_2026_Compliance_Statement.docx'), and packages them into a flat
    in-memory ZIP compressed archive.

    Inputs:
        all_clubs: List of club names to generate statements for.
        compliant_rosters: Dictionary mapping club names to compliant player DataFrames.

    Outputs:
        io.BytesIO: Flat in-memory ZIP archive.

    Helper Apps:
        finance_app.py.
    """
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for club in all_clubs:
            club_str = str(club).strip()
            if not club_str or club_str in ["Unknown", "TOTAL", "NCU Pathway XI"]:
                continue
            safe_club = sanitize_club_filename(club_str)
            comp_df = compliant_rosters.get(club_str, pd.DataFrame())

            # 1. Excel Roster Builder
            excel_bio = generate_club_compliant_roster_excel(club_str, comp_df)
            excel_name = f"{safe_club}_Season_2026_Compliant_Roster.xlsx"
            zf.writestr(excel_name, excel_bio.getvalue())

            # 2. Word Clearance Statement Builder
            word_bio = generate_club_compliance_statement_docx(club_str, comp_df)
            word_name = f"{safe_club}_Season_2026_Compliance_Statement.docx"
            zf.writestr(word_name, word_bio.getvalue())

    zip_buffer.seek(0)
    return zip_buffer


# ==========================================
# SUB-VIEW A: SHARON INVOICING DASHBOARD
# ==========================================
def render_sharon_invoicing_dashboard() -> None:
    """
    Renders Sharon's master invoicing dashboard with 4-category summaries,
    dynamic row-hiding, OpenPyXL spreadsheet export with =SUM() formulas,
    executive Word/ZIP invoice generators, and per-club player drill-down audit inspector.

    Inputs:
        None.

    Outputs:
        None.

    Helper Apps:
        finance_app.py, engine.py.
    """
    # 1. Collect file modification signatures for caching
    audit_files = [
        "1. NCU_Registered_Players.xlsx",
        "2. NCU_Validated_Aliases_Master.xlsx",
        "12. NCU_Validated_Women's Aliases_Master.xlsx",
        "NCU_Mens_Master_ID_Mapping.xlsx",
        "NCU_Womens_Master_ID_Mapping.xlsx",
        "4. Unregistered_Manual_Map.xlsx",
        eng.get_default_revenue_file(),
        "NV Play NCU League and Saturday Cup batting stats for season.xlsx",
        "NV Play Women's Fixtures batting stats for season.xlsx",
        "NV Play Midweek League batting stats for season.xlsx",
        "NV Play NCU League and Saturday Cup player appearances for abandoned games.xlsx"
    ]
    file_signatures = tuple(
        (f, os.path.getmtime(f) if os.path.exists(f) else 0)
        for f in audit_files
    )

    # Sidebar Controls
    with st.sidebar:
        st.subheader("⚙️ Dashboard Controls")
        st.info("💡 Data is dynamically reconciled in-memory against Sport80 revenue and NV Play scorecards.")

        filter_nonzero = st.checkbox("Show Invoiced Clubs Only (> £0)", value=False, key="sharon_filter_nonzero")
        search_query = st.text_input("🔍 Search Club:", placeholder="e.g., Holywood, CSNI", key="sharon_search_query")
        sort_by = st.selectbox(
            "Sort Matrix By:",
            options=["Club Name (A-Z)", "Total Invoice Due (High to Low)"],
            index=0,
            key="sharon_sort_by"
        )

        st.divider()
        if st.button("🔄 Refresh / Re-run Audit", width="stretch", key="sharon_refresh_btn"):
            st.cache_data.clear()
            st.rerun()

        st.caption("Developed for Northern Cricket Union &bull; Sharon's Audit Schedule")

    # Load data
    matrix_df, detail_dfs = load_invoicing_matrix(file_signatures)
    compliant_rosters = extract_compliant_rosters(detail_dfs)

    if matrix_df.empty:
        st.warning("⚠️ No invoicing records found or data sources are missing.")
        return

    # Filter and sort data for display
    display_df = matrix_df.copy()
    if filter_nonzero:
        display_df = display_df[display_df["Total Invoice Due"] > 0]
    if search_query.strip():
        display_df = display_df[display_df["Club"].str.contains(search_query.strip(), case=False, na=False)]

    if sort_by == "Club Name (A-Z)":
        display_df = display_df.sort_values(by="Club", ascending=True).reset_index(drop=True)
    else:
        display_df = display_df.sort_values(by=["Total Invoice Due", "Club"], ascending=[False, True]).reset_index(drop=True)

    # 2. KPI Summary Cards
    total_invoiced = int(matrix_df["Total Invoice Due"].sum())
    clubs_with_invoices = int((matrix_df["Total Invoice Due"] > 0).sum())
    total_infractions = int(
        matrix_df["Count_Unregistered"].sum()
        + matrix_df["Count_Youth"].sum()
        + matrix_df["Count_Adult_Unpaid"].sum()
        + matrix_df["Count_Adult_Paid5"].sum()
        + matrix_df.get("Count_Transfers", pd.Series(0, index=matrix_df.index)).sum()
    )
    if not matrix_df.empty:
        max_idx = matrix_df["Total Invoice Due"].idxmax()
        top_club_name = str(matrix_df.loc[max_idx, "Club"])
        top_club_amt = int(matrix_df.loc[max_idx, "Total Invoice Due"])
    else:
        top_club_name = "None"
        top_club_amt = 0

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    with kpi1:
        st.metric("Total Invoiced Amount", f"£{total_invoiced:,}", help="Aggregate sum across all clubs and infraction categories")
    with kpi2:
        st.metric("Clubs with Invoices", str(clubs_with_invoices), help="Total clubs carrying an invoice balance due")
    with kpi3:
        st.metric("Total Infraction Items", f"{total_infractions:,}", help="Total individual player appearances / registrations flagged")
    with kpi4:
        st.metric("Largest Club Invoice", f"£{top_club_amt:,}", delta=top_club_name, delta_color="normal", help=f"Highest invoice due: {top_club_name}")

    st.markdown("<br>", unsafe_allow_html=True)

    # 3. Main Invoicing Matrix Table
    c_head1, c_head2 = st.columns([3, 1])
    with c_head1:
        st.subheader("📋 Official Club Invoicing Matrix")
        st.caption(f"Sorted by {sort_by}. All amounts in GBP (£).")
    with c_head2:
        export_df = matrix_df.copy()
        if sort_by == "Club Name (A-Z)":
            export_df = export_df.sort_values(by="Club", ascending=True).reset_index(drop=True)
        else:
            export_df = export_df.sort_values(by=["Total Invoice Due", "Club"], ascending=[False, True]).reset_index(drop=True)

        excel_bytes = generate_invoicing_workbook(export_df, detail_dfs)
        st.download_button(
            label="📥 Download Excel Ledger (.xlsx)",
            data=excel_bytes.getvalue(),
            file_name="NCU_Season_2026_Invoicing_Schedule.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            width="stretch",
            help="Exports complete audit schedule with OpenPyXL formatting and =SUM() formulas.",
            key="dl_sharon_excel_btn"
        )

    disp_cols = [
        "Club",
        "Unregistered Scorecard Players",
        "Unpaid Youth in Adult Cricket",
        "Unpaid Adults (£10 Shortfall)",
        "Adults Paid Youth Rate (£5 Shortfall)",
    ]
    if "Transfer Fees (£25)" in display_df.columns:
        disp_cols.append("Transfer Fees (£25)")
    disp_cols.append("Total Invoice Due")

    st.dataframe(
        display_df[disp_cols],
        width="stretch",
        hide_index=True,
        column_config=get_standard_column_config()
    )

    # Word Invoices Exporter Action Row
    st.markdown("<br>", unsafe_allow_html=True)
    c_inv_btn1, c_inv_btn2 = st.columns([1, 1])
    with c_inv_btn1:
        generate_invoices_btn = st.button(
            "📥 Generate Club Word Invoices",
            type="primary",
            width="stretch",
            help="Loops through all invoiced clubs and generates styled individual .docx invoices into 'club_invoices/'.",
            key="gen_sharon_word_invoices_btn"
        )

    if generate_invoices_btn:
        with st.spinner("Generating individual club Word invoices (.docx)..."):
            gen_paths = generate_all_club_word_invoices(
                matrix_df,
                detail_dfs,
                output_dir="club_invoices",
                only_invoiced=False
            )
            st.session_state["generated_invoices"] = gen_paths

    if "generated_invoices" in st.session_state and st.session_state["generated_invoices"]:
        gen_paths = st.session_state["generated_invoices"]
        st.success(f"✅ Successfully generated **{len(gen_paths)}** club invoice documents in `./club_invoices/`!")
        with c_inv_btn2:
            zip_bytes = create_invoices_zip(gen_paths)
            st.download_button(
                label=f"📦 Download All {len(gen_paths)} Word Invoices (.zip)",
                data=zip_bytes.getvalue(),
                file_name="NCU_Season_2026_Club_Word_Invoices.zip",
                mime="application/zip",
                width="stretch",
                help="Download all generated .docx club invoices in a single ZIP file.",
                key="dl_sharon_zip_btn"
            )

    # 3b. Roster Compliance Clearance Reports Exporter Action Row
    st.markdown("<br>", unsafe_allow_html=True)
    c_clr_btn1, c_clr_btn2 = st.columns([1, 1])
    with c_clr_btn1:
        export_roster_btn = st.button(
            "📥 Export Roster Clearance Statements",
            type="primary",
            width="stretch",
            help="Generates both Excel rosters and Word clearance statements for all clubs and packages them into a single flat in-memory ZIP download archive.",
            key="export_roster_clearance_btn"
        )

    if export_roster_btn:
        with st.spinner("Generating Roster Compliance Clearance statements and rosters for all clubs..."):
            all_clubs_list = sorted(set(matrix_df["Club"].dropna().tolist()).union(compliant_rosters.keys()))
            clearance_zip_bio = generate_all_roster_clearance_archives(all_clubs_list, compliant_rosters)
            st.session_state["roster_clearance_zip"] = clearance_zip_bio.getvalue()

    if "roster_clearance_zip" in st.session_state and st.session_state["roster_clearance_zip"]:
        with c_clr_btn2:
            st.download_button(
                label="📦 Download All Roster Clearance Statements (.zip)",
                data=st.session_state["roster_clearance_zip"],
                file_name="NCU_Season_2026_Roster_Compliance_Statements.zip",
                mime="application/zip",
                type="primary",
                width="stretch",
                help="Download all club compliant rosters (.xlsx) and clearance statements (.docx) in a single flat ZIP file.",
                key="dl_roster_clearance_zip_btn"
            )

    st.divider()

    # 4. Club Drill-Down Inspector
    st.subheader("🔍 Inspect Club Player Breakdown")
    st.caption("Select a club to see the exact player names and matches contributing to their invoice.")

    available_clubs = sorted(matrix_df["Club"].unique())
    selected_inspect_club = st.selectbox(
        "Select Club to Inspect:",
        options=available_clubs,
        index=0,
        key="finance_inspect_club"
    )

    if selected_inspect_club:
        club_unreg = detail_dfs["unreg"][detail_dfs["unreg"]["Club_Clean"] == selected_inspect_club]
        club_youth = detail_dfs["youth"][detail_dfs["youth"]["Club_Clean"] == selected_inspect_club]
        club_adult_u = detail_dfs["adult_unpaid"][detail_dfs["adult_unpaid"]["Club_Clean"] == selected_inspect_club]
        club_adult_p5 = detail_dfs["adult_paid5"][detail_dfs["adult_paid5"]["Club_Clean"] == selected_inspect_club]
        club_comp = compliant_rosters.get(selected_inspect_club, pd.DataFrame())
        all_transfers_df = detail_dfs.get("transfers", pd.DataFrame())
        club_transfers = all_transfers_df[all_transfers_df["To Club"] == selected_inspect_club] if not all_transfers_df.empty else pd.DataFrame()

        t_unreg, t_youth, t_adult_u, t_adult_p5, t_trans, t_comp = st.tabs([
            f"⚠️ Unregistered Scorecard Players ({len(club_unreg)})",
            f"👦 Unpaid Youth ({len(club_youth)})",
            f"👤 Unpaid Adults ({len(club_adult_u)})",
            f"⚡ Adults Paid Youth Rate ({len(club_adult_p5)})",
            f"🔄 Rule A13 Transfers ({len(club_transfers)})",
            f"🟢 Compliant Cleared Roster ({len(club_comp)})"
        ])

        with t_unreg:
            if not club_unreg.empty:
                cols_to_show = [c for c in ["Match_Player_Display", "Total_Matches", "Competitions", "Matches Played"] if c in club_unreg.columns]
                st.dataframe(club_unreg[cols_to_show], width="stretch", hide_index=True)
            else:
                st.success(f"No unregistered scorecard players flagged for {selected_inspect_club}.")

        with t_youth:
            if not club_youth.empty:
                cols_to_show = [c for c in ["Full_Name", "Date of Birth", "Age at Registration", "Total_Matches", "Payment Details", "Matches Played"] if c in club_youth.columns]
                st.dataframe(club_youth[cols_to_show], width="stretch", hide_index=True)
            else:
                st.success(f"No unpaid youth players flagged for {selected_inspect_club}.")

        with t_adult_u:
            if not club_adult_u.empty:
                cols_to_show = [c for c in ["Full_Name", "Date of Birth", "Age at Registration", "Total_Matches", "Payment Details", "Matches Played"] if c in club_adult_u.columns]
                st.dataframe(club_adult_u[cols_to_show], width="stretch", hide_index=True)
            else:
                st.success(f"No unpaid adult players flagged for {selected_inspect_club}.")

        with t_adult_p5:
            if not club_adult_p5.empty:
                cols_to_show = [c for c in ["Full_Name", "Date of Birth", "Age at Registration", "Total_Matches", "Payment Details", "Matches Played"] if c in club_adult_p5.columns]
                st.dataframe(club_adult_p5[cols_to_show], width="stretch", hide_index=True)
            else:
                st.success(f"No adults paying youth rate flagged for {selected_inspect_club}.")

        with t_trans:
            if not club_transfers.empty:
                cols_to_show = [c for c in ["Player", "From Club", "Transfer Number", "Transfer Date", "Fee Due (£)", "Fee Infraction"] if c in club_transfers.columns]
                st.dataframe(
                    club_transfers[cols_to_show],
                    width="stretch",
                    hide_index=True,
                    column_config=get_standard_column_config()
                )
            else:
                st.success(f"No transfer infractions flagged for {selected_inspect_club}.")

        with t_comp:
            if not club_comp.empty:
                c_dl1, c_dl2 = st.columns([1, 1])
                safe_sel_club = sanitize_club_filename(selected_inspect_club)
                with c_dl1:
                    club_excel_bio = generate_club_compliant_roster_excel(selected_inspect_club, club_comp)
                    st.download_button(
                        label=f"📥 Download {safe_sel_club} Compliant Roster (.xlsx)",
                        data=club_excel_bio.getvalue(),
                        file_name=f"{safe_sel_club}_Season_2026_Compliant_Roster.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        width="stretch",
                        key=f"dl_comp_excel_{safe_sel_club}"
                    )
                with c_dl2:
                    club_docx_bio = generate_club_compliance_statement_docx(selected_inspect_club, club_comp)
                    st.download_button(
                        label=f"📄 Download {safe_sel_club} Clearance Statement (.docx)",
                        data=club_docx_bio.getvalue(),
                        file_name=f"{safe_sel_club}_Season_2026_Compliance_Statement.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        width="stretch",
                        key=f"dl_comp_docx_{safe_sel_club}"
                    )
                st.markdown("<br>", unsafe_allow_html=True)
                cols_comp_show = [c for c in ["Full_Name", "Player Category", "Date of Birth", "Age at Registration", "Date Registered", "Total_Paid", "Payment Details", "Matches Played", "Compliance Status"] if c in club_comp.columns]
                st.dataframe(
                    club_comp[cols_comp_show],
                    width="stretch",
                    hide_index=True,
                    column_config=get_standard_column_config()
                )
            else:
                st.info(f"No compliant cleared players recorded for {selected_inspect_club}.")


# ==========================================
# SUB-VIEW B: REGISTRATION FEE AUDIT
# ==========================================
def render_registration_fee_audit() -> None:
    """
    Renders the Registration Fee Audit sub-view within the financial console.
    Cross-references registrations, aliases, revenue, and scorecard match appearances
    to track payment shortfalls and £10 / £5 adult-to-youth discrepancies.

    Inputs:
        None.

    Outputs:
        None.

    Helper Apps:
        finance_app.py, engine.py.
    """
    st.info("💡 **Tip:** Cross-references registrations, aliases, revenue, and match appearances to generate a 100% reconciled fee audit.")

    with st.container(border=True):
        st.subheader("📁 Input Files")
        st.markdown("Ensure the following files are present in the working directory:")
        st.markdown("- `1. NCU_Registered_Players.xlsx`")
        st.markdown("- `2. NCU_Validated_Aliases_Master.xlsx`")
        st.markdown("- `12. NCU_Validated_Women's Aliases_Master.xlsx`")
        st.markdown("- Player Registrations with DOB file (e.g. `Player_Registrations_for_2026_with_DOB-*.csv`)")
        st.markdown("- `4. Unregistered_Manual_Map.xlsx` *(optional manual club mapping for unregistered players)*")
        st.markdown("- The raw Sport80 Revenue Report (e.g. `revenue_report_il_from_*.xlsx`)")
        st.markdown("- *Plus the standard NV Play stats files (Sat, Women, Midweek)*")

        if st.button("🚀 Run Registration Fee Audit", type="primary", width="stretch", key="run_reg_fee_audit_btn"):
            with st.spinner("Processing audit (this may take 10-20 seconds)..."):
                try:
                    audit_file, timestamped_filename, df_summary, _ = eng.run_registration_fee_audit()
                    st.session_state['audit_outputs'] = (audit_file, timestamped_filename)
                    st.session_state['audit_summary_df'] = df_summary

                    # Save physical copies directly to "Output Files"
                    output_dir = "Output Files"
                    os.makedirs(output_dir, exist_ok=True)
                    date_str = datetime.now().strftime('%Y%m%d_%H%M%S')
                    excel_path = os.path.join(output_dir, f"NCU_Registration_Fee_Audit_{date_str}.xlsx")
                    docx_path = os.path.join(output_dir, f"NCU_Revenue_Anomalies_Report_{date_str}.docx")

                    with open(excel_path, "wb") as f_out:
                        f_out.write(audit_file.getvalue())
                    with open(docx_path, "wb") as f_out:
                        f_out.write(timestamped_filename.getvalue())

                    st.session_state['saved_paths'] = (excel_path, docx_path)
                except Exception as e:
                    st.error(f"❌ Error during audit: {str(e)}")

        # Display Disambiguation Flagged Violations live on the fee dashboard
        flagged_violations = st.session_state.get('fee_audit_violations', [])
        if flagged_violations:
            with st.expander(f"🚨 Disambiguation Flagged Unregistered Violations ({len(flagged_violations)} logged)", expanded=True):
                v_cols = st.columns(3)
                v_cols[0].metric("Flagged Unregistered Players", len(flagged_violations))
                total_fines = sum(float(v.get('Fine', 10.0)) for v in flagged_violations)
                v_cols[1].metric("Total Fines Assessed", f"£{total_fines:.2f}")
                v_cols[2].metric("Clubs Impacted", len(set(v.get('Club', '') for v in flagged_violations)))
                df_v = pd.DataFrame(flagged_violations)[['Player', 'Club', 'Matches', 'Fine', 'Timestamp', 'Status']]
                st.dataframe(df_v, width="stretch")

        if 'audit_outputs' in st.session_state:
            excel_io, doc_io = st.session_state['audit_outputs']
            st.success("✅ Audit complete! Your reports are ready to download below.")

            if 'audit_summary_df' in st.session_state and not st.session_state['audit_summary_df'].empty:
                with st.expander("📊 Financial Audit Reconciliation Summary", expanded=False):
                    st.dataframe(st.session_state['audit_summary_df'], width="stretch")

            if 'saved_paths' in st.session_state:
                p_excel, p_docx = st.session_state['saved_paths']
                st.info(f"📁 **Files also saved directly to:**\n- `{p_excel}`\n- `{p_docx}`")

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                date_str = datetime.now().strftime('%Y%m%d_%H%M%S')
                zip_file.writestr(f"NCU_Registration_Fee_Audit_{date_str}.xlsx", excel_io.getvalue())
                zip_file.writestr(f"NCU_Revenue_Anomalies_Report_{date_str}.docx", doc_io.getvalue())

            st.download_button(
                label="📦 Download Audit Results (ZIP)",
                data=zip_buffer.getvalue(),
                file_name=f"NCU_Registration_Fee_Audit_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip",
                width="stretch",
                type="primary",
                key="dl_audit_zip"
            )


# ==========================================
# SUB-VIEW C: UNREGISTERED PLAYER FINES GENERATOR
# ==========================================
def render_unregistered_fines_generator() -> None:
    """
    Renders the Unregistered Player Fines Generator sub-view within the financial console.
    Runs the registration audit engine to isolate match-day scorecard multiplier evaluations
    exclusively for unregistered player appearances and exports itemized Word documents.

    Inputs:
        None.

    Outputs:
        None.

    Helper Apps:
        finance_app.py, engine.py.
    """
    st.markdown("Automatically run the registration audit engine to generate a standalone fines report isolated exclusively to unregistered players.")

    if not DOCX_AVAILABLE:
        st.error("The `python-docx` library is not installed. Please run `pip install python-docx` to use this feature.")
        return

    st.subheader("Select League Domain")
    domain = st.radio("Choose the dataset domain to audit:", ["Men's", "Women's", "Midweek"], horizontal=True, key="unreg_domain")

    with st.sidebar:
        st.subheader("Select Date Range")
        start_date = st.date_input("Start Date", value=datetime.today() - timedelta(days=7), key="unreg_start")
        end_date = st.date_input("End Date", value=datetime.today(), key="unreg_end")
        st.divider()
        c_files = eng.DEFAULT_FILES.get(domain, eng.DEFAULT_FILES["Men's"])
        with st.expander("📁 File Path Configurations", expanded=False):
            f_reg = st.text_input("Official Registry (Excel)", value=c_files["reg"], key=f"unreg_reg_{domain}")
            f_alias = st.text_input("Aliases Master (Excel)", value=c_files["alias"], key=f"unreg_alias_{domain}")
            f_id_map = st.text_input("ID Mapping Master (Excel)", value=c_files.get("id_map", ""), key=f"unreg_id_map_{domain}")
            f_bat = st.text_input("Batting Stats (Excel)", value=c_files["bat"], key=f"unreg_bat_{domain}")
            f_bowl = st.text_input("Bowling Stats (Excel)", value=c_files["bowl"], key=f"unreg_bowl_{domain}")
            f_abandoned = st.text_input("Abandoned Games Stats (Excel)", value=c_files.get("abandoned", ""), key=f"unreg_ab_{domain}")
            f_revenue = st.text_input("Official Revenue Report (Excel)", value=c_files.get("revenue", eng.get_default_revenue_file()), key=f"unreg_revenue_{domain}")

            if domain != "Midweek":
                f_starring = st.text_input("Starring Master (Excel)", value=c_files["starring"], key=f"unreg_starring_{domain}")
                f_league = st.text_input("League Structure (Excel)", value=c_files["league"], key=f"unreg_league_{domain}")
                f_cup = st.text_input("Cup Master (Excel)", value=c_files.get("cup", eng.DEFAULT_CUP_FILE), key=f"unreg_cup_{domain}")
            else:
                f_starring = st.text_input("Men's Starring Master (Excel)", value=eng.DEFAULT_FILES["Men's"]["starring"], key="unreg_mw_starring")
                f_weekend_league = st.text_input("Weekend League Structure (Excel)", value=eng.DEFAULT_FILES["Men's"]["league"], key="unreg_wknd_league")
                f_midweek_league = st.text_input("Midweek League Structure (Excel)", value=c_files["league"], key="unreg_mw_league")

    include_irish = False
    if domain == "Men's":
        include_irish = st.toggle("Include Irish Competitions in Audit?", value=False, key="unreg_irish_check")
        if include_irish:
            with st.sidebar:
                with st.expander("📁 Irish File Path Configurations", expanded=False):
                    f_irish_bat = st.text_input("Irish Batting Stats (Excel)", value=c_files.get("irish_bat", eng.DEFAULT_IRISH_BAT_FILE), key="unreg_irish_bat")
                    f_irish_bowl = st.text_input("Irish Bowling Stats (Excel)", value=c_files.get("irish_bowl", eng.DEFAULT_IRISH_BOWL_FILE), key="unreg_irish_bowl")

    st.divider()
    if st.button("📄 Run Engine & Generate Unregistered Report", type="primary", width="stretch", key="run_unreg_fines_btn"):
        files_to_check = [f_reg, f_alias, f_bat, f_bowl]
        if f_id_map:
            files_to_check.append(f_id_map)
        if f_abandoned:
            files_to_check.append(f_abandoned)
        if domain != "Midweek":
            files_to_check.extend([f_starring, f_league, f_cup])
            if domain == "Men's" and include_irish:
                files_to_check.extend([f_irish_bat, f_irish_bowl])
        else:
            files_to_check.extend([f_starring, f_weekend_league, f_midweek_league])

        missing_files = [f for f in files_to_check if f and not os.path.exists(f)]
        if missing_files:
            st.error("Cannot find the following files:\n\n" + "\n".join([f"- {f}" for f in missing_files]))
        elif start_date > end_date:
            st.error("Start Date cannot be after End Date.")
        else:
            with st.spinner("Running registration audit and compiling unregistered fines report..."):
                try:
                    start_ts = pd.to_datetime(start_date)
                    end_ts = pd.to_datetime(end_date)

                    if domain != "Midweek":
                        if domain == "Men's" and include_irish:
                            audit_excel_io, _ = eng.run_registration_audit(domain, start_ts, end_ts, f_reg, f_alias, f_starring, f_league, f_bat, f_bowl, f_irish_bat, f_irish_bowl, f_cup, f_abandoned=f_abandoned, f_id_map=f_id_map, f_revenue=f_revenue)
                        else:
                            audit_excel_io, _ = eng.run_registration_audit(domain, start_ts, end_ts, f_reg, f_alias, f_starring, f_league, f_bat, f_bowl, f_cup=f_cup, f_abandoned=f_abandoned, f_id_map=f_id_map, f_revenue=f_revenue)
                    else:
                        audit_excel_io, _ = eng.run_midweek_registration_audit(start_ts, end_ts, f_reg, f_alias, f_starring, f_weekend_league, f_midweek_league, f_bat, f_bowl, f_abandoned=f_abandoned, f_id_map=f_id_map, f_revenue=f_revenue)

                    audit_dfs = audit_excel_io.dfs if hasattr(audit_excel_io, 'dfs') else audit_excel_io
                    doc_io = eng.generate_unregistered_fines_only(audit_dfs)
                    st.success("✅ Unregistered Fines report generated successfully!")
                    st.download_button(
                        f"📥 Download {domain} Unregistered Fines Report (Word)",
                        data=doc_io.getvalue(),
                        file_name=f"NCU_{domain.replace('''s''', '')}_Unreg_Fines_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary",
                        width="stretch",
                        key="dl_unreg_fines_btn"
                    )
                except Exception as e:
                    st.error(f"An error occurred during processing: {str(e)}")


# ==========================================
# SUB-VIEW D: CLUB FINES GENERATOR
# ==========================================
def render_club_fines_generator() -> None:
    """
    Renders the Club Fines Generator sub-view within the financial console.
    Evaluates registration infractions and merges them with forfeited match schedules
    to generate an administrative club-by-club fines report.

    Inputs:
        None.

    Outputs:
        None.

    Helper Apps:
        finance_app.py, engine.py.
    """
    st.markdown("Automatically run the registration audit engine to find violations and merge them with forfeited matches to generate a club-by-club fines report.")

    if not DOCX_AVAILABLE:
        st.error("The `python-docx` library is not installed. Please run `pip install python-docx` to use this feature.")
        return

    st.subheader("Select League Domain")
    domain = st.radio("Choose the dataset domain to audit:", ["Men's", "Women's", "Midweek"], horizontal=True, key="club_domain")

    with st.sidebar:
        st.subheader("Select Date Range")
        start_date = st.date_input("Start Date", value=datetime.today() - timedelta(days=7), key="club_fines_start")
        end_date = st.date_input("End Date", value=datetime.today(), key="club_fines_end")
        st.divider()
        c_files = eng.DEFAULT_FILES.get(domain, eng.DEFAULT_FILES["Men's"])
        with st.expander("📁 File Path Configurations", expanded=False):
            f_reg = st.text_input("Official Registry (Excel)", value=c_files["reg"], key=f"club_fines_reg_{domain}")
            f_alias = st.text_input("Aliases Master (Excel)", value=c_files["alias"], key=f"club_fines_alias_{domain}")
            f_id_map = st.text_input("ID Mapping Master (Excel)", value=c_files.get("id_map", ""), key=f"club_fines_id_map_{domain}")
            f_bat = st.text_input("Batting Stats (Excel)", value=c_files["bat"], key=f"club_fines_bat_{domain}")
            f_bowl = st.text_input("Bowling Stats (Excel)", value=c_files["bowl"], key=f"club_fines_bowl_{domain}")
            f_abandoned = st.text_input("Abandoned Games Stats (Excel)", value=c_files.get("abandoned", ""), key=f"club_fines_ab_{domain}")
            f_revenue = st.text_input("Official Revenue Report (Excel)", value=c_files.get("revenue", eng.get_default_revenue_file()), key=f"club_fines_revenue_{domain}")

            if domain != "Midweek":
                f_starring = st.text_input("Starring Master (Excel)", value=c_files["starring"], key=f"club_fines_starring_{domain}")
                f_league = st.text_input("League Structure (Excel)", value=c_files["league"], key=f"club_fines_league_{domain}")
                f_cup = st.text_input("Cup Master (Excel)", value=c_files.get("cup", eng.DEFAULT_CUP_FILE), key=f"club_fines_cup_{domain}")
            else:
                f_starring = st.text_input("Men's Starring Master (Excel)", value=eng.DEFAULT_FILES["Men's"]["starring"], key="club_fines_mw_starring")
                f_weekend_league = st.text_input("Weekend League Structure (Excel)", value=eng.DEFAULT_FILES["Men's"]["league"], key="club_fines_wknd_league")
                f_midweek_league = st.text_input("Midweek League Structure (Excel)", value=c_files["league"], key="club_fines_mw_league")

    include_irish = False
    if domain == "Men's":
        include_irish = st.toggle("Include Irish Competitions in Audit?", value=False, key="club_fines_irish_check")
        if include_irish:
            with st.sidebar:
                with st.expander("📁 Irish File Path Configurations", expanded=False):
                    f_irish_bat = st.text_input("Irish Batting Stats (Excel)", value=c_files.get("irish_bat", eng.DEFAULT_IRISH_BAT_FILE), key="club_fines_irish_bat")
                    f_irish_bowl = st.text_input("Irish Bowling Stats (Excel)", value=c_files.get("irish_bowl", eng.DEFAULT_IRISH_BOWL_FILE), key="club_fines_irish_bowl")

    st.divider()
    st.subheader("Forfeited Matches Data")
    default_forfeit_path = os.path.join("test_data", "Team Fines for forfeiting matches 2026.xlsx") if os.environ.get("TEST_MODE", "0") == "1" else "Team Fines for forfeiting matches 2026.xlsx"
    col1, col2 = st.columns([1, 2])
    with col1:
        use_default_forfeit = st.toggle(f"Use local '{default_forfeit_path}'", value=os.path.exists(default_forfeit_path), key="club_fines_forfeit_toggle")
    with col2:
        f_forfeit = st.file_uploader("Or Upload Forfeits Excel File", type=["xlsx"], key="club_fines_forfeit_upload")

    st.divider()
    if st.button("📄 Run Engine & Generate Fines Report", type="primary", width="stretch", key="run_club_fines_btn"):
        forfeit_path = f_forfeit if f_forfeit is not None else (default_forfeit_path if use_default_forfeit and os.path.exists(default_forfeit_path) else None)
        files_to_check = [f_reg, f_alias, f_bat, f_bowl]
        if f_id_map:
            files_to_check.append(f_id_map)
        if f_abandoned:
            files_to_check.append(f_abandoned)
        if domain != "Midweek":
            files_to_check.extend([f_starring, f_league, f_cup])
            if domain == "Men's" and include_irish:
                files_to_check.extend([f_irish_bat, f_irish_bowl])
        else:
            files_to_check.extend([f_starring, f_weekend_league, f_midweek_league])

        missing_files = [f for f in files_to_check if f and not os.path.exists(f)]
        if missing_files:
            st.error("Cannot find the following files:\n\n" + "\n".join([f"- {f}" for f in missing_files]))
        elif start_date > end_date:
            st.error("Start Date cannot be after End Date.")
        else:
            with st.spinner("Running registration audit and compiling fines report..."):
                try:
                    start_ts = pd.to_datetime(start_date)
                    end_ts = pd.to_datetime(end_date)

                    if domain != "Midweek":
                        if domain == "Men's" and include_irish:
                            audit_excel_io, _ = eng.run_registration_audit(domain, start_ts, end_ts, f_reg, f_alias, f_starring, f_league, f_bat, f_bowl, f_irish_bat, f_irish_bowl, f_cup, f_abandoned=f_abandoned, f_id_map=f_id_map, f_revenue=f_revenue)
                        else:
                            audit_excel_io, _ = eng.run_registration_audit(domain, start_ts, end_ts, f_reg, f_alias, f_starring, f_league, f_bat, f_bowl, f_cup=f_cup, f_abandoned=f_abandoned, f_id_map=f_id_map, f_revenue=f_revenue)
                    else:
                        audit_excel_io, _ = eng.run_midweek_registration_audit(start_ts, end_ts, f_reg, f_alias, f_starring, f_weekend_league, f_midweek_league, f_bat, f_bowl, f_abandoned=f_abandoned, f_id_map=f_id_map, f_revenue=f_revenue)

                    audit_dfs = audit_excel_io.dfs if hasattr(audit_excel_io, 'dfs') else audit_excel_io
                    doc_io = eng.generate_club_fines_report(audit_dfs, forfeit_path, start_ts, end_ts)
                    st.success("✅ Fines report generated successfully!")
                    st.download_button(
                        f"📥 Download {domain} Fines Report (Word)",
                        data=doc_io.getvalue(),
                        file_name=f"NCU_{domain.replace('''s''', '')}_Fines_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary",
                        width="stretch",
                        key="dl_club_fines_btn"
                    )
                except Exception as e:
                    st.error(f"An error occurred during processing: {str(e)}")


# ==========================================
# MAIN APPLICATION INTERFACE
# ==========================================
def main() -> None:
    """
    Entry point for the NCU Finance & Invoicing Command Center.
    Initializes executive sub-navigation and orchestrates the 4 financial console sub-views.

    Inputs:
        None.

    Outputs:
        None.

    Helper Apps:
        None (Main application entry point).
    """
    with st.sidebar:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #1F4E78 0%, #15375B 100%); color: #FFFFFF; padding: 14px 18px; border-radius: 8px; margin-bottom: 15px; border-left: 5px solid #D4AF37;">
            <h3 style="margin: 0; color: #FFFFFF; font-weight: 700; font-size: 1.15rem;">💳 Finance Console</h3>
            <p style="margin: 3px 0 0 0; color: #E2E8F0; font-size: 0.82rem;">Northern Cricket Union &bull; Admin Suite</p>
        </div>
        """, unsafe_allow_html=True)

        finance_mode = st.radio(
            "Financial Console View:",
            [
                "Sharon Invoicing Dashboard",
                "Registration Fee Audit",
                "Unregistered Player Fines Generator",
                "Club Fines Generator"
            ],
            key="finance_app_nav"
        )
        st.divider()

    # Dynamic Header Banner based on active sub-view
    banner_titles = {
        "Sharon Invoicing Dashboard": (
            "💳 Sharon Invoicing Dashboard",
            "Official Club Invoicing Schedule & Audit Breakdown &bull; Season 2026"
        ),
        "Registration Fee Audit": (
            "💰 Registration Fee Audit",
            "Cross-Reference Registrations, Revenue & Scorecards &bull; £10/£5 Shortfall Tracker"
        ),
        "Unregistered Player Fines Generator": (
            "💸 Unregistered Player Fines Generator",
            "Match-Day Scorecard Infraction Evaluator & Standalone Word Document Exporter"
        ),
        "Club Fines Generator": (
            "🛡️ Club Fines Generator",
            "General Team Administrative Fine Manager & Forfeited Match Consolidator"
        )
    }

    title_text, subtitle_text = banner_titles.get(finance_mode, ("💳 NCU Finance Console", "Season 2026"))
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #1F4E78 0%, #15375B 100%); color: #FFFFFF; padding: 18px 24px; border-radius: 8px; margin-bottom: 20px; border-left: 6px solid #D4AF37; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.08);">
        <h2 style="margin: 0; color: #FFFFFF; font-weight: 700; font-size: 1.6rem;">{title_text}</h2>
        <p style="margin: 4px 0 0 0; color: #E2E8F0; font-size: 0.95rem;">{subtitle_text}</p>
    </div>
    """, unsafe_allow_html=True)

    if finance_mode == "Sharon Invoicing Dashboard":
        render_sharon_invoicing_dashboard()
    elif finance_mode == "Registration Fee Audit":
        render_registration_fee_audit()
    elif finance_mode == "Unregistered Player Fines Generator":
        render_unregistered_fines_generator()
    elif finance_mode == "Club Fines Generator":
        render_club_fines_generator()


if __name__ == "__main__":
    main()
