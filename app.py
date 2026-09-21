# ==========================================
# app.py
# ==========================================
import streamlit as st
import pandas as pd
import numpy as np
import os
import io
import re
import json
import zipfile
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Set, Tuple
import importlib

import engine as eng
import starring_rules as sr
importlib.reload(sr)

try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False
    
import warnings
warnings.filterwarnings('ignore')

# ==========================================
# STREAMLIT CACHED EXCEL LOADERS
# ==========================================
def get_excel_df(filepath):
    return eng.get_excel_df(filepath)

def get_excel_sheet_df(filepath, sheet_name=None, header='infer'):
    return eng.get_excel_sheet_df(filepath, sheet_name=sheet_name, header=header)

@st.cache_data(show_spinner="Loading and parsing contacts directory...")
def cached_parse_club_contacts(filepath, mtime):
    if not filepath or not os.path.exists(filepath):
        return pd.DataFrame(), {}, []
    return eng.parse_club_contacts_matrix(filepath)

@st.cache_data(show_spinner="Parsing uploaded contacts...")
def cached_parse_uploaded_contacts(file_bytes):
    return eng.parse_club_contacts_matrix(io.BytesIO(file_bytes))

def get_club_contacts_data(filepath):
    if not filepath or not os.path.exists(filepath):
        return pd.DataFrame(), {}, []
    mtime = os.path.getmtime(filepath)
    return cached_parse_club_contacts(filepath, mtime)

@st.cache_resource(show_spinner="Loading match scorecard data for eligibility tracker...")
def cached_starring_pipeline(domain, f_reg, f_alias, f_bat, f_bowl, f_abandoned, f_id_map, f_irish_bat, f_irish_bowl, mtimes):
    return eng.build_starring_inactivity_pipeline_data(
        domain=domain, f_reg=f_reg, f_alias=f_alias, f_bat=f_bat, f_bowl=f_bowl,
        f_irish_bat=f_irish_bat, f_irish_bowl=f_irish_bowl, f_abandoned=f_abandoned, f_id_map=f_id_map
    )

# ==========================================
# CONTACT LINK FORMATTERS
# ==========================================
def format_tel_link(phone_str):
    if not phone_str or str(phone_str).strip().lower() in ['nan', 'none', '-', '']:
        return "—"
    text = str(phone_str).strip().replace('.0', '')
    digits = re.sub(r'[^\d+]', '', text)
    if digits.startswith('0'):
        digits = '+44' + digits[1:]
    return f'<a href="tel:{digits}" style="text-decoration:none; font-weight:600; color:#0066cc;">📞 {text}</a>'

def format_mail_link(email_str):
    if not email_str or str(email_str).strip().lower() in ['nan', 'none', '-', '']:
        return "—"
    text = str(email_str).strip()
    return f'<a href="mailto:{text}" style="text-decoration:none; font-weight:600; color:#0066cc;">✉️ {text}</a>'

def get_tier_group(tier: Any) -> Tuple[int, str]:
    """
    Classifies a club role or team tier into a standardized group order and display label.
    Delegates directly to eng.get_tier_group for centralized category taxonomy.
    """
    return eng.get_tier_group(tier)

def resolve_starring_history_path() -> Optional[str]:
    """
    Dynamically resolves the absolute file path for 'NCU_Club_Starring_History.xlsx'
    across local development environments and production deployment directories.
    """
    candidates = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "NCU_Club_Starring_History.xlsx"),
        os.path.abspath("NCU_Club_Starring_History.xlsx"),
        os.path.join(os.getcwd(), "NCU_Club_Starring_History.xlsx"),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None

def render_contact_grid(df_items, num_cols=3):
    """
    Renders contact cards in row-by-row chunks of columns.
    Ensures that when viewed on mobile/tablet devices, cards stack in exact
    sequential order rather than unrolling column-by-column vertically.
    """
    if df_items.empty:
        return
    for i in range(0, len(df_items), num_cols):
        batch = df_items.iloc[i : i + num_cols]
        cols = st.columns(num_cols)
        for c_idx, (_, row) in enumerate(batch.iterrows()):
            with cols[c_idx]:
                with st.container(border=True):
                    role_title = row.get('Role', 'Club Official')
                    official_name = row.get('Name', 'Not Listed')
                    phone_val = row.get('Phone', '')
                    email_val = row.get('Email', '')

                    st.markdown(f"### {role_title}")
                    st.markdown(f"**👤 {official_name}**")
                    st.markdown(f"**Category:** `{row.get('Team Tier', 'General')}`")
                    st.markdown(format_tel_link(phone_val), unsafe_allow_html=True)
                    st.markdown(format_mail_link(email_val), unsafe_allow_html=True)

# ==========================================
# USER CONFIGURATIONS & PERSISTENCE
# ==========================================
MAIN_HEADER_SIZE = "28px" 
CONFIG_FILE = "threshold_settings.json"

PAGE_TITLES = {
    "dashboard": "📊 2026 Season Summary Dashboard",
    "player_doc": "📄 Player Word Doc Generator",
    "reg_checks": "🛡️ Weekend Registration and Starring Checks",
    "midweek_checks": "🛡️ Midweek Registration & Starring Check",
    "starring_reports": "🚨 Club Starring & Inactivity Exporter",
    "starring_registry": "🏏 Club Starring Registry & Historical Eligibility Tracker",
    "milestones_report": "🏆 League Milestones Report",
    "club_contacts": "📇 Club Contacts & Officials Directory",
    "csv_importer": "📥 NV Play CSV Match Stats Importer",
    "player_disambiguation": "🔗 Player Disambiguation & ID Mapping"
}

DEFAULT_THRESHOLDS = {
    "t1_runs": 200, "t1_bmat": 5, "t1_wick": 15, "t1_mmat": 5,
    "t2_runs": 150, "t2_bmat": 5, "t2_wick": 10, "t2_mmat": 5,
    "t3_runs": 100, "t3_bmat": 3, "t3_wick": 5,  "t3_mmat": 3,
    "t4_runs": 50,  "t4_bmat": 3, "t4_wick": 3,  "t4_mmat": 3,
    "w1_runs": 100, "w1_bmat": 5, "w1_wick": 10, "w1_mmat": 5,
    "w2_runs": 25,  "w2_bmat": 2, "w2_wick": 2,  "w2_mmat": 2,
    "mw_min_runs": 50, "mw_min_innings": 0, "mw_min_wickets": 5
}

def init_threshold_store():
    if "threshold_store" not in st.session_state:
        store = dict(DEFAULT_THRESHOLDS)
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    saved = json.load(f)
                    for k, v in saved.items():
                        if k in DEFAULT_THRESHOLDS:
                            store[k] = int(v)
            except Exception:
                pass
        st.session_state["threshold_store"] = store

def get_threshold_val(key):
    init_threshold_store()
    is_zero = st.session_state.get("disable_thresholds", False)
    if is_zero: return 0
    if key in st.session_state: st.session_state["threshold_store"][key] = st.session_state[key]
    return st.session_state["threshold_store"].get(key, DEFAULT_THRESHOLDS.get(key, 0))

def save_threshold_settings():
    init_threshold_store()
    if not st.session_state.get("disable_thresholds", False):
        for k in DEFAULT_THRESHOLDS:
            if k in st.session_state: st.session_state["threshold_store"][k] = st.session_state[k]
    with open(CONFIG_FILE, "w") as f:
        json.dump(st.session_state["threshold_store"], f, indent=4)

def reset_threshold_settings():
    st.session_state["threshold_store"] = dict(DEFAULT_THRESHOLDS)
    for k, v in DEFAULT_THRESHOLDS.items(): st.session_state[k] = v
    if os.path.exists(CONFIG_FILE):
        try: os.remove(CONFIG_FILE)
        except Exception: pass

def toggle_zero_thresholds():
    init_threshold_store()
    is_zero = st.session_state.get("disable_thresholds", False)
    store = st.session_state["threshold_store"]
    for k in DEFAULT_THRESHOLDS:
        st.session_state[k] = 0 if is_zero else store.get(k, DEFAULT_THRESHOLDS[k])

# ==========================================
# PAGE CONFIGURATION & STYLING
# ==========================================
st.set_page_config(page_title="NCU Cricket Hub", page_icon="🏏", layout="wide")

def inject_custom_styles() -> None:
    """
    Injects standardized Streamlit CSS styling into the active runtime session.
    Harmonizes headers, action buttons, metrics, and prevents text truncation
    on status badges without external CSS/JS dependencies.
    """
    st.markdown(f"""
    <style>
        /* Typography & Header hierarchy */
        h1 {{ font-size: {MAIN_HEADER_SIZE} !important; font-weight: 700; }}
        
        /* Action buttons & download controls */
        div.stButton > button {{ white-space: nowrap !important; }}
        div.stButton > button[kind="primary"] {{ border-radius: 8px; padding: 0.5rem 1.5rem; }}
        div.stDownloadButton > button:first-child {{
            background-color: #0066cc;
            color: #ffffff;
            border-radius: 8px;
            border: none;
            padding: 0.5rem 1.5rem;
        }}
        div.stDownloadButton > button:first-child:hover {{
            background-color: #0052a3;
            color: #ffffff;
        }}
        
        /* Metric cards */
        [data-testid="stMetricValue"] {{ font-size: 1.8rem; font-weight: 700; }}
        [data-testid="metric-container"] {{
            background-color: rgba(250, 250, 250, 0.1);
            border: 1px solid rgba(128, 128, 128, 0.2);
            padding: 15px;
            border-radius: 10px;
        }}
        
        /* Status Badges & Tag Protections: Prevents text clipping and truncation */
        span[data-testid="stBadge"],
        div[data-testid="stBadge"],
        .status-badge {{
            white-space: nowrap !important;
            overflow: visible !important;
            text-overflow: unset !important;
            font-weight: 600 !important;
            display: inline-flex !important;
            align-items: center !important;
        }}

        /* Center alignment for DOM/HTML table headers */
        th, th[role="columnheader"] {{
            text-align: center !important;
        }}
        th.col-left, th[data-col-align="left"], td.col-left, td[data-col-align="left"] {{
            text-align: left !important;
        }}
    </style>
    """, unsafe_allow_html=True)

def center_col_label(label: str, target_width: int = 12) -> str:
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
    for st.dataframe and st.data_editor across this application.
    Centered columns have center-balanced header labels mirroring centered cells,
    while left-aligned columns remain strictly left-aligned.
    """
    return {
        "Rank": st.column_config.Column(center_col_label("Rank", 10), alignment="center", width="small"),
        "XI_Level": st.column_config.Column(center_col_label("XI Level", 12), alignment="center", width="small"),
        "Club": st.column_config.TextColumn("Club Name", alignment="left", width="medium"),
        "Club Name": st.column_config.TextColumn("Club Name", alignment="left", width="medium"),
        "Player": st.column_config.TextColumn("Player", alignment="left", width="medium"),
        "Full Name": st.column_config.TextColumn("Player Name", alignment="left", width="medium"),
        "Full_Name": st.column_config.TextColumn("Player Name", alignment="left", width="medium"),
        "Starred Tier": st.column_config.Column(center_col_label("Starred Tier", 16), alignment="center", width="small"),
        "Registered": st.column_config.TextColumn(
            center_col_label("Registered", 14),
            help="Official Sport80 registry verification (✔️ / ✅ Registered | ❌ Unregistered)",
            alignment="center",
            width="small"
        ),
        "Administrative Status": st.column_config.TextColumn(
            "Administrative Status",
            help="NCU Rule A11/A12 Roster Eligibility & De-starring Requirement status",
            width="large",
            alignment="left"
        ),
        "Compliance Status": st.column_config.TextColumn(
            center_col_label("Compliance Status", 20),
            help="Audit verification state",
            alignment="center",
            width="medium"
        ),
        "Eligible Appearances": st.column_config.NumberColumn(center_col_label("Eligible Appearances", 24), format="%d", alignment="center", width="small"),
        "Total_Matches": st.column_config.NumberColumn(center_col_label("Matches Played", 18), format="%d", alignment="center", width="small"),
        "Missed Matches": st.column_config.NumberColumn(center_col_label("Missed Matches", 18), format="%d", alignment="center", width="small"),
        "Days Inactive": st.column_config.NumberColumn(center_col_label("Days Inactive", 16), format="%d", alignment="center", width="small"),
        "Last Played Date": st.column_config.Column(center_col_label("Last Played Date", 20), alignment="center", width="medium"),
        "Transfer Number": st.column_config.Column(center_col_label("Transfer Number", 18), alignment="center", width="small"),
        "Transfer Date": st.column_config.Column(center_col_label("Transfer Date", 16), alignment="center", width="medium"),
        "Fee Due (£)": st.column_config.NumberColumn(center_col_label("Fee Due (£)", 14), format="£%.2f", alignment="center", width="small"),
        "Fee Infraction": st.column_config.CheckboxColumn("£25 Late Fee Infraction (>= 1 Apr)", width="medium"),
        "Total_Paid": st.column_config.NumberColumn(center_col_label("Fee Cleared", 14), format="£%.2f", alignment="center", width="small"),
    }

inject_custom_styles()

# ==========================================
# SIDEBAR NAVIGATION
# ==========================================
with st.sidebar:
    st.title("🏏 NCU Cricket Hub")
    # TEST MODE banner — only shown when launched via Launch_Test_Suite.bat
    if os.environ.get("TEST_MODE", "0") == "1":
        st.warning("⚠️ **TEST MODE**\n\nUsing sample data from `test_data/`.\nNo production files are being read or written.")
        st.divider()

    st.header("🛠️ Navigation")
    
    app_mode = st.radio(
        "Choose a module to run:",
        [
            "2026 Season Summary Dashboard",
            "Club Starring Registry & Historical Eligibility Tracker",
            "CSV Match Stats Importer",
            "Player Disambiguation & ID Mapping",
            "Registration Checks",
            "Midweek Registration & Starring Check",
            "Starring & Inactivity Reports",
            "Player Word Doc Generator",
            "Club Contacts Directory",
        ]
    )
    st.divider()

# ==========================================
# TOOL 1: WORD DOC GENERATOR
# ==========================================
if app_mode == "Player Word Doc Generator":
    st.title(PAGE_TITLES["player_doc"])
    
    if not DOCX_AVAILABLE:
        st.error("The `python-docx` library is not installed. Please run `pip install python-docx` to use this feature.")
    else:
        st.info("💡 **Tip:** This utility extracts match-by-match stats and builds a formatted season report ready for print.")
        st.subheader("Select League Domain")
        domain = st.radio("Choose the dataset domain to search:", ["Men's", "Women's", "Midweek"], horizontal=True)

        with st.sidebar:
            c_files = eng.DEFAULT_FILES[domain]
            with st.expander("📁 File Path Configurations", expanded=False):
                f_reg = st.text_input("Official Registry (Excel)", value=c_files["reg"], key=f"doc_reg_{domain}")
                f_alias = st.text_input("Aliases Master (Excel)", value=c_files["alias"], key=f"doc_alias_{domain}")
                f_id_map = st.text_input("ID Mapping Master (Excel)", value=c_files.get("id_map", ""), key=f"doc_id_map_{domain}")
                f_bat = st.text_input("Batting Stats (Excel)", value=c_files["bat"], key=f"doc_bat_{domain}")
                f_bowl = st.text_input("Bowling Stats (Excel)", value=c_files["bowl"], key=f"doc_bowl_{domain}")
                f_abandoned = st.text_input("Abandoned Games Stats (Excel)", value=c_files.get("abandoned", ""), key=f"doc_ab_{domain}")
                f_league = st.text_input("League Structure (Excel)", value=c_files["league"], key=f"doc_league_{domain}")
                f_cup = st.text_input("Cup Master (Excel)", value=c_files.get("cup", eng.DEFAULT_CUP_FILE), key=f"doc_cup_{domain}")
        
        include_irish = False
        if domain == "Men's":
            include_irish = st.toggle("Include Irish Competitions in Player Report?", value=False)
            if include_irish:
                with st.sidebar:
                    with st.expander("📁 Irish File Path Configurations", expanded=False):
                        f_irish_bat = st.text_input("Irish Batting Stats (Excel)", value=c_files.get("irish_bat", eng.DEFAULT_IRISH_BAT_FILE), key="doc_irish_bat")
                        f_irish_bowl = st.text_input("Irish Bowling Stats (Excel)", value=c_files.get("irish_bowl", eng.DEFAULT_IRISH_BOWL_FILE), key="doc_irish_bowl")

        if 'doc_last_domain' not in st.session_state or st.session_state.doc_last_domain != domain:
            st.session_state.player_search_active = False
            st.session_state.doc_last_domain = domain

        with st.container(border=True):
            st.subheader("🔍 Search Player Database")
            search_query = st.text_input("Enter the player's full name or scorecard alias:", placeholder="e.g., Joe Bloggs")
            col_btn, _ = st.columns([1.5, 4])
            with col_btn:
                execute_search = st.button("🔍 Search Player", type="primary", width="stretch")

        if 'player_search_active' not in st.session_state:
            st.session_state.player_search_active = False

        if execute_search:
            st.session_state.player_search_active = True
            st.session_state.player_search_query = search_query
            st.session_state.data_loaded = False 

        if st.session_state.player_search_active:
            current_query = st.session_state.player_search_query
            files_to_check = [f_reg, f_alias, f_bat, f_bowl, f_league, f_cup]
            if f_id_map: files_to_check.append(f_id_map)
            if domain == "Men's" and include_irish: files_to_check.extend([f_irish_bat, f_irish_bowl])
                
            missing_files = [f for f in files_to_check if not os.path.exists(f)]
            if missing_files:
                st.error(f"Cannot find the following files:\n\n" + "\n".join([f"- {f}" for f in missing_files]))
            elif not current_query:
                st.warning("Please enter a player name to search.")
            else:
                if not st.session_state.get('data_loaded'):
                    with st.spinner("Searching datasets and building options..."):
                        reg_players = get_excel_df(f_reg)
                        aliases = get_excel_df(f_alias)
                        id_map_df = get_excel_df(f_id_map) if f_id_map and os.path.exists(f_id_map) else pd.DataFrame()
                        id_map = eng.build_id_map(id_map_df)
                        batting = get_excel_df(f_bat)
                        bowling = get_excel_df(f_bowl)
                        abandoned_df = get_excel_df(f_abandoned) if f_abandoned and os.path.exists(f_abandoned) else pd.DataFrame()
                        
                        if domain == "Men's" and include_irish:
                            if os.path.exists(f_irish_bat): batting = pd.concat([batting, get_excel_df(f_irish_bat)], ignore_index=True)
                            if os.path.exists(f_irish_bowl): bowling = pd.concat([bowling, get_excel_df(f_irish_bowl)], ignore_index=True)

                        alias_map = eng.build_alias_map(aliases, domain)
                        player_club_map = eng.build_player_club_map(reg_players, alias_map, domain, id_map_df=id_map_df)

                        batting['Name'] = batting.apply(lambda r: eng.resolve_player_from_row(r, r['Name'], id_map, alias_map, player_club_map)[0], axis=1)
                        bowling['Bowler'] = bowling.apply(lambda r: eng.resolve_player_from_row(r, r['Bowler'], id_map, alias_map, player_club_map)[0], axis=1)
                        if not abandoned_df.empty:
                            ab_name_col = 'Name' if 'Name' in abandoned_df.columns else abandoned_df.columns[1]
                            abandoned_df['Cleaned Name'] = abandoned_df.apply(lambda r: eng.resolve_player_from_row(r, r[ab_name_col], id_map, alias_map, player_club_map)[0], axis=1)
                            ab_grp_col = 'Group' if 'Group' in abandoned_df.columns else ('Match' if 'Match' in abandoned_df.columns else abandoned_df.columns[0])
                            abandoned_df['Group'] = abandoned_df[ab_grp_col].apply(lambda x: eng.doc_format_cricket_names(x, domain))

                        batting['Group'] = batting['Group'].apply(lambda x: eng.doc_format_cricket_names(x, domain))
                        bowling['Group'] = bowling['Group'].apply(lambda x: eng.doc_format_cricket_names(x, domain))

                        clean_q = current_query.strip().lower()
                        target_official_names = set()

                        if 'Input Name (Scorecard/Stats)' in aliases.columns and 'Official Registered Name' in aliases.columns:
                            alias_matches = aliases[
                                aliases['Input Name (Scorecard/Stats)'].astype(str).str.contains(clean_q, case=False, na=False) |
                                aliases['Official Registered Name'].astype(str).str.contains(clean_q, case=False, na=False)
                            ]
                            target_official_names.update(alias_matches['Official Registered Name'].dropna().astype(str).str.strip().tolist())
                        else:
                            for _, row in aliases.iterrows():
                                if clean_q in str(row.iloc[0]).lower() or clean_q in str(row.iloc[1]).lower():
                                    target_official_names.add(str(row.iloc[1]).strip())

                        if 'Full Name' in reg_players.columns:
                            reg_matches = reg_players[reg_players['Full Name'].astype(str).str.contains(clean_q, case=False, na=False)]
                            target_official_names.update(reg_matches['Full Name'].dropna().astype(str).str.strip().tolist())

                        if id_map_df is not None and not id_map_df.empty:
                            if 'NV_Play_Name' in id_map_df.columns and 'Sport80_Name' in id_map_df.columns:
                                id_matches = id_map_df[
                                    id_map_df['NV_Play_Name'].astype(str).str.contains(clean_q, case=False, na=False) |
                                    id_map_df['Sport80_Name'].astype(str).str.contains(clean_q, case=False, na=False)
                                ]
                                target_official_names.update(id_matches['Sport80_Name'].dropna().astype(str).str.strip().tolist())

                        bat_direct = batting[batting['Name'].astype(str).str.contains(clean_q, case=False, na=False)]['Name'].unique().tolist()
                        bowl_direct = bowling[bowling['Bowler'].astype(str).str.contains(clean_q, case=False, na=False)]['Bowler'].unique().tolist()
                        target_official_names.update(bat_direct + bowl_direct)

                        matched_batting_list, matched_bowling_list, matched_abandoned_list = [], [], []
                        for off_name in target_official_names:
                            matched_batting_list.append(batting[batting['Name'].astype(str).str.contains(off_name, case=False, na=False)])
                            matched_bowling_list.append(bowling[bowling['Bowler'].astype(str).str.contains(off_name, case=False, na=False)])
                            if not abandoned_df.empty:
                                matched_abandoned_list.append(abandoned_df[abandoned_df['Cleaned Name'].astype(str).str.contains(off_name, case=False, na=False)])

                        matched_batting = pd.concat(matched_batting_list) if matched_batting_list else pd.DataFrame()
                        matched_bowling = pd.concat(matched_bowling_list) if matched_bowling_list else pd.DataFrame()
                        matched_abandoned = pd.concat(matched_abandoned_list) if matched_abandoned_list else pd.DataFrame()

                        if not matched_batting.empty: matched_batting = matched_batting[~matched_batting.index.duplicated(keep='first')].reset_index(drop=True)
                        if not matched_bowling.empty: matched_bowling = matched_bowling[~matched_bowling.index.duplicated(keep='first')].reset_index(drop=True)

                        found_batters = matched_batting['Name'].dropna().unique().tolist() if not matched_batting.empty else []
                        found_bowlers = matched_bowling['Bowler'].dropna().unique().tolist() if not matched_bowling.empty else []
                        found_ab = matched_abandoned['Cleaned Name'].dropna().unique().tolist() if not matched_abandoned.empty else []
                        unique_dict = {}
                        for p in (found_batters + found_bowlers + found_ab):
                            p_str = str(p).strip()
                            k = p_str.lower()
                            if k not in unique_dict or ('Jnr' in p_str or 'Snr' in p_str):
                                unique_dict[k] = p_str
                        raw_unique_players = list(unique_dict.values())
                        
                        def player_sort_key(name):
                            pure_name = eng.extract_pure_player_name(name)
                            mapped = alias_map.get(name.lower(), name.lower())
                            club = player_club_map.get(mapped.lower(), "Unknown Club").lower()
                            parts = pure_name.split()
                            surname = parts[-1].lower() if len(parts) > 1 else (parts[0].lower() if parts else "")
                            firstnames = " ".join(parts[:-1]).lower() if len(parts) > 1 else ""
                            return (surname, firstnames, club)
                            
                        st.session_state.matched_batting = matched_batting
                        st.session_state.matched_bowling = matched_bowling
                        st.session_state.matched_abandoned = matched_abandoned
                        st.session_state.unique_players = sorted(raw_unique_players, key=player_sort_key)
                        st.session_state.reg_players = reg_players
                        st.session_state.aliases_df = aliases
                        st.session_state.player_club_map = player_club_map
                        st.session_state.id_map_df = id_map_df
                        st.session_state.data_loaded = True

                matched_batting = st.session_state.matched_batting
                matched_bowling = st.session_state.matched_bowling
                matched_abandoned = st.session_state.matched_abandoned
                unique_players = st.session_state.unique_players
                reg_players = st.session_state.reg_players
                aliases_df = st.session_state.aliases_df
                id_map_df = st.session_state.get('id_map_df', None)

                if matched_batting.empty and matched_bowling.empty and matched_abandoned.empty:
                    st.error(f"No statistics found for '{current_query}'. Please try another name.")
                else:
                    def get_club_for_player(name):
                        qualifier = eng.extract_player_club_qualifier(name)
                        if qualifier: return qualifier
                        club = st.session_state.player_club_map.get(name.lower(), None)
                        if not club or str(club).lower() in ['nan', 'none', '', 'unknown club']:
                            a_map = eng.build_alias_map(aliases_df, domain)
                            mapped = a_map.get(name.lower(), name.lower())
                            club = st.session_state.player_club_map.get(mapped.lower(), None)
                        if club and str(club).lower() not in ['nan', 'none', '', 'unknown club']:
                            return str(club).replace(" Cricket Club", "").replace(" CC", "").strip()
                        
                        p_bat = matched_batting[matched_batting['Name'] == name] if not matched_batting.empty else None
                        p_bowl = matched_bowling[matched_bowling['Bowler'] == name] if not matched_bowling.empty else None
                        inferred = eng.infer_player_club(name, p_bat, p_bowl, domain)
                        if inferred != "Unknown_Club": return inferred
                        
                        return "Unknown Club"

                    def format_player_display(name):
                        pure = eng.extract_pure_player_name(name)
                        club_clean = get_club_for_player(name)
                        playing_name = eng.get_player_playing_name(pure, aliases=aliases_df, id_map_df=id_map_df, club=club_clean)
                        if club_clean and str(club_clean).lower() not in ['unknown club', 'nan', 'none', '']:
                            return f"{playing_name} ({club_clean})"
                        return playing_name

                    if len(unique_players) == 1:
                        active_player = unique_players[0]
                        pure_registered_name = eng.extract_pure_player_name(active_player)
                        club_clean = get_club_for_player(active_player)
                        p_aliases = eng.get_player_aliases(pure_registered_name, aliases=aliases_df, id_map_df=id_map_df, club=club_clean)
                        st.success(f"Found Match: {format_player_display(active_player)}")
                        
                        p_bat = matched_batting[matched_batting['Name'].astype(str).str.lower() == active_player.lower()] if not matched_batting.empty else pd.DataFrame()
                        p_bowl = matched_bowling[matched_bowling['Bowler'].astype(str).str.lower() == active_player.lower()] if not matched_bowling.empty else pd.DataFrame()
                        p_ab = matched_abandoned[matched_abandoned['Cleaned Name'].astype(str).str.lower() == active_player.lower()] if not matched_abandoned.empty else pd.DataFrame()
                        
                        league_df = get_excel_df(f_league)
                        cup_df = get_excel_df(f_cup)
                        if not league_df.empty and 'Team' in league_df.columns and 'League' in league_df.columns:
                            league_dict = dict(zip(league_df['Team'], league_df['League']))
                        else:
                            league_dict = None

                        playing_name = eng.get_player_playing_name(pure_registered_name, aliases=aliases_df, id_map_df=id_map_df, club=club_clean)
                        doc_io, filename = eng.generate_single_player_doc(active_player, p_bat, p_bowl, reg_players, domain, aliases_list=p_aliases, player_abandoned=p_ab, league_dict=league_dict, cup_df=cup_df, id_map_df=id_map_df, playing_name=playing_name)
                        st.download_button("📥 Download Player Word Document", data=doc_io.getvalue(), file_name=filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary")
                    else:
                        st.warning(f"Multiple players match '{current_query}'. Please select the players to generate reports for.")
                        select_all = st.toggle("Select all players")
                        selected_players = st.multiselect("Select players:", options=unique_players, default=unique_players if select_all else [], format_func=format_player_display)
                        
                        if selected_players:
                            if len(selected_players) == 1:
                                active_player = selected_players[0]
                                pure_registered_name = eng.extract_pure_player_name(active_player)
                                club_clean = get_club_for_player(active_player)
                                p_aliases = eng.get_player_aliases(pure_registered_name, aliases=aliases_df, id_map_df=id_map_df, club=club_clean)
                                p_bat = matched_batting[matched_batting['Name'].astype(str).str.lower() == active_player.lower()] if not matched_batting.empty else pd.DataFrame()
                                p_bowl = matched_bowling[matched_bowling['Bowler'].astype(str).str.lower() == active_player.lower()] if not matched_bowling.empty else pd.DataFrame()
                                p_ab = matched_abandoned[matched_abandoned['Cleaned Name'].astype(str).str.lower() == active_player.lower()] if not matched_abandoned.empty else pd.DataFrame()
                                
                                league_df = get_excel_df(f_league)
                                cup_df = get_excel_df(f_cup)
                                if not league_df.empty and 'Team' in league_df.columns and 'League' in league_df.columns:
                                    league_dict = dict(zip(league_df['Team'], league_df['League']))
                                else:
                                    league_dict = None

                                playing_name = eng.get_player_playing_name(pure_registered_name, aliases=aliases_df, id_map_df=id_map_df, club=club_clean)
                                doc_io, filename = eng.generate_single_player_doc(active_player, p_bat, p_bowl, reg_players, domain, aliases_list=p_aliases, player_abandoned=p_ab, league_dict=league_dict, cup_df=cup_df, id_map_df=id_map_df, playing_name=playing_name)
                                st.download_button(f"📥 Download Report for {format_player_display(active_player)}", data=doc_io.getvalue(), file_name=filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", key="dl_single_multi")
                            else:
                                zip_buffer = io.BytesIO()
                                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                                    for active_player in selected_players:
                                        pure = eng.extract_pure_player_name(active_player)
                                        club_clean = get_club_for_player(active_player)
                                        p_aliases = eng.get_player_aliases(pure, aliases=aliases_df, id_map_df=id_map_df, club=club_clean)
                                        p_bat = matched_batting[matched_batting['Name'].astype(str).str.lower() == active_player.lower()] if not matched_batting.empty else pd.DataFrame()
                                        p_bowl = matched_bowling[matched_bowling['Bowler'].astype(str).str.lower() == active_player.lower()] if not matched_bowling.empty else pd.DataFrame()
                                        p_ab = matched_abandoned[matched_abandoned['Cleaned Name'].astype(str).str.lower() == active_player.lower()] if not matched_abandoned.empty else pd.DataFrame()
                                        league_df = get_excel_df(f_league)
                                        cup_df = get_excel_df(f_cup)
                                        if not league_df.empty and 'Team' in league_df.columns and 'League' in league_df.columns:
                                            league_dict = dict(zip(league_df['Team'], league_df['League']))
                                        else:
                                            league_dict = None

                                        playing_name = eng.get_player_playing_name(pure, aliases=aliases_df, id_map_df=id_map_df, club=club_clean)
                                        doc_io, filename = eng.generate_single_player_doc(active_player, p_bat, p_bowl, reg_players, domain, aliases_list=p_aliases, player_abandoned=p_ab, league_dict=league_dict, cup_df=cup_df, id_map_df=id_map_df, playing_name=playing_name)
                                        zip_file.writestr(filename, doc_io.getvalue())
                                        
                                st.download_button(f"📦 Download Reports for {len(selected_players)} Players (ZIP)", data=zip_buffer.getvalue(), file_name=f"Player_Reports_{current_query.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.zip", mime="application/zip", type="primary", key="dl_zip_multi")

# ==========================================
# 2026 SEASON SUMMARY DASHBOARD
# ==========================================
elif app_mode == "2026 Season Summary Dashboard":
    eng.render_season_summary_dashboard()

# ==========================================
# TOOL 2: WEEKEND REGISTRATION CHECKS
# ==========================================
elif app_mode == "Registration Checks":
    st.title(PAGE_TITLES["reg_checks"])
    st.markdown("Automate weekend audits by cross-referencing match logs against master registries and starring lists.")
    
    if not DOCX_AVAILABLE:
        st.error("The `python-docx` library is not installed. Please run `pip install python-docx` to use this feature.")
    else:
        st.subheader("Select League Domain")
        domain = st.radio("Choose the dataset domain to audit:", ["Men's", "Women's"], horizontal=True)

        with st.sidebar:
            st.subheader("Select Date Range")
            start_date = st.date_input("Start Date", value=datetime.today() - timedelta(days=1))
            end_date = st.date_input("End Date", value=datetime.today())
            st.divider()           

        with st.sidebar:
            c_files = eng.DEFAULT_FILES[domain]
            with st.expander("📁 File Path Configurations", expanded=False):
                f_reg = st.text_input("Official Registry (Excel)", value=c_files["reg"], key=f"reg_check_reg_{domain}")
                f_alias = st.text_input("Aliases Master (Excel)", value=c_files["alias"], key=f"reg_check_alias_{domain}")
                f_id_map = st.text_input("ID Mapping Master (Excel)", value=c_files.get("id_map", ""), key=f"reg_check_id_map_{domain}")
                f_starring = st.text_input("Starring Master (Excel)", value=c_files["starring"], key=f"reg_check_starring_{domain}")
                f_cup = st.text_input("Cup Master (Excel)", value=c_files.get("cup", eng.DEFAULT_CUP_FILE), key=f"reg_check_cup_{domain}")
                f_league = st.text_input("League Structure (Excel)", value=c_files["league"], key=f"reg_check_league_{domain}")
                f_bat = st.text_input("Batting Stats (Excel)", value=c_files["bat"], key=f"reg_check_bat_{domain}")
                f_bowl = st.text_input("Bowling Stats (Excel)", value=c_files["bowl"], key=f"reg_check_bowl_{domain}")
                f_abandoned = st.text_input("Abandoned Games Stats (Excel)", value=c_files.get("abandoned", ""), key=f"reg_check_ab_{domain}")
                f_revenue = st.text_input("Official Revenue Report (Excel)", value=c_files.get("revenue", eng.get_default_revenue_file()), key=f"reg_check_revenue_{domain}")
        
        include_irish = False
        if domain == "Men's":
            include_irish = st.toggle("Include Irish Competitions in Audit?", value=False)
            if include_irish:
                with st.sidebar:
                    with st.expander("📁 Irish File Path Configurations", expanded=False):
                        f_irish_bat = st.text_input("Irish Batting Stats (Excel)", value=c_files.get("irish_bat", eng.DEFAULT_IRISH_BAT_FILE), key="reg_irish_bat")
                        f_irish_bowl = st.text_input("Irish Bowling Stats (Excel)", value=c_files.get("irish_bowl", eng.DEFAULT_IRISH_BOWL_FILE), key="reg_irish_bowl")

        st.subheader("Run Audit Engine")
        if st.button("🚀 Execute Security Audit", type="primary"):
            files_to_check = [f for f in [f_reg, f_alias, f_league, f_bat, f_bowl, f_cup] if f]
            if f_id_map: files_to_check.append(f_id_map)
            if f_starring: files_to_check.append(f_starring)
            if f_abandoned: files_to_check.append(f_abandoned)
            if domain == "Men's" and include_irish:
                if f_irish_bat: files_to_check.append(f_irish_bat)
                if f_irish_bowl: files_to_check.append(f_irish_bowl)
                
            missing_files = [f for f in files_to_check if not os.path.exists(f)]
            if missing_files:
                st.error(f"Cannot find the following files:\n\n" + "\n".join([f"- {f}" for f in missing_files]))
            elif start_date > end_date:
                st.error("Start Date cannot be after End Date.")
            else:
                with st.spinner("Parsing match logs and evaluating registration rules..."):
                    try:
                        start_ts = pd.to_datetime(start_date)
                        end_ts = pd.to_datetime(end_date)
                        
                        if domain == "Men's" and include_irish:
                            excel_io, doc_io = eng.run_registration_audit(domain, start_ts, end_ts, f_reg, f_alias, f_starring, f_league, f_bat, f_bowl, f_irish_bat, f_irish_bowl, f_cup, f_abandoned=f_abandoned, f_id_map=f_id_map, f_revenue=f_revenue)
                        else:
                            excel_io, doc_io = eng.run_registration_audit(domain, start_ts, end_ts, f_reg, f_alias, f_starring, f_league, f_bat, f_bowl, f_cup=f_cup, f_abandoned=f_abandoned, f_id_map=f_id_map, f_revenue=f_revenue)
                        
                        try:
                            if hasattr(excel_io, 'dfs'):
                                df_unreg = excel_io.dfs.get("Unregistered Matches", pd.DataFrame())
                                df_deemed = excel_io.dfs.get("Deemed Registered", pd.DataFrame())
                                df_starring_viols = excel_io.dfs.get("Starring Violations", pd.DataFrame())
                            else:
                                excel_io.seek(0)
                                with pd.ExcelFile(excel_io) as xf:
                                    df_unreg = xf.parse("Unregistered Matches") if "Unregistered Matches" in xf.sheet_names else pd.DataFrame()
                                    df_deemed = xf.parse("Deemed Registered") if "Deemed Registered" in xf.sheet_names else pd.DataFrame()
                                    df_starring_viols = xf.parse("Starring Violations") if "Starring Violations" in xf.sheet_names else pd.DataFrame()
                            
                            unreg_count = len(df_unreg) if not df_unreg.empty and 'Status' not in df_unreg.columns else 0
                            deemed_count = len(df_deemed) if not df_deemed.empty and 'Status' not in df_deemed.columns else 0
                            star_count = len(df_starring_viols) if not df_starring_viols.empty and 'Status' not in df_starring_viols.columns else 0
                        except:
                            unreg_count, deemed_count, star_count = 0, 0, 0

                        st.success("✅ Audit complete!")
                        st.subheader("📊 Audit Discrepancy Overview")
                        m_col1, m_col2, m_col3 = st.columns(3)
                        with m_col1: st.metric(label="⚠️ Unregistered Match Appearances", value=unreg_count, delta=f"{unreg_count} Flagged", delta_color="inverse")
                        with m_col2: st.metric(label="ℹ️ Deemed Registered Records", value=deemed_count, delta=f"{deemed_count} Tracked", delta_color="off")
                        with m_col3: st.metric(label="🚨 Starring Violations", value=star_count, delta=f"{star_count} Flagged", delta_color="inverse")

                        if unreg_count > 0 or deemed_count > 0 or star_count > 0:
                            st.subheader("📋 Audit Report Previews")
                            if unreg_count > 0:
                                with st.expander("⚠️ Unregistered Matches"): st.dataframe(df_unreg, width="stretch", hide_index=True)
                            if deemed_count > 0:
                                with st.expander("ℹ️ Deemed Registered Players"): st.dataframe(df_deemed, width="stretch", hide_index=True)
                            if star_count > 0:
                                with st.expander("🚨 Starring Violations"): st.dataframe(df_starring_viols, width="stretch", hide_index=True)

                        st.divider()
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                            date_str = f"{start_ts.strftime('%d-%m-%Y')}_to_{end_ts.strftime('%d-%m-%Y')}"
                            prefix = domain.replace("'", "")
                            zip_file.writestr(f"{prefix}_Audit_Database_{date_str}.xlsx", excel_io.getvalue())
                            zip_file.writestr(f"{prefix}_Audit_Report_{date_str}.docx", doc_io.getvalue())
                                
                        st.download_button("📦 Download Audit Results (ZIP)", data=zip_buffer.getvalue(), file_name=f"{domain.replace('''s''', '')}_Registration_Audit_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip", mime="application/zip", type="primary")
                    except Exception as e:
                        st.error(f"An error occurred during processing: {str(e)}")

# ==========================================
# TOOL 3: MIDWEEK LEAGUE REGISTRATION CHECKS
# ==========================================
elif app_mode == "Midweek Registration & Starring Check":
    st.title(PAGE_TITLES["midweek_checks"])
    st.markdown("Automate Midweek audits, enforcing the strict Junior League 3 weekend ceiling logic.")
    
    if not DOCX_AVAILABLE:
        st.error("The `python-docx` library is not installed. Please run `pip install python-docx` to use this feature.")
    else:
        with st.sidebar:
            st.subheader("Select Date Range")
            start_date = st.date_input("Start Date", value=datetime.today() - timedelta(days=1))
            end_date = st.date_input("End Date", value=datetime.today())
            st.divider() 
            with st.expander("📁 File Path Configurations", expanded=False):
                f_reg = st.text_input("Official Registry (Excel)", value=eng.DEFAULT_FILES["Midweek"]["reg"], key="mw_check_reg")
                f_alias = st.text_input("Aliases Master (Excel)", value=eng.DEFAULT_FILES["Midweek"]["alias"], key="mw_check_alias")
                f_id_map = st.text_input("ID Mapping Master (Excel)", value=eng.DEFAULT_FILES["Midweek"].get("id_map", ""), key="mw_check_id_map")
                f_starring = st.text_input("Men's Starring Master (Excel)", value=eng.DEFAULT_FILES["Men's"]["starring"], key="mw_check_starring")
                f_weekend_league = st.text_input("Weekend League Structure (Excel)", value=eng.DEFAULT_FILES["Men's"]["league"], key="mw_check_wknd_league")
                f_midweek_league = st.text_input("Midweek League Structure (Excel)", value=eng.DEFAULT_FILES["Midweek"]["league"], key="mw_check_mw_league")
                f_bat = st.text_input("Midweek Batting Stats (Excel)", value=eng.DEFAULT_FILES["Midweek"]["bat"], key="mw_check_bat")
                f_bowl = st.text_input("Midweek Bowling Stats (Excel)", value=eng.DEFAULT_FILES["Midweek"]["bowl"], key="mw_check_bowl")
                f_abandoned = st.text_input("Abandoned Games Stats (Excel)", value=eng.DEFAULT_FILES["Midweek"].get("abandoned", ""), key="mw_check_ab")
                f_revenue = st.text_input("Official Revenue Report (Excel)", value=eng.DEFAULT_FILES["Midweek"].get("revenue", eng.get_default_revenue_file()), key="mw_check_revenue")

        st.subheader("Run Midweek Audit Engine")
        if st.button("🚀 Execute Midweek Audit", type="primary"):
            files_to_check = [f for f in [f_reg, f_alias, f_starring, f_weekend_league, f_midweek_league, f_bat, f_bowl] if f]
            if f_id_map: files_to_check.append(f_id_map)
            if f_abandoned: files_to_check.append(f_abandoned)
            missing_files = [f for f in files_to_check if not os.path.exists(f)]
            
            if missing_files:
                st.error(f"Cannot find the following files:\n\n" + "\n".join([f"- {f}" for f in missing_files]))
            elif start_date > end_date:
                st.error("Start Date cannot be after End Date.")
            else:
                with st.spinner("Parsing match logs and evaluating Midweek eligibility rules..."):
                    try:
                        start_ts = pd.to_datetime(start_date)
                        end_ts = pd.to_datetime(end_date)
                        excel_io, doc_io = eng.run_midweek_registration_audit(start_ts, end_ts, f_reg, f_alias, f_starring, f_weekend_league, f_midweek_league, f_bat, f_bowl, f_abandoned=f_abandoned, f_id_map=f_id_map, f_revenue=f_revenue)
                        
                        try:
                            excel_io.seek(0)
                            with pd.ExcelFile(excel_io) as xf:
                                df_unreg = xf.parse("Unregistered Matches") if "Unregistered Matches" in xf.sheet_names else pd.DataFrame()
                                df_deemed = xf.parse("Deemed Registered") if "Deemed Registered" in xf.sheet_names else pd.DataFrame()
                                df_starring_viols = xf.parse("Starring Violations") if "Starring Violations" in xf.sheet_names else pd.DataFrame()
                            unreg_count = len(df_unreg) if not df_unreg.empty and 'Status' not in df_unreg.columns else 0
                            deemed_count = len(df_deemed) if not df_deemed.empty and 'Status' not in df_deemed.columns else 0
                            star_count = len(df_starring_viols) if not df_starring_viols.empty and 'Status' not in df_starring_viols.columns else 0
                        except:
                            unreg_count, deemed_count, star_count = 0, 0, 0

                        st.success("✅ Audit complete!")
                        st.subheader("📊 Midweek Discrepancy Overview")
                        m_col1, m_col2, m_col3 = st.columns(3)
                        with m_col1: st.metric(label="⚠️ Unregistered Midweek Players", value=unreg_count, delta=f"{unreg_count} Flagged", delta_color="inverse")
                        with m_col2: st.metric(label="ℹ️ Deemed Registered Players", value=deemed_count, delta=f"{deemed_count} Tracked", delta_color="off")
                        with m_col3: st.metric(label="🚨 Midweek Starring Ceiling Violations", value=star_count, delta=f"{star_count} Flagged", delta_color="inverse")

                        if unreg_count > 0 or deemed_count > 0 or star_count > 0:
                            st.subheader("📋 Audit Report Previews")
                            if unreg_count > 0:
                                with st.expander("⚠️ Unregistered Midweek Matches"): st.dataframe(df_unreg, width="stretch", hide_index=True)
                            if deemed_count > 0:
                                with st.expander("ℹ️ Deemed Registered Players"): st.dataframe(df_deemed, width="stretch", hide_index=True)
                            if star_count > 0:
                                with st.expander("🚨 Midweek Ceiling Violations (Junior 3 & Above Starred players)"): st.dataframe(df_starring_viols, width="stretch", hide_index=True)

                        st.divider()
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                            date_str = f"{start_ts.strftime('%d-%m-%Y')}_to_{end_ts.strftime('%d-%m-%Y')}"
                            zip_file.writestr(f"Midweek_Audit_Database_{date_str}.xlsx", excel_io.getvalue())
                            zip_file.writestr(f"Midweek_Audit_Report_{date_str}.docx", doc_io.getvalue())
                                
                        st.download_button("📦 Download Audit Results (ZIP)", data=zip_buffer.getvalue(), file_name=f"Midweek_Registration_Audit_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip", mime="application/zip", type="primary")
                    except Exception as e:
                        st.error(f"An error occurred during processing: {str(e)}")

# ==========================================
# TOOL: CLUB STARRING REGISTRY & HISTORICAL ELIGIBILITY TRACKER (RULES A10–A14)
# ==========================================
elif app_mode == "Club Starring Registry & Historical Eligibility Tracker":
    st.title(PAGE_TITLES["starring_registry"])
    st.markdown("Automated validation and administrative tracking engine based on official **NCU Rules A10–A14**.")

    st.subheader("Select Dataset Domain & Club")
    col_sel1, col_sel2 = st.columns([1, 2])
    with col_sel1:
        domain = st.radio("Choose domain:", ["Men's", "Women's"], horizontal=True, key="registry_domain")

    c_files = eng.DEFAULT_FILES[domain]
    f_reg = c_files.get("reg", "")
    f_alias = c_files.get("alias", "")
    f_starring = c_files.get("starring", "")
    f_bat = c_files.get("bat", "")
    f_bowl = c_files.get("bowl", "")

    # Load starring lists
    starring_df = pd.DataFrame()
    parsed_club_dict = {}
    if f_starring and os.path.exists(f_starring):
        starring_df, parsed_club_dict = eng.cached_parse_starring_data(f_starring, os.path.getmtime(f_starring))

    available_clubs = sorted(list(parsed_club_dict.keys())) if parsed_club_dict else sorted(list(eng.NCU_ALL_CLUBS))

    with col_sel2:
        selected_club = st.selectbox(
            "Select Club:",
            options=available_clubs,
            index=0 if available_clubs else None,
            key=f"registry_club_select_{domain}"
        )

    club_star_df = parsed_club_dict.get(selected_club, pd.DataFrame()) if selected_club else pd.DataFrame()

    # Determine automated senior team count
    all_club_counts = eng.get_all_club_team_counts()
    auto_num_teams = sr.get_club_senior_team_count(
        club_name=selected_club,
        domain=domain,
        club_starring_df=club_star_df,
        all_club_counts=all_club_counts
    )

    # Tabs for the 3 main rule components
    is_womens = "women" in domain.lower()
    rule_code = "Rule WA10" if is_womens else "Rule A10"
    tab_a10, tab_a11, tab_a13 = st.tabs([
        f"📋 {rule_code}: Quota Checker",
        "⏱️ Rule A11 & A12: Absence & Eligibility",
        "🔄 Rule A13: Transfer Monitor & Finance Link"
    ])

    # ----------------------------------------------------
    # TAB 1: RULE A10 / WA10 AUTOMATIC QUOTA CHECKER
    # ----------------------------------------------------
    with tab_a10:
        st.subheader(f"📋 {rule_code}: Automatic Starring Quota Enforcer")
        if is_womens:
            st.markdown(
                "Enforces exact starring slots per tier based on the club's total team count according to **Rule WA10**. "
                "Clubs with 2 teams field 7 starred players for 1st XI; clubs with 3 teams field 7 for 1st XI and 9 for 2nd XI; "
                "clubs with 4 teams field 7 for 1st XI, 9 for 2nd XI, and 11 for 3rd XI."
            )
        else:
            st.markdown(
                "Enforces exact starring slots per tier based on the club's total team count. "
                "Clubs with 2 teams field 8 starred players for 1st XI; clubs with 3+ teams complete upper tiers and star downstream."
            )

        c_q1, c_q2 = st.columns([1.5, 3])
        with c_q1:
            st.metric(
                label="Total Senior Teams Fielded",
                value=f"{auto_num_teams} Teams",
                delta="Auto-Detected",
                delta_color="normal",
                help=f"Automatically calculated for {selected_club} from NCU League Structure and starring records."
            )
            min_override = 1 if is_womens else 2
            max_override = 4 if is_womens else 6
            val_override = max(min_override, min(int(auto_num_teams), max_override))
            with st.expander("⚙️ Override Team Count"):
                num_teams_override = st.number_input(
                    "Simulate / Override Count:",
                    min_value=min_override,
                    max_value=max_override,
                    value=val_override,
                    step=1,
                    key=f"a10_num_teams_{domain}_{selected_club}"
                )
            num_teams = num_teams_override
        with c_q2:
            if is_womens:
                if num_teams == 2:
                    tier_expl = "2 Teams: First 7 players normally selected for 1st XI (Rule WA10(a)). Automatic starring applies to NCU Senior / Future Series (inc. Ireland U17) reps (≥2 times in past 12 mos)."
                elif num_teams == 3:
                    tier_expl = "3 Teams: 1st XI: 7; 2nd XI: 4 remaining 1st XI + first 5 normally selected for 2nd XI = 9 players (Rule WA10(b))."
                elif num_teams >= 4:
                    tier_expl = "4 Teams: 1st XI: 7; 2nd XI: 9; 3rd XI: Next 6 from 2nd XI + first 5 normally selected for 3rd XI = 11 players (Rule WA10(c))."
                else:
                    tier_expl = "1 Team: No starring quotas required."
                st.info(f"**Rule WA10 Tier Rules for {num_teams} Teams:** {tier_expl}")
            else:
                st.info(
                    f"**Rule A10 Tier Rules for {num_teams} Teams:** "
                    + ("2 Teams: Exactly 8 for 1st XI." if num_teams == 2 else "")
                    + ("3 Teams: 3 remaining 1st XI + 7 for 2nd XI (1st XI: 8, 2nd XI: 10)." if num_teams == 3 else "")
                    + ("4 Teams: Next 3 from 2nd XI + 5 for 3rd XI (1st XI: 8, 2nd XI: 10, 3rd XI: 8)." if num_teams == 4 else "")
                    + ("5 Teams: Next 5 from 3rd XI + 2 for 4th XI (1st XI: 8, 2nd XI: 10, 3rd XI: 8, 4th XI: 7)." if num_teams == 5 else "")
                    + ("6 Teams: Next 7 from 4th XI for 5th XI (1st XI: 8, 2nd XI: 10, 3rd XI: 8, 4th XI: 7, 5th XI: 7)." if num_teams >= 6 else "")
                )

        quota_validation = sr.validate_club_starring_quotas(club_star_df, num_teams, domain=domain)

        # Display status indicators
        st.markdown("#### 🎯 Starring Slot Quota Fulfillment")
        q_cols = st.columns(len(quota_validation)) if quota_validation else [st.container()]
        all_complete = True
        for idx, (tier, q_data) in enumerate(quota_validation.items()):
            if not q_data["is_complete"]:
                all_complete = False
            with q_cols[idx]:
                delta_str = "Complete" if q_data["is_complete"] else (f"Short {abs(q_data['delta'])}" if q_data["delta"] < 0 else f"+{q_data['delta']} Over")
                delta_color = "normal" if q_data["is_complete"] else "inverse"
                st.metric(
                    label=f"{tier} Quota",
                    value=f"{q_data['actual']} / {q_data['expected']}",
                    delta=delta_str,
                    delta_color=delta_color
                )
                st.caption(f"Status: **{q_data['status_label']}**")

        if all_complete and quota_validation:
            st.success(f"✅ All starring quotas for **{selected_club}** are 100% compliant with {rule_code} ({sum(q['actual'] for q in quota_validation.values())} total starred players).")
        elif not all_complete:
            st.warning(f"⚠️ Quota discrepancy detected for **{selected_club}**. Adjust rosters to satisfy {rule_code} quotas.")

        st.markdown("#### 👥 Current Starred Roster Breakdown")
        if not club_star_df.empty:
            df_reg_all = eng.get_excel_df(f_reg) if f_reg and os.path.exists(f_reg) else pd.DataFrame()
            df_alias = eng.get_excel_df(f_alias) if f_alias and os.path.exists(f_alias) else pd.DataFrame()
            alias_map = eng.build_alias_map(df_alias, domain) if not df_alias.empty else {}

            starred_display_df = sr.check_starred_roster_registration(club_star_df, df_reg_all, alias_map)
            starred_display_df = sr.sort_starring_roster_dataframe(starred_display_df)
            disp_star_cols = [c for c in ["Rank", "Surname", "Forename", "XI_Level", "Full Name", "Registered"] if c in starred_display_df.columns]

            total_starred = len(starred_display_df)
            reg_count = sum(1 for v in starred_display_df["Registered"] if v in ["✅", "✔️"])
            unreg_count = total_starred - reg_count

            if unreg_count > 0:
                st.warning(f"⚠️ **Registration Notice:** {unreg_count} of {total_starred} starred players for **{selected_club}** are not currently registered with NCU.")
            else:
                st.success(f"✅ **Registration Verified:** All {total_starred} starred players for **{selected_club}** are officially registered.")

            st.dataframe(
                starred_display_df[disp_star_cols],
                width="stretch",
                hide_index=True,
                column_config=get_standard_column_config()
            )
        else:
            st.info(f"No starred roster records loaded for {selected_club}.")

        st.divider()
        if "starring_save_success" in st.session_state:
            st.success(st.session_state.pop("starring_save_success"))
        if "starring_save_warning" in st.session_state:
            st.warning(st.session_state.pop("starring_save_warning"))

        st.markdown("#### 🔄 Roster Modification & Player Replacement")

        is_locked_now = sr.is_roster_modification_locked()
        if is_locked_now:
            st.warning(
                "⚠️ **Post-July 31st Roster Modification Notice:** Standard roster modifications are locked per NCU Rule A12."
            )
            enable_override = st.checkbox(
                "🔓 Enable Roster Edits (Emergency / Injury Adjustment)",
                value=False,
                key=f"a10_enable_override_{domain}_{selected_club}"
            )
        else:
            enable_override = True

        can_edit, edit_msg = sr.can_edit_roster(admin_override=enable_override)

        if not can_edit:
            st.info("🔒 Roster modification controls are currently locked. Check '🔓 Enable Roster Edits (Emergency / Injury Adjustment)' above to activate player replacement inputs.")
        else:
            if is_locked_now:
                st.info("🔓 **Administrative Override Active:** Player replacement inputs and save controls unlocked for emergency / injury adjustments.")

            rc1, rc2, rc3 = st.columns([1.5, 2, 2])
            with rc1:
                replace_tier = st.selectbox(
                    "Team Tier:",
                    options=list(quota_validation.keys()) if quota_validation else ["1st XI", "2nd XI", "3rd XI", "4th XI", "5th XI"],
                    key=f"a10_replace_tier_{domain}_{selected_club}"
                )
            with rc2:
                tier_players = []
                if not club_star_df.empty and "XI_Level" in club_star_df.columns:
                    tier_sub = club_star_df[club_star_df["XI_Level"].astype(str).str.strip().str.lower() == replace_tier.lower()]
                    tier_sub_sorted = sr.sort_starring_roster_dataframe(tier_sub)
                    tier_players = tier_sub_sorted["Full Name"].dropna().tolist()

                # Check if this tier has an open vacancy under Rule A10 quotas
                tier_quota = quota_validation.get(replace_tier, {}) if quota_validation else {}
                has_vacancy = tier_quota.get("actual", 0) < tier_quota.get("expected", 0)
                out_options = (["[Vacant Slot]"] + tier_players) if has_vacancy else tier_players

                if out_options:
                    player_out = st.selectbox(
                        "Outgoing / De-Starred Player:",
                        options=out_options,
                        key=f"a10_player_out_{domain}_{selected_club}"
                    )
                else:
                    player_out = st.text_input(
                        "Outgoing / De-Starred Player:",
                        key=f"a10_player_out_txt_{domain}_{selected_club}"
                    )

            # Build current starred tier map for the selected club (including alias variants)
            starred_tier_map = {}
            if not club_star_df.empty and "Full Name" in club_star_df.columns and "XI_Level" in club_star_df.columns:
                for _, s_row in club_star_df.iterrows():
                    s_fn = str(s_row.get("Full Name", "")).strip()
                    s_tier = str(s_row.get("XI_Level", "")).strip()
                    if s_fn:
                        starred_tier_map[s_fn.lower()] = s_tier
                        if 'df_alias' in locals() and df_alias is not None and not df_alias.empty:
                            col_s80 = next((c for c in ["Official Registered Name", "Registered Name"] if c in df_alias.columns), None)
                            col_nv = next((c for c in ["NV Play Name", "Input Name (Scorecard/Stats)"] if c in df_alias.columns), None)
                            if col_s80 and col_nv:
                                match_rows = df_alias[(df_alias[col_nv].astype(str).str.strip().str.lower() == s_fn.lower()) | (df_alias[col_s80].astype(str).str.strip().str.lower() == s_fn.lower())]
                                for _, m_r in match_rows.iterrows():
                                    m_s80 = str(m_r.get(col_s80, "")).strip().lower()
                                    m_nv = str(m_r.get(col_nv, "")).strip().lower()
                                    if m_s80:
                                        starred_tier_map[m_s80] = s_tier
                                    if m_nv:
                                        starred_tier_map[m_nv] = s_tier

            # Retrieve candidate registered club players
            df_reg_lookup = df_reg_all if 'df_reg_all' in locals() and not df_reg_all.empty else (eng.get_excel_df(f_reg) if f_reg and os.path.exists(f_reg) else pd.DataFrame())
            df_alias_lookup = df_alias if 'df_alias' in locals() and not df_alias.empty else (eng.get_excel_df(f_alias) if f_alias and os.path.exists(f_alias) else None)
            candidate_players = sr.get_club_registered_players_list(
                df_reg=df_reg_lookup,
                club_name=selected_club,
                clean_club_fn=eng.club_matches_team_base,
                club_star_df=club_star_df,
                df_alias=df_alias_lookup
            )

            with rc3:
                if candidate_players:
                    def format_player_in_opt(p_name: str) -> str:
                        curr_tier = starred_tier_map.get(p_name.strip().lower())
                        if curr_tier:
                            return f"{p_name} ⭐ (Starred: {curr_tier})"
                        return p_name

                    player_in = st.selectbox(
                        "Incoming Replacement Player:",
                        options=candidate_players,
                        format_func=format_player_in_opt,
                        key=f"a10_player_in_select_{domain}_{selected_club}"
                    )
                else:
                    player_in = st.text_input(
                        "Incoming Replacement Player:",
                        placeholder="e.g. John Smith",
                        key=f"a10_player_in_txt_{domain}_{selected_club}"
                    )

            # Check if selected incoming player is currently starred in another tier
            prev_tier = starred_tier_map.get(player_in.strip().lower()) if player_in else None
            if prev_tier:
                if prev_tier.lower() == replace_tier.lower():
                    st.warning(f"⚠️ **Note:** **{player_in}** is already currently starred for **{replace_tier}**.")
                else:
                    st.info(
                        f"ℹ️ **Promoted / Re-Starred Player:** **{player_in}** is currently starred in the **{prev_tier}**. "
                        f"Assigning them to **{replace_tier}** will promote them and leave a vacancy in **{prev_tier}** that will need to be filled."
                    )

            if is_locked_now:
                admin_reason = st.text_input(
                    "Mandatory Override Reason / Explanation (required for league records):",
                    placeholder="e.g., Season-ending injury replacement approved by NCU",
                    key=f"a10_override_reason_{domain}_{selected_club}"
                )
            else:
                admin_reason = "Standard pre-deadline starring roster modification"

            if st.button("💾 Save Starring Adjustment & Update Audit Log", width="stretch", key=f"a10_save_override_{domain}_{selected_club}"):
                if not player_in.strip():
                    st.error("⚠️ Please select or enter a valid incoming replacement player name.")
                elif not player_out.strip():
                    st.error("⚠️ Please select or enter the outgoing player.")
                elif player_in.strip().lower() == player_out.strip().lower() and not any(k in player_out.lower() for k in ["vacant", "empty"]):
                    st.error("⚠️ Incoming replacement player cannot be the same as the outgoing player.")
                elif is_locked_now and not admin_reason.strip():
                    st.error("⚠️ An explanation comment is mandatory for post-July 31st administrative overrides.")
                else:
                    try:
                        # 1. Update the master Excel workbook on disk
                        sr.update_master_starring_roster(
                            file_path=f_starring,
                            club_name=selected_club,
                            tier=replace_tier,
                            player_out=player_out,
                            player_in=player_in,
                            prev_tier=prev_tier if prev_tier and prev_tier.lower() != replace_tier.lower() else None
                        )

                        # 2. Log override in NCU_Club_Starring_History.xlsx
                        audit_record = sr.log_starring_override_change(
                            club=selected_club,
                            tier=replace_tier,
                            player_out=player_out,
                            player_in=player_in,
                            admin_comment=admin_reason,
                            prev_tier=prev_tier if prev_tier and prev_tier.lower() != replace_tier.lower() else None
                        )

                        # 3. Store flash messages in session state across rerun
                        f_name = os.path.basename(f_starring) if f_starring else "Master Starring Excel"
                        if any(k in player_out.lower() for k in ["vacant", "empty"]):
                            action_desc = f"Assigned **{player_in}** into open vacant slot in **{replace_tier}**."
                        else:
                            action_desc = f"Replaced **{player_out}** with **{player_in}** in **{replace_tier}**."

                        st.session_state["starring_save_success"] = (
                            f"✅ **Starring Adjustment Saved & Roster Updated!** {action_desc}\n\n"
                            f"Master workbook (`{f_name}`) and audit trail (`NCU_Club_Starring_History.xlsx`) updated successfully at {audit_record['Timestamp']}."
                        )
                        if prev_tier and prev_tier.lower() != replace_tier.lower():
                            st.session_state["starring_save_warning"] = (
                                f"🚨 **Starring Slot Vacancy Alert:** **{player_in}** was moved from **{prev_tier}** into **{replace_tier}**. "
                                f"To maintain Rule A10 starring quotas, please assign a replacement player into **{prev_tier}**."
                            )

                        # 4. Trigger Streamlit rerun so that freshly modified file is re-read and page displays the new player immediately
                        st.rerun()
                    except Exception as ex:
                        st.error(f"Error logging override change: {ex}")

        starring_history_path = resolve_starring_history_path()
        if starring_history_path and os.path.exists(starring_history_path):
            with st.expander("📜 View Audit Trail & Config (NCU_Club_Starring_History.xlsx)"):
                try:
                    df_audit_view = eng.read_excel_calamine(starring_history_path, sheet_name="Starring Audit Trail")
                    st.markdown("**Starring Audit Trail:**")
                    st.dataframe(df_audit_view, width="stretch", hide_index=True)
                except Exception:
                    pass
                try:
                    df_intl_view = eng.read_excel_calamine(starring_history_path, sheet_name="International Exemptions")
                    st.markdown("**International Duty Exemptions:**")
                    st.dataframe(df_intl_view, width="stretch", hide_index=True)
                except Exception:
                    pass

    # ----------------------------------------------------
    # TAB 2: RULE A11 & A12 ABSENCE TRACKER & DEADLINE LOCK
    # ----------------------------------------------------
    with tab_a11:
        st.subheader("⏱️ Rule A11 & A12: Absence Tracker & Eligibility Monitor")
        st.markdown(
            "Tracks player appearances across scorecard match logs. "
            "Flags **⚠️ De-Starring Action Required** if a starred player does not appear for their team or a higher one "
            "for **3 consecutive matches or 3 weeks** (whichever is greater). "
            "International duty overrides this absence constraint."
        )

        # Hard Deadline Lock Banner
        is_locked_now = sr.is_roster_modification_locked()
        if is_locked_now:
            st.warning("⚠️ **Hard Deadline Lock Active (NCU Rule A12):** Standard roster modifications and starring changes are locked after 31st July.")
            override_t2 = st.checkbox(
                "🔓 Enable Roster Edits (Emergency / Injury Adjustment)",
                value=False,
                key="a12_admin_override_t2"
            )
            if override_t2:
                st.info("🔓 **Administrative Override Enabled:** Roster modification controls are unlocked in Rule A10 tab.")
        else:
            st.success("🔓 **Roster Modification Window Open:** Standard starring adjustments permitted prior to 31st July.")

        col_cfg1, col_cfg2 = st.columns([1, 2])
        with col_cfg1:
            eval_date_input = st.date_input("Audit / Evaluation Date:", value=datetime.now().date(), key="a11_eval_date")
            eval_date = datetime.combine(eval_date_input, datetime.min.time())
        with col_cfg2:
            # 1. Dynamically populate with whole club registration list (including starred players), sorted alphabetically (Surname, First Name)
            df_reg_lookup = df_reg_all if 'df_reg_all' in locals() and not df_reg_all.empty else (eng.get_excel_df(f_reg) if f_reg and os.path.exists(f_reg) else pd.DataFrame())
            df_alias_lookup = df_alias if 'df_alias' in locals() and not df_alias.empty else (eng.get_excel_df(f_alias) if f_alias and os.path.exists(f_alias) else None)
            club_candidate_players = sr.get_club_registered_players_list(
                df_reg=df_reg_lookup,
                club_name=selected_club,
                clean_club_fn=eng.club_matches_team_base,
                club_star_df=club_star_df,
                df_alias=df_alias_lookup
            )
            if not club_candidate_players:
                club_candidate_players = sr.get_club_starred_roster_sorted(club_star_df)

            # 2. Transactional persistence: Load saved exemptions from NCU_Club_Starring_History.xlsx
            persisted_exemptions = sr.load_international_exemptions(selected_club, domain=domain)
            if persisted_exemptions is None:
                default_intl_pool = ["Mark Adair", "Paul Stirling"] if domain == "Men's" else ["Cara Murray"]
                initial_exempt = [
                    p for p in club_candidate_players
                    if any(d.lower() == p.lower() or d.lower() in p.lower() for d in default_intl_pool)
                ]
                sr.save_international_exemptions(selected_club, initial_exempt, domain=domain)
                persisted_exemptions = initial_exempt

            # Ensure any persisted exemptions are preserved in available options
            dropdown_options = list(club_candidate_players)
            if persisted_exemptions:
                for pe in persisted_exemptions:
                    if pe and pe not in dropdown_options:
                        dropdown_options.append(pe)
                dropdown_options = sorted(
                    dropdown_options,
                    key=lambda x: (x.strip().split()[-1].lower() if x.strip().split() else "", x.strip().lower())
                )

            active_defaults = [p for p in (persisted_exemptions or []) if p in dropdown_options]

            selected_intl_players = st.multiselect(
                "Assign International Duty Exemptions (Club-Specific)",
                options=dropdown_options,
                default=active_defaults,
                key=f"a11_intl_exempt_{domain}_{selected_club}",
                help="Assign individual international duty exemptions for players representing Ireland. Exempts players from Rule A11/A12 de-starring absence thresholds."
            )

            # 3. Transactional persistence: Save dynamically if modified in multiselect widget
            if set(selected_intl_players) != set(active_defaults):
                sr.save_international_exemptions(selected_club, selected_intl_players, domain=domain)

        # Evaluate absences for players in selected club via proven engine pipeline
        df_absence = pd.DataFrame()
        if not club_star_df.empty:
            f_abandoned = c_files.get("abandoned", "")
            f_id_map = c_files.get("id_map", "")
            mtimes = tuple(
                os.path.getmtime(p) for p in [f_reg, f_alias, f_bat, f_bowl, f_abandoned, f_id_map] if p and os.path.exists(p)
            )
            all_app, override_map, comp_map, get_official_name, _, _, _ = cached_starring_pipeline(
                domain, f_reg, f_alias, f_bat, f_bowl, f_abandoned, f_id_map, None, None, mtimes
            )

            active_intl_list = [p.strip().lower() for p in selected_intl_players if str(p).strip()]

            df_absence = eng.evaluate_club_starring_inactivity(
                club_name=selected_club,
                club_star_df=club_star_df,
                all_app=all_app,
                override_map=override_map,
                comp_map=comp_map,
                get_official_name_func=get_official_name,
                international_players=active_intl_list,
                eval_date=eval_date,
                is_intl_override=False
            )

        # Summary KPIs
        k1, k2, k3, k4 = st.columns(4)
        total_p = len(df_absence)
        flagged_p = len(df_absence[df_absence["Administrative Status"].str.startswith("⚠️")]) if not df_absence.empty else 0
        intl_p = len(df_absence[df_absence["Administrative Status"].str.contains("International", na=False)]) if not df_absence.empty else 0
        eligible_p = total_p - flagged_p

        with k1: st.metric("Total Starred Players", total_p)
        with k2: st.metric("Eligible Players", eligible_p)
        with k3: st.metric("⚠️ De-Starring Required", flagged_p, delta=f"{flagged_p} Action Items" if flagged_p else "None", delta_color="inverse" if flagged_p else "normal")
        with k4: st.metric("International Exemptions", intl_p)

        if not df_absence.empty:
            df_absence_sorted = sr.sort_starring_roster_dataframe(df_absence)
            st.dataframe(
                df_absence_sorted,
                width="stretch",
                hide_index=True,
                column_config=get_standard_column_config()
            )
        else:
            st.info("No player absence records to display.")

    # ----------------------------------------------------
    # TAB 3: RULE A13 TRANSFER MONITOR & FINANCE LINK
    # ----------------------------------------------------
    with tab_a13:
        st.subheader("🔄 Rule A13: Seasonal Transfer Monitor & Financial Penalty Bridge")
        st.markdown(
            "Monitors seasonal player movements per **NCU Rule A13**:\n"
            "- **Transfer Limit:** Maximum of 2 transfers allowed per player per season. Any 3rd+ transfer attempt is strictly blocked.\n"
            "- **Financial Penalty:** For any valid transfer occurring **on or after 1st April**, an automatic **£25.00 Transfer Fee infraction** is assessed to the destination club and piped into `finance_app.py`."
        )

        df_reg_all = eng.get_excel_df(f_reg) if f_reg and os.path.exists(f_reg) else pd.DataFrame()
        transfer_audit = sr.evaluate_player_transfers(df_reg_all, season_year=2026, clean_club_fn=eng.extract_base_club_name)

        df_valid_trans = transfer_audit.get("valid_transfers", pd.DataFrame())
        df_blocked_trans = transfer_audit.get("blocked_transfers", pd.DataFrame())
        total_fees = transfer_audit.get("total_fees_assessed", 0.0)

        # KPI row
        m1, m2, m3, m4 = st.columns(4)
        tot_trans = len(df_valid_trans)
        fee_trans = len(df_valid_trans[df_valid_trans["Fee Infraction"]]) if not df_valid_trans.empty else 0
        tot_blocked = len(df_blocked_trans)

        with m1: st.metric("Total Valid Transfers", tot_trans)
        with m2: st.metric("In-Season Transfers (>= 1 Apr)", fee_trans)
        with m3: st.metric("Total Transfer Fees Due", f"£{total_fees:,.2f}", help="Assessed at £25.00 per late transfer on or after 1st April")
        with m4: st.metric("Rule A13 Blocked Attempts", tot_blocked, delta=f"{tot_blocked} Blocked" if tot_blocked else "None", delta_color="inverse" if tot_blocked else "normal")

        if not df_blocked_trans.empty:
            st.error("🚨 **Rule A13 Transfer Limit Violations Detected:** The following transfers exceed the maximum 2 transfers per season limit and are blocked:")
            df_blocked_sorted = sr.sort_transfer_records_dataframe(df_blocked_trans)
            st.dataframe(
                df_blocked_sorted,
                width="stretch",
                hide_index=True,
                column_config=get_standard_column_config()
            )

        st.markdown("#### 🔄 Seasonal Player Transfers")
        col_vt1, _ = st.columns([3, 1])
        with col_vt1:
            trans_scope = st.radio(
                "Transfer View Scope:",
                options=[f"🎯 {selected_club} Transfers", "🌐 All NCU Transfers (League-Wide)"],
                horizontal=True,
                key="a13_trans_scope"
            )

        def is_club_match(c_val: Any) -> bool:
            if not c_val or pd.isna(c_val):
                return False
            return eng.club_matches_team_base(selected_club, str(c_val))

        if trans_scope.startswith("🎯"):
            filtered_trans = df_valid_trans[
                df_valid_trans["To Club"].apply(is_club_match) | df_valid_trans["From Club"].apply(is_club_match)
            ].copy() if not df_valid_trans.empty else pd.DataFrame()
            empty_msg = f"No player transfers recorded involving **{selected_club}** for the 2026 season. Toggle **'All NCU Transfers'** above to review all {tot_trans} league-wide transfers."
        else:
            filtered_trans = df_valid_trans.copy()
            empty_msg = "No valid transfer records found."

        if not filtered_trans.empty:
            filtered_trans = sr.sort_transfer_records_dataframe(filtered_trans)
            disp_trans_cols = [c for c in ["Player", "From Club", "To Club", "Transfer Number", "Transfer Date", "Fee Due (£)", "Fee Infraction"] if c in filtered_trans.columns]
            st.dataframe(
                filtered_trans[disp_trans_cols],
                width="stretch",
                hide_index=True,
                column_config=get_standard_column_config()
            )
        else:
            st.info(empty_msg)


        st.divider()
        st.markdown("#### 💳 Dynamic Integration with Finance Command Center (`finance_app.py`)")
        st.success("🔗 **Active Ledger Bridge:** All £25.00 transfer fees are dynamically linked with Sharon's official invoicing schedule matrix and itemized on club invoices.")

# ==========================================
# TOOL 4: STARRING & INACTIVITY REPORTS
# ==========================================
elif app_mode == "Starring & Inactivity Reports":
    st.title(PAGE_TITLES["starring_reports"])
    st.markdown("Generate club-by-club Excel files highlighting inactive starred players (Red/Yellow) and tracking international exemptions (Green).")
    
    st.subheader("Select League Domain")
    domain = st.radio("Choose the dataset domain to audit:", ["Men's", "Women's"], horizontal=True)

    include_irish = False
    if domain == "Men's":
        include_irish = st.toggle("Include Irish Competitions in Inactivity Reports?", value=False, key="star_include_irish")

    with st.sidebar:
        c_files = eng.DEFAULT_FILES[domain]
        with st.expander("📁 File Path Configurations", expanded=False):
            f_reg = st.text_input("Official Registry (Excel)", value=c_files["reg"], key=f"star_reg_{domain}")
            f_alias = st.text_input("Aliases Master (Excel)", value=c_files["alias"], key=f"star_alias_{domain}")
            f_starring = st.text_input("Starring Master (Excel)", value=c_files["starring"], key=f"star_starring_{domain}")
            f_bat = st.text_input("Batting Stats (Excel)", value=c_files["bat"], key=f"star_bat_{domain}")
            f_bowl = st.text_input("Bowling Stats (Excel)", value=c_files["bowl"], key=f"star_bowl_{domain}")
            f_abandoned = st.text_input("Abandoned Games Stats (Excel)", value=c_files.get("abandoned", ""), key=f"star_ab_{domain}")
            if domain == "Men's" and include_irish:
                f_irish_bat = st.text_input("Irish Batting Stats (Excel)", value=c_files.get("irish_bat", eng.DEFAULT_IRISH_BAT_FILE), key="star_irish_bat")
                f_irish_bowl = st.text_input("Irish Bowling Stats (Excel)", value=c_files.get("irish_bowl", eng.DEFAULT_IRISH_BOWL_FILE), key="star_irish_bowl")

    st.subheader("Generate Reports")
    if st.button("📦 Process All Clubs & Download ZIP", type="primary"):
        files_to_check = [f for f in [f_reg, f_alias, f_starring, f_bat, f_bowl] if f]
        if f_abandoned: files_to_check.append(f_abandoned)
        if domain == "Men's" and include_irish: files_to_check.extend([f_irish_bat, f_irish_bowl])
        missing_files = [f for f in files_to_check if not os.path.exists(f)]
        
        if missing_files:
            st.error(f"Cannot find the following files:\n\n" + "\n".join([f"- {f}" for f in missing_files]))
        else:
            with st.spinner(f"Generating {domain} Starring Reports for all clubs..."):
                try:
                    irish_bat_path = f_irish_bat if (domain == "Men's" and include_irish) else None
                    irish_bowl_path = f_irish_bowl if (domain == "Men's" and include_irish) else None
                    
                    zip_buffer = eng.generate_starring_inactivity_reports(
                        domain=domain, f_reg=f_reg, f_alias=f_alias, f_starring=f_starring,
                        f_bat=f_bat, f_bowl=f_bowl, f_irish_bat=irish_bat_path, f_irish_bowl=irish_bowl_path, f_abandoned=f_abandoned
                    )
                    zip_buffer.seek(0)
                    with zipfile.ZipFile(zip_buffer, 'r') as z_file:
                        name_list = z_file.namelist()
                        workbooks_count = sum(1 for item in name_list if item.startswith("NCU_Master_Audit_"))
                        has_unreg = "Unregistered_Starred_Players.xlsx" in name_list

                    st.success("✅ Reports generated successfully!")
                    st.subheader("📊 Exporter Output Summary")
                    col_star1, col_star2 = st.columns(2)
                    with col_star1: st.metric(label="Clubs Workbooks Created", value=workbooks_count)
                    with col_star2: st.metric(label="Flagged Unregistered Starred Players List", value="Yes" if has_unreg else "No")
                    
                    st.divider()
                    st.download_button("📥 Download Club Reports (ZIP)", data=zip_buffer.getvalue(), file_name=f"{domain.replace('''s''', '')}_Starring_Reports_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip", mime="application/zip", type="primary")
                except Exception as e:
                    st.error(f"An error occurred during processing: {str(e)}")

# ==========================================
# TOOL 6: CLUB CONTACTS & OFFICIALS DIRECTORY
# ==========================================
elif app_mode == "Club Contacts Directory":
    st.title(PAGE_TITLES["club_contacts"])
    st.markdown("Search and filter official union contacts by Club, Team Tier, or Union-Wide Role with one-tap calling and messaging.")

    with st.sidebar:
        with st.expander("📁 Contacts Master Configuration", expanded=False):
            f_contacts = st.text_input("Contacts Excel File", value=eng.DEFAULT_CONTACTS_FILE, key="contacts_filepath")

    if not os.path.exists(f_contacts):
        st.warning(f"⚠️ Contacts spreadsheet `{f_contacts}` was not found in the root directory. Please upload it below.")
        uploaded_contacts = st.file_uploader("Upload Club Contacts Excel File", type=["xlsx", "xls"], key="contacts_upload_manual")
        if uploaded_contacts:
            df_contacts, club_grounds, ordered_roles = cached_parse_uploaded_contacts(uploaded_contacts.getvalue())
        else:
            df_contacts, club_grounds, ordered_roles = pd.DataFrame(), {}, []
    else:
        df_contacts, club_grounds, ordered_roles = get_club_contacts_data(f_contacts)

    if df_contacts.empty:
        st.info("Please provide a valid `2026 Season Club Contacts.xlsx` file to load directory contacts.")
    else:
        tab_team_view, tab_role_view, tab_search = st.tabs([
            "🏏 Team & Club Filter",
            "👥 Union-Wide Role View",
            "🔍 Global Directory Search"
        ])

        contact_clubs = [c for c in df_contacts['Club'].unique() if c and str(c).lower() != 'nan']
        all_clubs = sorted(list(set(contact_clubs) | set(eng.NCU_ALL_CLUBS))) if contact_clubs else sorted(list(eng.NCU_ALL_CLUBS))
        present_tiers = [t for t in eng.NCU_CONTACT_TIER_HIERARCHY if t == "All Roles & Officials" or t in df_contacts['Team Tier'].unique()]

        # ----------------------------------------------------
        # TAB 1: TEAM-LEVEL & CLUB FILTERS
        # ----------------------------------------------------
        with tab_team_view:
            st.subheader("Club & Team Tier Lookups")
            col_c1, col_c2 = st.columns([1.5, 1.5])
            
            with col_c1:
                selected_club = st.selectbox("Select Club:", options=all_clubs, index=0 if all_clubs else None)
            with col_c2:
                selected_tier = st.selectbox("Select Scope / Team Tier:", options=present_tiers)

            if selected_club:
                grounds = club_grounds.get(selected_club, {})
                if grounds:
                    with st.expander(f"📍 {selected_club} Ground Locations", expanded=False):
                        for g_label, g_val in grounds.items():
                            st.markdown(f"**{g_label}:** {g_val}")

                club_matches = df_contacts[df_contacts['Club'] == selected_club].sort_values(by='Role Order')
                st.divider()
                st.subheader(f"📌 {selected_club} — {selected_tier}")

                if selected_tier != "All Roles & Officials":
                    filtered_view = club_matches[club_matches['Team Tier'] == selected_tier]
                    if filtered_view.empty:
                        st.info(f"No contact records found for {selected_club} under {selected_tier}.")
                    else:
                        render_contact_grid(filtered_view.sort_values(by='Role Order'))
                else:
                    if club_matches.empty:
                        st.info(f"No contact records found for {selected_club}.")
                    else:
                        matches_copy = club_matches.copy()
                        matches_copy['Group Order'] = matches_copy['Team Tier'].apply(lambda x: get_tier_group(x)[0])
                        matches_copy['Group Name'] = matches_copy['Team Tier'].apply(lambda x: get_tier_group(x)[1])

                        for (grp_order, grp_name), grp_df in matches_copy.groupby(['Group Order', 'Group Name'], sort=True):
                            st.markdown(f"#### {grp_name}")
                            render_contact_grid(grp_df.sort_values(by='Role Order'))
                            st.markdown("<br>", unsafe_allow_html=True)

        # ----------------------------------------------------
        # TAB 2: ROLE-SPECIFIC UNION-WIDE VIEWS
        # ----------------------------------------------------
        with tab_role_view:
            st.subheader("Union-Wide Role Directory")
            st.markdown("Select a role to view every assigned official across all union clubs in official spreadsheet sequence.")

            selected_role = st.selectbox(
                "Select Role to Inspect:", 
                options=ordered_roles, 
                index=0 if ordered_roles else None
            )

            if selected_role:
                role_df = df_contacts[df_contacts['Role'] == selected_role].copy()
                
                if role_df.empty:
                    st.info(f"No officials listed under: **{selected_role}**")
                else:
                    role_df['Direct Phone'] = role_df['Phone'].apply(format_tel_link)
                    role_df['Direct Email'] = role_df['Email'].apply(format_mail_link)
                    
                    display_table = role_df[['Club', 'Name', 'Direct Phone', 'Direct Email', 'Team Tier']].sort_values(by='Club')
                    st.markdown(
                        display_table.to_html(escape=False, index=False), 
                        unsafe_allow_html=True
                    )

        # ----------------------------------------------------
        # TAB 3: GLOBAL DIRECTORY SEARCH
        # ----------------------------------------------------
        with tab_search:
            st.subheader("Search Across Directory")
            q = st.text_input("Enter any name, club, phone number, or role keyword:", placeholder="e.g., Grounds, 07812, Cliftonville, Safeguarding")

            if q:
                clean_q = q.strip().lower()
                matched_rows = df_contacts[
                    df_contacts['Club'].str.lower().str.contains(clean_q, na=False) |
                    df_contacts['Role'].str.lower().str.contains(clean_q, na=False) |
                    df_contacts['Name'].str.lower().str.contains(clean_q, na=False) |
                    df_contacts['Phone'].str.lower().str.contains(clean_q, na=False) |
                    df_contacts['Email'].str.lower().str.contains(clean_q, na=False) |
                    df_contacts['Team Tier'].str.lower().str.contains(clean_q, na=False)
                ].sort_values(by=['Club', 'Role Order']).copy()

                if matched_rows.empty:
                    st.warning(f"No contact results found matching '{q}'.")
                else:
                    st.success(f"Found {len(matched_rows)} matching official(s):")
                    matched_rows['Direct Phone'] = matched_rows['Phone'].apply(format_tel_link)
                    matched_rows['Direct Email'] = matched_rows['Email'].apply(format_mail_link)
                    
                    st.markdown(
                        matched_rows[['Club', 'Role', 'Team Tier', 'Name', 'Direct Phone', 'Direct Email']].to_html(escape=False, index=False),
                        unsafe_allow_html=True
                    )
# ==========================================
# TOOL 7: NV PLAY CSV STATS IMPORTER
# ==========================================
elif app_mode == "CSV Match Stats Importer":
    st.title(PAGE_TITLES["csv_importer"])
    st.info(
        "💡 **Quick Match Importer:** Import NV Play `.csv` export files directly into active master season Excel workbooks. "
        "Automatically preserves cricket bowling figures (`@` text), player IDs, frozen headers (`A2`), and auto-fitted column widths."
    )

    st.subheader("1. Select Competition Domain")
    importer_domain = st.radio(
        "Target Dataset Domain:",
        ["Men's", "Women's", "Midweek"],
        horizontal=True,
        key="importer_domain_radio"
    )

    domain_files = eng.DEFAULT_FILES.get(importer_domain, {})
    target_bat_file = domain_files.get("bat", "")
    target_bowl_file = domain_files.get("bowl", "")

    with st.expander("📁 Target Master Excel Files (Configured)", expanded=False):
        c_bat, c_bowl = st.columns(2)
        with c_bat:
            st.markdown(f"**Batting Master:** `{target_bat_file}`")
            if os.path.exists(target_bat_file):
                st.caption("Status: ✅ Found on disk")
            else:
                st.caption("Status: ⚠️ File not found")
        with c_bowl:
            st.markdown(f"**Bowling Master:** `{target_bowl_file}`")
            if os.path.exists(target_bowl_file):
                st.caption("Status: ✅ Found on disk")
            else:
                st.caption("Status: ⚠️ File not found")

    st.subheader("2. Provide NV Play CSV Files")
    input_method = st.radio(
        "Source Method:",
        ["Upload CSV Files", "Scan Workspace for Recent CSVs"],
        horizontal=True,
        key="csv_input_method"
    )

    batting_file = None
    bowling_file = None

    if input_method == "Upload CSV Files":
        col1, col2 = st.columns(2)
        with col1:
            batting_file = st.file_uploader(
                "Upload Batting Stats CSV (`*-batting-stats-group-by-match.csv`)",
                type=["csv"],
                key="importer_upload_bat"
            )
        with col2:
            bowling_file = st.file_uploader(
                "Upload Bowling Stats CSV (`*-bowling-stats-group-by-match.csv`)",
                type=["csv"],
                key="importer_upload_bowl"
            )
    else:
        workspace_csvs = [f for f in os.listdir(".") if f.endswith(".csv")]
        bat_candidates = [f for f in workspace_csvs if "batting" in f.lower()]
        bowl_candidates = [f for f in workspace_csvs if "bowling" in f.lower()]

        col1, col2 = st.columns(2)
        with col1:
            if bat_candidates:
                bat_sel = st.selectbox(
                    "Select Batting CSV from Workspace:",
                    ["(None)"] + sorted(bat_candidates, reverse=True),
                    key="sel_bat_csv"
                )
                if bat_sel != "(None)":
                    batting_file = bat_sel
            else:
                st.info("No batting CSV files found in workspace root.")
        with col2:
            if bowl_candidates:
                bowl_sel = st.selectbox(
                    "Select Bowling CSV from Workspace:",
                    ["(None)"] + sorted(bowl_candidates, reverse=True),
                    key="sel_bowl_csv"
                )
                if bowl_sel != "(None)":
                    bowling_file = bowl_sel
            else:
                st.info("No bowling CSV files found in workspace root.")

    if batting_file or bowling_file:
        st.subheader("3. Pre-Flight Inspection & Duplicate Check")

        preview_data = []

        if batting_file:
            bat_info = eng.inspect_nv_play_csv(batting_file)
            bat_check = eng.check_csv_matches_against_excel(importer_domain, list(bat_info["groups"].keys()), "batting")
            for grp, count in bat_info["groups"].items():
                is_dup = grp in bat_check.get("existing_matches", [])
                preview_data.append({
                    "Type": "Batting",
                    "Match / Group": grp,
                    "Records": count,
                    "Status": "⚠️ Duplicate (Already in Excel)" if is_dup else "✅ New Match"
                })

        if bowling_file:
            bowl_info = eng.inspect_nv_play_csv(bowling_file)
            bowl_check = eng.check_csv_matches_against_excel(importer_domain, list(bowl_info["groups"].keys()), "bowling")
            for grp, count in bowl_info["groups"].items():
                is_dup = grp in bowl_check.get("existing_matches", [])
                preview_data.append({
                    "Type": "Bowling",
                    "Match / Group": grp,
                    "Records": count,
                    "Status": "⚠️ Duplicate (Already in Excel)" if is_dup else "✅ New Match"
                })

        if preview_data:
            df_preview = pd.DataFrame(preview_data)
            st.dataframe(df_preview, width="stretch", hide_index=True)

            has_duplicates = any("Duplicate" in r["Status"] for r in preview_data)
            allow_dups = False
            if has_duplicates:
                st.warning("⚠️ Some matches are already present in the target master workbook. By default, duplicate matches will be skipped to protect your data integrity.")
                allow_dups = st.checkbox("Force import duplicate matches anyway", value=False)

            st.subheader("4. Execute Import")
            if st.button("🚀 Import & Append to Season Master", type="primary", width="stretch"):
                with st.spinner("Importing records and enforcing Excel formatting rules..."):
                    result = eng.import_nv_play_stats(
                        domain=importer_domain,
                        batting_source=batting_file,
                        bowling_source=bowling_file,
                        allow_duplicates=allow_dups
                    )

                    if result["success"]:
                        st.success("✅ Import completed successfully!")
                        r_bat = result.get("batting_result")
                        r_bowl = result.get("bowling_result")

                        # Selectively update only the specific match data frame cache
                        bat_path = r_bat.get("excel_path") if r_bat else None
                        bowl_path = r_bowl.get("excel_path") if r_bowl else None
                        eng.update_match_data_cache(batting_path=bat_path, bowling_path=bowl_path)

                        # Reset session-level player search and audit state to force fresh reads of new matches
                        if "data_loaded" in st.session_state:
                            st.session_state.data_loaded = False
                        st.session_state.pop("audit_outputs", None)
                        st.session_state.pop("saved_paths", None)

                        m1, m2 = st.columns(2)
                        if r_bat:
                            with m1:
                                st.metric("Batting Rows Added", r_bat.get("rows_appended", 0))
                                if r_bat.get("matches_added"):
                                    st.write("**Matches added:**", r_bat["matches_added"])
                                if r_bat.get("matches_skipped"):
                                    st.write("**Duplicates skipped:**", r_bat["matches_skipped"])
                                if r_bat.get("backup_file"):
                                    st.caption(f"Backup created: `{r_bat['backup_file']}`")
                        if r_bowl:
                            with m2:
                                st.metric("Bowling Rows Added", r_bowl.get("rows_appended", 0))
                                if r_bowl.get("matches_added"):
                                    st.write("**Matches added:**", r_bowl["matches_added"])
                                if r_bowl.get("matches_skipped"):
                                    st.write("**Duplicates skipped:**", r_bowl["matches_skipped"])
                                if r_bowl.get("backup_file"):
                                    st.caption(f"Backup created: `{r_bowl['backup_file']}`")
                    else:
                        st.error("❌ An error occurred during import:")
                        for err in result.get("errors", []):
                            st.write(f"- {err}")

# ==========================================
# TOOL 8: PLAYER DISAMBIGUATION & PROFILE MAPPING
# ==========================================
elif app_mode == "Player Disambiguation & ID Mapping":
    st.title(PAGE_TITLES["player_disambiguation"])
    st.info(
        "💡 **Player Disambiguation & ID Mapping:** Detect active NV Play player UUIDs on match scorecards that lack a verified match "
        "in Sport80 registration files. Link profiles to permanently update master lookups or flag violations to assess fines."
    )

    # 1. State Structure Initialization
    if "unlinked_player_records" not in st.session_state:
        st.session_state["unlinked_player_records"] = {}
    if "fee_audit_violations" not in st.session_state:
        st.session_state["fee_audit_violations"] = []

    # 2. Controls & Sidebar Configurations
    col_d1, col_d2, col_d3 = st.columns([2, 1, 1])
    with col_d1:
        disambig_domain = st.radio(
            "Select Competition Domain:",
            ["Men's", "Women's", "Midweek"],
            horizontal=True,
            key="disambig_domain_radio"
        )
    with col_d2:
        st.write("")
        st.write("")
        scan_btn = st.button("🔄 Scan Scorecards", type="secondary", width="stretch")
    with col_d3:
        st.write("")
        st.write("")
        autolink_btn = st.button(
            "⚡ Auto-Link Verified",
            type="primary",
            width="stretch",
            help="Automatically links 100% exact name matches at the same club or verified transfer club"
        )

    c_files = eng.DEFAULT_FILES.get(disambig_domain, eng.DEFAULT_FILES["Men's"])
    with st.sidebar:
        st.subheader("⚙️ Disambiguation Settings")
        with st.expander("📁 Master File Configurations", expanded=False):
            f_reg = st.text_input("Official Registry (Excel)", value=c_files.get("reg", "1. NCU_Registered_Players.xlsx"), key=f"dis_reg_{disambig_domain}")
            f_id_map = st.text_input("ID Mapping Master (Excel)", value=c_files.get("id_map", "NCU_Mens_Master_ID_Mapping.xlsx"), key=f"dis_id_map_{disambig_domain}")
            f_alias = st.text_input("Aliases Master (Excel)", value=c_files.get("alias", "2. NCU_Validated_Aliases_Master.xlsx"), key=f"dis_alias_{disambig_domain}")
            f_bat = st.text_input("Batting Stats (Excel)", value=c_files.get("bat", ""), key=f"dis_bat_{disambig_domain}")
            f_bowl = st.text_input("Bowling Stats (Excel)", value=c_files.get("bowl", ""), key=f"dis_bowl_{disambig_domain}")
            f_ab = st.text_input("Abandoned Games Stats (Excel)", value=c_files.get("abandoned", ""), key=f"dis_ab_{disambig_domain}")
            f_unreg = st.text_input("Unregistered Manual Map (Excel)", value=c_files.get("unreg", "4. Unregistered_Manual_Map.xlsx"), key=f"dis_unreg_{disambig_domain}")

    # Auto-link action
    if autolink_btn:
        with st.spinner(f"Auto-linking high-confidence profiles for {disambig_domain}..."):
            auto_res = eng.auto_link_high_confidence_players(
                domain=disambig_domain,
                f_reg=f_reg,
                f_id_map=f_id_map,
                f_alias=f_alias,
                f_bat=f_bat,
                f_bowl=f_bowl,
                f_abandoned=f_ab,
                f_unreg=f_unreg,
            )
            scanned = eng.scan_unlinked_nvplay_players(
                domain=disambig_domain,
                f_reg=f_reg,
                f_id_map=f_id_map,
                f_bat=f_bat,
                f_bowl=f_bowl,
                f_abandoned=f_ab,
                f_unreg=f_unreg,
            )
            st.session_state["unlinked_player_records"][disambig_domain] = scanned
            st.success(f"⚡ Successfully auto-linked {auto_res['linked_count']} verified player profile(s)! {auto_res['remaining_count']} remaining.")

    # Auto-scan if domain not yet scanned or scan button pressed
    domain_records = st.session_state["unlinked_player_records"].get(disambig_domain, None)
    if domain_records is None or scan_btn:
        with st.spinner(f"Scanning {disambig_domain} scorecards for unlinked player UUIDs..."):
            scanned = eng.scan_unlinked_nvplay_players(
                domain=disambig_domain,
                f_reg=f_reg,
                f_id_map=f_id_map,
                f_bat=f_bat,
                f_bowl=f_bowl,
                f_abandoned=f_ab,
                f_unreg=f_unreg,
            )
            st.session_state["unlinked_player_records"][disambig_domain] = scanned
            domain_records = scanned

            for r in domain_records:
                if r.get('status') == 'flagged_unregistered':
                    if not any(v.get('NV_Play_ID') == r.get('nv_id') or v.get('Player') == r.get('nv_name') for v in st.session_state['fee_audit_violations']):
                        st.session_state['fee_audit_violations'].append({
                            "Player": r['nv_name'],
                            "Club": r['club'],
                            "NV_Play_ID": r.get('nv_id', ''),
                            "Matches": r.get('matches_played', 1),
                            "Match_Fixtures": r.get('match_groups', []),
                            "Fine": 10.0 * max(r.get('matches_played', 1), 1),
                            "Reason": "Fielding an unregistered player (Unmapped Profile)",
                            "Timestamp": "Persisted Audit Violation",
                            "Status": "Audit Violation Logged"
                        })

            if scan_btn:
                st.toast(f"Found {len(scanned)} unlinked player UUIDs across {disambig_domain} matches!", icon="🔍")

    reg_df = eng.get_excel_df(f_reg) if f_reg and os.path.exists(f_reg) else pd.DataFrame()

    tab_resolve, tab_verified, tab_violations = st.tabs([
        "⚠️ Unmapped Profile Queue",
        "✅ Verified & Linked Mappings",
        "🚨 Flagged Unregistered Violations"
    ])

    # ----------------------------------------------------
    # TAB 1: UNMAPPED PROFILE QUEUE (THE RESOLVE INTERFACE)
    # ----------------------------------------------------
    with tab_resolve:
        active_unlinked = [r for r in domain_records if r.get('status') == 'unlinked']
        linked_count = len([r for r in domain_records if r.get('status') == 'linked'])
        flagged_count = len([r for r in domain_records if r.get('status') == 'flagged_unregistered'])

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Unmapped Profiles", len(active_unlinked))
        m2.metric("Clubs Affected", len(set(r['club'] for r in active_unlinked)))
        m3.metric("Matches Affected", sum(r['matches_played'] for r in active_unlinked))
        m4.metric("Session Resolved", linked_count + flagged_count)

        st.divider()

        # Queue Filters
        col_f1, col_f2, col_f3 = st.columns([1.5, 2, 1])
        all_clubs = sorted(list(set(r['club'] for r in domain_records)))
        with col_f1:
            selected_club = st.selectbox(
                "Filter by Inferred Club:",
                options=["All Clubs"] + all_clubs,
                key=f"filter_club_{disambig_domain}"
            )
        with col_f2:
            search_query = st.text_input(
                "Search Player Name or UUID:",
                placeholder="Type name or UUID substring...",
                key=f"search_q_{disambig_domain}"
            )
        with col_f3:
            show_status = st.selectbox(
                "Queue Scope:",
                options=["Unmapped Only", "All Records", "Linked Only", "Flagged Unregistered"],
                key=f"scope_{disambig_domain}"
            )

        # Apply Filters
        filtered_records = domain_records
        if show_status == "Unmapped Only":
            filtered_records = [r for r in filtered_records if r.get('status') == 'unlinked']
        elif show_status == "Linked Only":
            filtered_records = [r for r in filtered_records if r.get('status') == 'linked']
        elif show_status == "Flagged Unregistered":
            filtered_records = [r for r in filtered_records if r.get('status') == 'flagged_unregistered']

        if selected_club != "All Clubs":
            filtered_records = [r for r in filtered_records if r['club'] == selected_club]

        if search_query.strip():
            sq = search_query.strip().lower()
            filtered_records = [
                r for r in filtered_records
                if sq in r['nv_name'].lower() or sq in r['nv_id'].lower() or sq in r['club'].lower()
            ]

        if not filtered_records:
            if not active_unlinked:
                st.success("🎉 All active NV Play player profiles for this domain are resolved and linked!")
            else:
                st.info("No records match the current filter criteria.")
        else:
            st.caption(f"Showing {len(filtered_records)} player profile(s) to resolve:")

            for record in filtered_records:
                is_resolved = record.get('status') in ['linked', 'flagged_unregistered']
                status_badge = "✅ Linked" if record.get('status') == 'linked' else ("🚨 Flagged Unregistered" if record.get('status') == 'flagged_unregistered' else "⚠️ Unmapped")

                with st.container(border=True):
                    col_left, col_right = st.columns([1, 1])

                    # ----------------------------------------------------
                    # LEFT PANEL: NV PLAY PROFILE DATA
                    # ----------------------------------------------------
                    with col_left:
                        st.markdown(f"### 🏏 {record['nv_name']}  `{status_badge}`")
                        st.markdown(f"**Club:** `{record['club']}`")
                        st.markdown(f"**NV Play UUID:** `{record['nv_id']}`")
                        st.markdown(f"**Matches Played:** **{record['matches_played']}** match(es)")

                        if record.get('match_groups'):
                            with st.expander(f"📋 View {len(record['match_groups'])} Fixture Appearances", expanded=False):
                                for fix in record['match_groups']:
                                    st.markdown(f"- `{fix}`")

                        if is_resolved:
                            if record.get('status') == 'linked':
                                st.success(f"Linked to: **{record.get('sport80_name', '')}** (ID: {record.get('sport80_id', '')}) — {record.get('sport80_club', '')}")
                            else:
                                st.warning("Definitive audit violation logged for fielding an unregistered player.")

                    # ----------------------------------------------------
                    # RIGHT PANEL: SPORT80 MATCHER & QUICK ACTIONS
                    # ----------------------------------------------------
                    with col_right:
                        st.markdown("### 🔍 Sport80 Profile Matcher")

                        if not is_resolved:
                            show_all_pool = st.checkbox(
                                "Show Union-Wide Pool (Ignore Club Filter)",
                                key=f"all_pool_{disambig_domain}_{record['nv_id']}",
                                value=False
                            )
                            target_club = "" if show_all_pool else record['club']
                            candidates = eng.get_sport80_candidates_for_club(reg_df, club=target_club)

                            options = ["-- Select Sport80 Registration --"]
                            opt_lookup = {}
                            auto_match_idx = 0

                            for idx, c in enumerate(candidates):
                                opt_text = c['display']
                                options.append(opt_text)
                                opt_lookup[opt_text] = c
                                if c['sport80_name'].lower() == record['nv_name'].lower():
                                    auto_match_idx = idx + 1

                            selected_sport80 = st.selectbox(
                                "Select Sport80 Player:",
                                options=options,
                                index=auto_match_idx if auto_match_idx < len(options) else 0,
                                key=f"sel_{disambig_domain}_{record['nv_id']}"
                            )

                            if auto_match_idx > 0 and selected_sport80 == options[auto_match_idx]:
                                st.caption("✨ *Auto-suggested based on matching registered name*")

                            st.write("")
                            btn_col1, btn_col2 = st.columns(2)

                            with btn_col1:
                                if st.button(
                                    "🔗 Link Profiles",
                                    key=f"btn_link_{disambig_domain}_{record['nv_id']}",
                                    type="primary",
                                    width="stretch"
                                ):
                                    if selected_sport80 == "-- Select Sport80 Registration --":
                                        st.error("Please select an official Sport80 profile from the dropdown before linking.")
                                    else:
                                        c_data = opt_lookup[selected_sport80]
                                        eng.save_nvplay_sport80_mapping(
                                            domain=disambig_domain,
                                            nv_id=record['nv_id'],
                                            nv_name=record['nv_name'],
                                            sport80_id=c_data['sport80_id'],
                                            sport80_name=c_data['sport80_name'],
                                            sport80_club=c_data['sport80_club'],
                                            f_id_map=f_id_map,
                                            f_alias=f_alias,
                                            f_unreg=f_unreg,
                                        )
                                        record['status'] = 'linked'
                                        record['sport80_id'] = c_data['sport80_id']
                                        record['sport80_name'] = c_data['sport80_name']
                                        record['sport80_club'] = c_data['sport80_club']
                                        st.success(f"✅ Successfully linked {record['nv_name']} to {c_data['sport80_name']} ({c_data['sport80_club']})!")
                                        st.rerun()

                            with btn_col2:
                                if st.button(
                                    "🚨 Flag as Unregistered",
                                    key=f"btn_flag_{disambig_domain}_{record['nv_id']}",
                                    type="secondary",
                                    width="stretch"
                                ):
                                    v_rec = eng.flag_player_as_unregistered(
                                        player_name=record['nv_name'],
                                        club=record['club'],
                                        nv_id=record['nv_id'],
                                        match_groups=record['match_groups'],
                                        f_unreg=f_unreg,
                                        f_id_map=f_id_map,
                                    )
                                    record['status'] = 'flagged_unregistered'
                                    record['flagged_unregistered'] = True
                                    st.session_state['fee_audit_violations'].append(v_rec)
                                    st.warning(f"⚠️ Flagged {record['nv_name']} as unregistered. £{v_rec['Fine']:.2f} penalty recorded.")
                                    st.rerun()
                        else:
                            st.info("This profile has already been processed in the current session.")
                            if st.button("↩️ Re-open for Editing", key=f"reopen_{disambig_domain}_{record['nv_id']}", width="stretch"):
                                record['status'] = 'unlinked'
                                st.rerun()

    # ----------------------------------------------------
    # TAB 2: VERIFIED & LINKED MAPPINGS
    # ----------------------------------------------------
    with tab_verified:
        st.subheader("📋 Master ID Mapping Directory")
        st.markdown("Verified pairing connections between NV Play UUIDs and Sport80 accounts stored in master Excel records.")

        if f_id_map and os.path.exists(f_id_map):
            df_map = eng.get_excel_df(f_id_map)
            if not df_map.empty:
                v_search = st.text_input("Filter Verified Mappings:", placeholder="Search by name, club, or UUID...", key=f"v_search_{disambig_domain}")
                if v_search.strip():
                    vs = v_search.strip().lower()
                    mask = (
                        df_map['NV_Play_Name'].astype(str).str.contains(vs, case=False, na=False) |
                        df_map['Sport80_Name'].astype(str).str.contains(vs, case=False, na=False) |
                        df_map['Sport80_Club'].astype(str).str.contains(vs, case=False, na=False) |
                        df_map['NV_Play_ID'].astype(str).str.contains(vs, case=False, na=False)
                    )
                    df_map = df_map[mask]

                st.dataframe(df_map, width="stretch", height=400)
                st.caption(f"Displaying {len(df_map)} verified mapping record(s).")
            else:
                st.info("Master ID Mapping workbook is empty.")
        else:
            st.warning(f"Master ID Mapping file `{f_id_map}` was not found.")

    # ----------------------------------------------------
    # TAB 3: FLAGGED UNREGISTERED VIOLATIONS & FEE DASHBOARD
    # ----------------------------------------------------
    with tab_violations:
        st.subheader("🚨 Unregistered Player Audit Violations & Financial Impact")
        st.markdown("Definitive audit violations logged against clubs for fielding unregistered players, automatically feeding into the fee reconciliation system.")

        session_violations = st.session_state.get('fee_audit_violations', [])
        f1, f2, f3 = st.columns(3)
        total_viol_fine = sum(float(v.get('Fine', 10.0)) for v in session_violations)
        f1.metric("Violations Logged (Session)", len(session_violations))
        f2.metric("Total Fee Penalties Assessed", f"£{total_viol_fine:.2f}")
        f3.metric("Clubs Impacted", len(set(v.get('Club', '') for v in session_violations)))

        if session_violations:
            st.write("### Session Logged Violations")
            df_s_viol = pd.DataFrame(session_violations)[['Player', 'Club', 'Matches', 'Fine', 'Timestamp', 'Status']]
            st.dataframe(df_s_viol, width="stretch")

        st.divider()
        st.write("### Unregistered Manual Map File (`4. Unregistered_Manual_Map.xlsx`)")
        if f_unreg and os.path.exists(f_unreg):
            df_unreg_map = eng.get_excel_df(f_unreg)
            if not df_unreg_map.empty:
                st.dataframe(df_unreg_map, width="stretch")
            else:
                st.info("Unregistered manual map workbook is empty.")
        else:
            st.caption(f"Manual map file `{f_unreg}` not yet created.")

