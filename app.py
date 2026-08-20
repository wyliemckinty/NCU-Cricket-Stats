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
import importlib

import engine as eng
importlib.reload(eng)  # Force Python to reload engine.py on every rerun

try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False
    
import warnings
warnings.filterwarnings('ignore')

# ==========================================
# UI CUSTOMIZATION SETTINGS
# ==========================================
# Change these values to adjust text sizes in the Contacts Directory. 
# You can use standard CSS sizes like "18px", "24px", "1.2rem", etc.
UI_CLUB_HEADER_SIZE = "22px"
UI_ROLE_TITLE_SIZE = "18px"
UI_OFFICIAL_NAME_SIZE = "16px"

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

# ==========================================
# USER CONFIGURATIONS & PERSISTENCE
# ==========================================
MAIN_HEADER_SIZE = "28px" 
CONFIG_FILE = "threshold_settings.json"

PAGE_TITLES = {
    "bulk_averages": "📊 League Bulk Averages Calculator",
    "player_doc": "📄 Player Word Doc Generator",
    "reg_checks": "🛡️ Weekend Registration and Starring Checks",
    "midweek_checks": "🛡️ Midweek Registration & Starring Check",
    "starring_reports": "🚨 Club Starring & Inactivity Exporter",
    "fines_generator": "💸 Club Fines Generator",
    "unregistered_fines": "💸 Unregistered Player Fines Generator",
    "milestones_report": "🏆 League Milestones Report",
    "club_contacts": "📇 Club Contacts & Officials Directory"
}

DEFAULT_THRESHOLDS = {
    "t1_runs": 200, "t1_bmat": 5, "t1_wick": 15, "t1_mmat": 5,
    "t2_runs": 150, "t2_bmat": 5, "t2_wick": 10, "t2_mmat": 5,
    "t3_runs": 100, "t3_bmat": 3, "t3_wick": 5,  "t3_mmat": 3,
    "t4_runs": 50,  "t4_bmat": 3, "t4_wick": 3,  "t4_mmat": 3,
    "w1_runs": 100, "w1_bmat": 5, "w1_wick": 10, "w1_mmat": 5,
    "w2_runs": 25,  "w2_bmat": 2, "w2_wick": 2,  "w2_mmat": 2,
    "mw_min_runs": 50, "mw_min_innings": 0, "mw_min_wickets": 5, "mw_min_bowl_innings": 0
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

st.markdown(f"""
<style>
    h1 {{ font-size: {MAIN_HEADER_SIZE} !important; font-weight: 700; }}
    div.stButton > button {{ white-space: nowrap !important; }}
    div.stButton > button[kind="primary"] {{ border-radius: 8px; padding: 0.5rem 1.5rem; }}
    div.stDownloadButton > button:first-child {{ background-color: #0066cc; color: white; border-radius: 8px; border: none; padding: 0.5rem 1.5rem; }}
    div.stDownloadButton > button:first-child:hover {{ background-color: #0052a3; color: white; }}
    [data-testid="stMetricValue"] {{ font-size: 1.8rem; font-weight: 700; }}
    [data-testid="metric-container"] {{ background-color: rgba(250, 250, 250, 0.1); border: 1px solid rgba(128, 128, 128, 0.2); padding: 15px; border-radius: 10px; }}
</style>
""", unsafe_allow_html=True)

# ==========================================
# SIDEBAR NAVIGATION
# ==========================================
with st.sidebar:
    st.title("🏏 NCU Cricket Hub")
    st.header("🛠️ Navigation")
    
    app_mode = st.radio(
        "Choose a module to run:",
        [
            "Player Word Doc Generator", 
            "Registration Checks",
            "Midweek Registration & Starring Check",
            "Starring & Inactivity Reports",
            "Club Fines Generator",
            "Unregistered Player Fines Generator",
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
            st.divider() 
            c_files = eng.DEFAULT_FILES[domain]
            with st.expander("📁 File Path Configurations", expanded=False):
                f_reg = st.text_input("Official Registry (Excel)", value=c_files["reg"], key=f"doc_reg_{domain}")
                f_alias = st.text_input("Aliases Master (Excel)", value=c_files["alias"], key=f"doc_alias_{domain}")
                f_bat = st.text_input("Batting Stats (Excel)", value=c_files["bat"], key=f"doc_bat_{domain}")
                f_bowl = st.text_input("Bowling Stats (Excel)", value=c_files["bowl"], key=f"doc_bowl_{domain}")
                f_abandoned = st.text_input("Abandoned Games Stats (Excel)", value=c_files.get("abandoned", ""), key=f"doc_ab_{domain}")
        
        include_irish = False
        if domain == "Men's":
            include_irish = st.toggle("Include Irish Competitions in Player Report?", value=False)
            if include_irish:
                with st.sidebar:
                    with st.expander("📁 Irish File Path Configurations", expanded=False):
                        f_irish_bat = st.text_input("Irish Batting Stats (Excel)", value="Irish Competitions 2026 Batting stats.xlsx", key="doc_irish_bat")
                        f_irish_bowl = st.text_input("Irish Bowling Stats (Excel)", value="Irish Competitions 2026 Bowling stats.xlsx", key="doc_irish_bowl")

        if 'doc_last_domain' not in st.session_state or st.session_state.doc_last_domain != domain:
            st.session_state.player_search_active = False
            st.session_state.doc_last_domain = domain

        with st.container(border=True):
            st.subheader("🔍 Search Player Database")
            search_query = st.text_input("Enter the player's full name or scorecard alias:", placeholder="e.g., Joe Bloggs")
            col_btn, _ = st.columns([1.5, 4])
            with col_btn:
                execute_search = st.button("🔍 Search Player", type="primary", use_container_width=True)

        if 'player_search_active' not in st.session_state:
            st.session_state.player_search_active = False

        if execute_search:
            st.session_state.player_search_active = True
            st.session_state.player_search_query = search_query
            st.session_state.data_loaded = False 

        if st.session_state.player_search_active:
            current_query = st.session_state.player_search_query
            files_to_check = [f_reg, f_alias, f_bat, f_bowl]
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
                        batting = get_excel_df(f_bat)
                        bowling = get_excel_df(f_bowl)
                        abandoned_df = get_excel_df(f_abandoned) if f_abandoned and os.path.exists(f_abandoned) else pd.DataFrame()
                        
                        if domain == "Men's" and include_irish:
                            if os.path.exists(f_irish_bat): batting = pd.concat([batting, get_excel_df(f_irish_bat)], ignore_index=True)
                            if os.path.exists(f_irish_bowl): bowling = pd.concat([bowling, get_excel_df(f_irish_bowl)], ignore_index=True)

                        alias_map = eng.build_alias_map(aliases, domain)
                        f_unreg = eng.DEFAULT_FILES.get(domain, {}).get("unreg", "")
                        unreg_df = get_excel_df(f_unreg) if os.path.exists(f_unreg) else None
                        player_club_map = eng.build_player_club_map(reg_players, alias_map, domain, unreg_map_df=unreg_df)
                        player_club_map = eng.infer_unregistered_player_clubs(batting, bowling, player_club_map, min_matches=2)
                        
                        def resolve_duplicates(row, name_col):
                            name = str(row[name_col])
                            row_team = str(row.get('Team', '')).lower()
                            match_grp = str(row.get('Group', row.get('Match', ''))).lower()
                            combined_context = row_team + ' ' + match_grp
                            if domain == "Men's" and name in eng.KNOWN_DUPLICATES:
                                for club in eng.KNOWN_DUPLICATES[name]:
                                    # Check all known aliases/abbreviations for this club
                                    # Use word-boundary regex to avoid 'CI' matching inside 'City' or 'CSNI'
                                    variants = eng.CLUB_ALIASES.get(club, [club])
                                    for variant in variants:
                                        if re.search(r'\b' + re.escape(variant.lower()) + r'\b', combined_context):
                                            return f"{name} ({club})"
                            return name
                        
                        batting['Name'] = batting.apply(lambda r: eng.cleanse_name_contextual(r['Name'], r, alias_map, player_club_map), axis=1)
                        bowling['Bowler'] = bowling.apply(lambda r: eng.cleanse_name_contextual(r['Bowler'], r, alias_map, player_club_map), axis=1)
                        if not abandoned_df.empty:
                            ab_name_col = 'Name' if 'Name' in abandoned_df.columns else abandoned_df.columns[1]
                            abandoned_df['Cleaned Name'] = abandoned_df.apply(lambda r: eng.cleanse_name_contextual(r[ab_name_col], r, alias_map, player_club_map), axis=1)
                            ab_grp_col = 'Group' if 'Group' in abandoned_df.columns else ('Match' if 'Match' in abandoned_df.columns else abandoned_df.columns[0])
                            abandoned_df['Group'] = abandoned_df[ab_grp_col].apply(lambda x: eng.doc_format_cricket_names(x, domain))

                        batting['Name'] = batting.apply(lambda x: resolve_duplicates(x, 'Name'), axis=1)
                        bowling['Bowler'] = bowling.apply(lambda x: resolve_duplicates(x, 'Bowler'), axis=1)
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

                        if '_computed_name' in reg_players.columns:
                            reg_matches = reg_players[reg_players['_computed_name'].astype(str).str.contains(clean_q, case=False, na=False)]
                            target_official_names.update(reg_matches['_computed_name'].dropna().astype(str).str.strip().tolist())
                        elif 'Full Name' in reg_players.columns:
                            reg_matches = reg_players[reg_players['Full Name'].astype(str).str.contains(clean_q, case=False, na=False)]
                            target_official_names.update(reg_matches['Full Name'].dropna().astype(str).str.strip().tolist())
                        elif 'First Name' in reg_players.columns and 'Last Name' in reg_players.columns:
                            comp = reg_players['First Name'].astype(str).str.strip() + ' ' + reg_players['Last Name'].astype(str).str.strip()
                            reg_matches = reg_players[comp.str.contains(clean_q, case=False, na=False)]
                            target_official_names.update(comp[reg_matches.index].dropna().astype(str).str.strip().tolist())
                        elif 'First Name' in reg_players.columns and 'Surname' in reg_players.columns:
                            comp = reg_players['First Name'].astype(str).str.strip() + ' ' + reg_players['Surname'].astype(str).str.strip()
                            reg_matches = reg_players[comp.str.contains(clean_q, case=False, na=False)]
                            target_official_names.update(comp[reg_matches.index].dropna().astype(str).str.strip().tolist())

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
                        raw_unique_players = list(set(found_batters + found_bowlers + found_ab))
                        
                        def player_sort_key(name):
                            pure_name = name.split(' (')[0].strip()
                            club = player_club_map.get(name.lower(), "Unknown Club").lower()
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
                        st.session_state.data_loaded = True

                matched_batting = st.session_state.matched_batting
                matched_bowling = st.session_state.matched_bowling
                matched_abandoned = st.session_state.matched_abandoned
                unique_players = st.session_state.unique_players
                reg_players = st.session_state.reg_players
                aliases_df = st.session_state.aliases_df

                if matched_batting.empty and matched_bowling.empty and matched_abandoned.empty:
                    st.error(f"No statistics found for '{current_query}'. Please try another name.")
                else:
                    def get_club_for_player(name):
                        if '(' in name and ')' in name: return name.split('(')[-1].replace(')', '').strip()
                        club = st.session_state.player_club_map.get(name.lower(), None)
                        if club and str(club).lower() not in ['nan', 'none', '', 'unknown club']:
                            return str(club).replace(" Cricket Club", "").replace(" CC", "").strip()
                        return "Unknown Club"

                    def format_player_display(name):
                        pure = name.split(' (')[0].strip()
                        club_clean = get_club_for_player(name)
                        
                        # Check for transfer date
                        transfer_suffix = ""
                        try:
                            df_reg = st.session_state.reg_players
                            name_col = '_computed_name' if '_computed_name' in df_reg.columns else ('Full Name' if 'Full Name' in df_reg.columns else df_reg.columns[0])
                            reg_match = df_reg[df_reg[name_col].astype(str).str.strip().str.lower() == pure.lower()]
                            if not reg_match.empty and 'Transfer Date' in reg_match.columns:
                                t_date = reg_match.iloc[0]['Transfer Date']
                                if pd.notna(t_date):
                                    day = t_date.day
                                    suffix = 'th' if 11 <= day <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
                                    formatted_date = f"{day}{suffix} {t_date.strftime('%B %Y')}"
                                    transfer_suffix = f" (transferred {formatted_date})"
                        except:
                            pass

                        p_aliases = eng.get_player_aliases(pure, aliases_df)
                        if p_aliases: return f"{pure} / {' / '.join(p_aliases)} ({club_clean}){transfer_suffix}"
                        return f"{pure} ({club_clean}){transfer_suffix}"

                    if len(unique_players) == 1:
                        active_player = unique_players[0]
                        pure_registered_name = active_player.split(' (')[0].strip()
                        p_aliases = eng.get_player_aliases(pure_registered_name, aliases_df)
                        st.success(f"Found Match: {format_player_display(active_player)}")
                        
                        p_bat = matched_batting[matched_batting['Name'] == active_player]
                        p_bowl = matched_bowling[matched_bowling['Bowler'] == active_player]
                        p_ab = matched_abandoned[matched_abandoned['Cleaned Name'] == active_player] if not matched_abandoned.empty else pd.DataFrame()
                        
                        doc_io, filename = eng.generate_single_player_doc(active_player, p_bat, p_bowl, reg_players, domain, aliases_list=p_aliases, player_abandoned=p_ab)
                        st.download_button("📥 Download Player Word Document", data=doc_io.getvalue(), file_name=filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary")
                    else:
                        st.warning(f"Multiple players match '{current_query}'. Please select the players to generate reports for.")
                        select_all = st.toggle("Select all players")
                        selected_players = st.multiselect("Select players:", options=unique_players, default=unique_players if select_all else [], format_func=format_player_display)
                        
                        if selected_players:
                            if len(selected_players) == 1:
                                active_player = selected_players[0]
                                pure_registered_name = active_player.split(' (')[0].strip()
                                p_aliases = eng.get_player_aliases(pure_registered_name, aliases_df)
                                p_bat = matched_batting[matched_batting['Name'] == active_player]
                                p_bowl = matched_bowling[matched_bowling['Bowler'] == active_player]
                                p_ab = matched_abandoned[matched_abandoned['Cleaned Name'] == active_player] if not matched_abandoned.empty else pd.DataFrame()
                                
                                doc_io, filename = eng.generate_single_player_doc(active_player, p_bat, p_bowl, reg_players, domain, aliases_list=p_aliases, player_abandoned=p_ab)
                                st.download_button(f"📥 Download Report for {format_player_display(active_player)}", data=doc_io.getvalue(), file_name=filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", key="dl_single_multi")
                            else:
                                zip_buffer = io.BytesIO()
                                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                                    for active_player in selected_players:
                                        pure = active_player.split(' (')[0].strip()
                                        p_aliases = eng.get_player_aliases(pure, aliases_df)
                                        p_bat = matched_batting[matched_batting['Name'] == active_player]
                                        p_bowl = matched_bowling[matched_bowling['Bowler'] == active_player]
                                        p_ab = matched_abandoned[matched_abandoned['Cleaned Name'] == active_player] if not matched_abandoned.empty else pd.DataFrame()
                                        doc_io, filename = eng.generate_single_player_doc(active_player, p_bat, p_bowl, reg_players, domain, aliases_list=p_aliases, player_abandoned=p_ab)
                                        zip_file.writestr(filename, doc_io.getvalue())
                                        
                                st.download_button(f"📦 Download Reports for {len(selected_players)} Players (ZIP)", data=zip_buffer.getvalue(), file_name=f"Player_Reports_{current_query.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.zip", mime="application/zip", type="primary", key="dl_zip_multi")

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
            st.divider()           
            st.subheader("Select Date Range")
            start_date = st.date_input("Start Date", value=datetime.today() - timedelta(days=1))
            end_date = st.date_input("End Date", value=datetime.today())
            st.divider()           

        with st.sidebar:
            c_files = eng.DEFAULT_FILES[domain]
            with st.expander("📁 File Path Configurations", expanded=False):
                f_reg = st.text_input("Official Registry (Excel)", value=c_files["reg"], key=f"reg_check_reg_{domain}")
                f_alias = st.text_input("Aliases Master (Excel)", value=c_files["alias"], key=f"reg_check_alias_{domain}")
                f_starring = st.text_input("Starring Master (Excel)", value=c_files["starring"], key=f"reg_check_starring_{domain}")
                f_cup = st.text_input("Cup Master (Excel)", value="NCU_Cup_Fixtures.xlsx", key=f"reg_check_cup_{domain}")
                f_league = st.text_input("League Structure (Excel)", value=c_files["league"], key=f"reg_check_league_{domain}")
                f_bat = st.text_input("Batting Stats (Excel)", value=c_files["bat"], key=f"reg_check_bat_{domain}")
                f_bowl = st.text_input("Bowling Stats (Excel)", value=c_files["bowl"], key=f"reg_check_bowl_{domain}")
                f_abandoned = st.text_input("Abandoned Games Stats (Excel)", value=c_files.get("abandoned", ""), key=f"reg_check_ab_{domain}")
        
        include_irish = False
        if domain == "Men's":
            include_irish = st.toggle("Include Irish Competitions in Audit?", value=False)
            if include_irish:
                with st.sidebar:
                    with st.expander("📁 Irish File Path Configurations", expanded=False):
                        f_irish_bat = st.text_input("Irish Batting Stats (Excel)", value="Irish Competitions 2026 Batting stats.xlsx", key="reg_irish_bat")
                        f_irish_bowl = st.text_input("Irish Bowling Stats (Excel)", value="Irish Competitions 2026 Bowling stats.xlsx", key="reg_irish_bowl")

        st.subheader("Run Audit Engine")
        if st.button("🚀 Execute Security Audit", type="primary"):
            files_to_check = [f for f in [f_reg, f_alias, f_league, f_bat, f_bowl, f_cup] if f]
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
                            excel_io, doc_io = eng.run_registration_audit(domain, start_ts, end_ts, f_reg, f_alias, f_starring, f_league, f_bat, f_bowl, f_irish_bat, f_irish_bowl, f_cup, f_abandoned=f_abandoned)
                        else:
                            excel_io, doc_io = eng.run_registration_audit(domain, start_ts, end_ts, f_reg, f_alias, f_starring, f_league, f_bat, f_bowl, f_cup=f_cup, f_abandoned=f_abandoned)
                        
                        try:
                            excel_io.seek(0)
                            df_unreg = pd.read_excel(excel_io, sheet_name="Unregistered Matches")
                            df_deemed = pd.read_excel(excel_io, sheet_name="Deemed Registered")
                            df_starring_viols = pd.read_excel(excel_io, sheet_name="Starring Violations")
                            
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
                                
                        st.download_button("📦 Download Audit Results (ZIP)", data=zip_buffer.getvalue(), file_name=f"{prefix}_Registration_Audit_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip", mime="application/zip", type="primary")
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
            st.divider()           
            st.subheader("Select Date Range")
            start_date = st.date_input("Start Date", value=datetime.today() - timedelta(days=1))
            end_date = st.date_input("End Date", value=datetime.today())
            st.divider() 
            with st.expander("📁 File Path Configurations", expanded=False):
                f_reg = st.text_input("Official Registry (Excel)", value=eng.DEFAULT_FILES["Midweek"]["reg"], key="mw_check_reg")
                f_alias = st.text_input("Aliases Master (Excel)", value=eng.DEFAULT_FILES["Midweek"]["alias"], key="mw_check_alias")
                f_starring = st.text_input("Men's Starring Master (Excel)", value=eng.DEFAULT_FILES["Men's"]["starring"], key="mw_check_starring")
                f_weekend_league = st.text_input("Weekend League Structure (Excel)", value=eng.DEFAULT_FILES["Men's"]["league"], key="mw_check_wknd_league")
                f_midweek_league = st.text_input("Midweek League Structure (Excel)", value=eng.DEFAULT_FILES["Midweek"]["league"], key="mw_check_mw_league")
                f_bat = st.text_input("Midweek Batting Stats (Excel)", value=eng.DEFAULT_FILES["Midweek"]["bat"], key="mw_check_bat")
                f_bowl = st.text_input("Midweek Bowling Stats (Excel)", value=eng.DEFAULT_FILES["Midweek"]["bowl"], key="mw_check_bowl")
                f_abandoned = st.text_input("Abandoned Games Stats (Excel)", value=eng.DEFAULT_FILES["Midweek"].get("abandoned", ""), key="mw_check_ab")

        st.subheader("Run Midweek Audit Engine")
        if st.button("🚀 Execute Midweek Audit", type="primary"):
            files_to_check = [f for f in [f_reg, f_alias, f_starring, f_weekend_league, f_midweek_league, f_bat, f_bowl] if f]
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
                        excel_io, doc_io = eng.run_midweek_registration_audit(start_ts, end_ts, f_reg, f_alias, f_starring, f_weekend_league, f_midweek_league, f_bat, f_bowl, f_abandoned=f_abandoned)
                        
                        try:
                            excel_io.seek(0)
                            df_unreg = pd.read_excel(excel_io, sheet_name="Unregistered Matches")
                            df_deemed = pd.read_excel(excel_io, sheet_name="Deemed Registered")
                            df_starring_viols = pd.read_excel(excel_io, sheet_name="Starring Violations")
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
        st.divider() 
        c_files = eng.DEFAULT_FILES[domain]
        with st.expander("📁 File Path Configurations", expanded=False):
            f_reg = st.text_input("Official Registry (Excel)", value=c_files["reg"], key=f"star_reg_{domain}")
            f_alias = st.text_input("Aliases Master (Excel)", value=c_files["alias"], key=f"star_alias_{domain}")
            f_starring = st.text_input("Starring Master (Excel)", value=c_files["starring"], key=f"star_starring_{domain}")
            f_bat = st.text_input("Batting Stats (Excel)", value=c_files["bat"], key=f"star_bat_{domain}")
            f_bowl = st.text_input("Bowling Stats (Excel)", value=c_files["bowl"], key=f"star_bowl_{domain}")
            f_abandoned = st.text_input("Abandoned Games Stats (Excel)", value=c_files.get("abandoned", ""), key=f"star_ab_{domain}")
            if domain == "Men's" and include_irish:
                f_irish_bat = st.text_input("Irish Batting Stats (Excel)", value="Irish Competitions 2026 Batting stats.xlsx", key="star_irish_bat")
                f_irish_bowl = st.text_input("Irish Bowling Stats (Excel)", value="Irish Competitions 2026 Bowling stats.xlsx", key="star_irish_bowl")

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
                    prefix = domain.replace("'", "")
                    st.download_button("📥 Download Club Reports (ZIP)", data=zip_buffer.getvalue(), file_name=f"{prefix}_Starring_Reports_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip", mime="application/zip", type="primary")
                except Exception as e:
                    st.error(f"An error occurred during processing: {str(e)}")

# ==========================================
# TOOL 5: CLUB FINES GENERATOR
# ==========================================
elif app_mode == "Club Fines Generator":
    st.title(PAGE_TITLES["fines_generator"])
    st.markdown("Automatically run the registration audit engine to find violations and merge them with forfeited matches to generate a club-by-club fines report.")
    
    if not DOCX_AVAILABLE:
        st.error("The `python-docx` library is not installed. Please run `pip install python-docx` to use this feature.")
    else:
        st.subheader("Select League Domain")
        domain = st.radio("Choose the dataset domain to audit:", ["Men's", "Women's", "Midweek"], horizontal=True)

        with st.sidebar:
            st.divider()           
            st.subheader("Select Date Range")
            start_date = st.date_input("Start Date", value=datetime.today() - timedelta(days=7), key="fines_start")
            end_date = st.date_input("End Date", value=datetime.today(), key="fines_end")
            st.divider() 
            c_files = eng.DEFAULT_FILES.get(domain, eng.DEFAULT_FILES["Men's"])
            with st.expander("📁 File Path Configurations", expanded=False):
                f_reg = st.text_input("Official Registry (Excel)", value=c_files["reg"], key=f"fines_reg_{domain}")
                f_alias = st.text_input("Aliases Master (Excel)", value=c_files["alias"], key=f"fines_alias_{domain}")
                f_bat = st.text_input("Batting Stats (Excel)", value=c_files["bat"], key=f"fines_bat_{domain}")
                f_bowl = st.text_input("Bowling Stats (Excel)", value=c_files["bowl"], key=f"fines_bowl_{domain}")
                f_abandoned = st.text_input("Abandoned Games Stats (Excel)", value=c_files.get("abandoned", ""), key=f"fines_ab_{domain}")
                
                if domain != "Midweek":
                    f_starring = st.text_input("Starring Master (Excel)", value=c_files["starring"], key=f"fines_starring_{domain}")
                    f_league = st.text_input("League Structure (Excel)", value=c_files["league"], key=f"fines_league_{domain}")
                    f_cup = st.text_input("Cup Master (Excel)", value="NCU_Cup_Fixtures.xlsx", key=f"fines_cup_{domain}")
                else:
                    f_starring = st.text_input("Men's Starring Master (Excel)", value=eng.DEFAULT_FILES["Men's"]["starring"], key="fines_mw_starring")
                    f_weekend_league = st.text_input("Weekend League Structure (Excel)", value=eng.DEFAULT_FILES["Men's"]["league"], key="fines_wknd_league")
                    f_midweek_league = st.text_input("Midweek League Structure (Excel)", value=c_files["league"], key="fines_mw_league")

        include_irish = False
        if domain == "Men's":
            include_irish = st.toggle("Include Irish Competitions in Audit?", value=False, key="fines_irish_check")
            if include_irish:
                with st.sidebar:
                    with st.expander("📁 Irish File Path Configurations", expanded=False):
                        f_irish_bat = st.text_input("Irish Batting Stats (Excel)", value="Irish Competitions 2026 Batting stats.xlsx", key="fines_irish_bat")
                        f_irish_bowl = st.text_input("Irish Bowling Stats (Excel)", value="Irish Competitions 2026 Bowling stats.xlsx", key="fines_irish_bowl")
        
        st.divider()
        st.subheader("Forfeited Matches Data")
        default_forfeit_path = "Team Fines for forfeiting matches 2026.xlsx"
        col1, col2 = st.columns([1, 2])
        with col1:
            use_default_forfeit = st.toggle(f"Use local '{default_forfeit_path}'", value=os.path.exists(default_forfeit_path))
        with col2:
            f_forfeit = st.file_uploader("Or Upload Forfeits Excel File", type=["xlsx"], key="fines_forfeit_upload")

        st.divider()
        if st.button("📄 Run Engine & Generate Fines Report", type="primary"):
            forfeit_path = f_forfeit if f_forfeit is not None else (default_forfeit_path if use_default_forfeit and os.path.exists(default_forfeit_path) else None)
            files_to_check = [f_reg, f_alias, f_bat, f_bowl]
            if f_abandoned: files_to_check.append(f_abandoned)
            if domain != "Midweek":
                files_to_check.extend([f_starring, f_league])
                if domain == "Men's" and include_irish: files_to_check.extend([f_irish_bat, f_irish_bowl])
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
                                audit_excel_io, _ = eng.run_registration_audit(domain, start_ts, end_ts, f_reg, f_alias, f_starring, f_league, f_bat, f_bowl, f_irish_bat, f_irish_bowl, f_cup, f_abandoned=f_abandoned)
                            else:
                                audit_excel_io, _ = eng.run_registration_audit(domain, start_ts, end_ts, f_reg, f_alias, f_starring, f_league, f_bat, f_bowl, f_cup=f_cup, f_abandoned=f_abandoned)
                        else:
                            audit_excel_io, _ = eng.run_midweek_registration_audit(start_ts, end_ts, f_reg, f_alias, f_starring, f_weekend_league, f_midweek_league, f_bat, f_bowl, f_abandoned=f_abandoned)
                            
                        audit_excel_io.seek(0)
                        doc_io = eng.generate_club_fines_report(audit_excel_io, forfeit_path, start_ts, end_ts)
                        prefix = domain.replace("'", "")
                        st.success("✅ Fines report generated successfully!")
                        st.download_button(f"📥 Download {domain} Fines Report (Word)", data=doc_io.getvalue(), file_name=f"NCU_{prefix}_Fines_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary")  
                    except Exception as e:
                        st.error(f"An error occurred during processing: {str(e)}")

# ==========================================
# TOOL 6: UNREGISTERED FINES GENERATOR
# ==========================================
elif app_mode == "Unregistered Player Fines Generator":
    st.title(PAGE_TITLES["unregistered_fines"])
    st.markdown("Automatically run the registration audit engine to generate a standalone fines report isolated exclusively to unregistered players.")
    
    if not DOCX_AVAILABLE:
        st.error("The `python-docx` library is not installed. Please run `pip install python-docx` to use this feature.")
    else:
        st.subheader("Select League Domain")
        domain = st.radio("Choose the dataset domain to audit:", ["Men's", "Women's", "Midweek"], horizontal=True, key="unreg_domain")

        with st.sidebar:
            st.divider()           
            st.subheader("Select Date Range")
            start_date = st.date_input("Start Date", value=datetime.today() - timedelta(days=7), key="unreg_start")
            end_date = st.date_input("End Date", value=datetime.today(), key="unreg_end")
            st.divider() 
            c_files = eng.DEFAULT_FILES.get(domain, eng.DEFAULT_FILES["Men's"])
            with st.expander("📁 File Path Configurations", expanded=False):
                f_reg = st.text_input("Official Registry (Excel)", value=c_files["reg"], key=f"unreg_reg_{domain}")
                f_alias = st.text_input("Aliases Master (Excel)", value=c_files["alias"], key=f"unreg_alias_{domain}")
                f_bat = st.text_input("Batting Stats (Excel)", value=c_files["bat"], key=f"unreg_bat_{domain}")
                f_bowl = st.text_input("Bowling Stats (Excel)", value=c_files["bowl"], key=f"unreg_bowl_{domain}")
                f_abandoned = st.text_input("Abandoned Games Stats (Excel)", value=c_files.get("abandoned", ""), key=f"unreg_ab_{domain}")
                
                if domain != "Midweek":
                    f_starring = st.text_input("Starring Master (Excel)", value=c_files["starring"], key=f"unreg_starring_{domain}")
                    f_league = st.text_input("League Structure (Excel)", value=c_files["league"], key=f"unreg_league_{domain}")
                    f_cup = st.text_input("Cup Master (Excel)", value="NCU_Cup_Fixtures.xlsx", key=f"unreg_cup_{domain}")
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
                        f_irish_bat = st.text_input("Irish Batting Stats (Excel)", value="Irish Competitions 2026 Batting stats.xlsx", key="unreg_irish_bat")
                        f_irish_bowl = st.text_input("Irish Bowling Stats (Excel)", value="Irish Competitions 2026 Bowling stats.xlsx", key="unreg_irish_bowl")
        
        st.divider()
        if st.button("📄 Run Engine & Generate Unregistered Report", type="primary"):
            files_to_check = [f_reg, f_alias, f_bat, f_bowl]
            if f_abandoned: files_to_check.append(f_abandoned)
            if domain != "Midweek":
                files_to_check.extend([f_starring, f_league])
                if domain == "Men's" and include_irish: files_to_check.extend([f_irish_bat, f_irish_bowl])
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
                                audit_excel_io, _ = eng.run_registration_audit(domain, start_ts, end_ts, f_reg, f_alias, f_starring, f_league, f_bat, f_bowl, f_irish_bat, f_irish_bowl, f_cup, f_abandoned=f_abandoned)
                            else:
                                audit_excel_io, _ = eng.run_registration_audit(domain, start_ts, end_ts, f_reg, f_alias, f_starring, f_league, f_bat, f_bowl, f_cup=f_cup, f_abandoned=f_abandoned)
                        else:
                            audit_excel_io, _ = eng.run_midweek_registration_audit(start_ts, end_ts, f_reg, f_alias, f_starring, f_weekend_league, f_midweek_league, f_bat, f_bowl, f_abandoned=f_abandoned)
                            
                        audit_excel_io.seek(0)
                        doc_io = eng.generate_unregistered_fines_only(audit_excel_io)
                        prefix = domain.replace("'", "")
                        st.success("✅ Unregistered Fines report generated successfully!")
                        st.download_button(f"📥 Download {domain} Unregistered Fines Report (Word)", data=doc_io.getvalue(), file_name=f"NCU_{prefix}_Unreg_Fines_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary")
                    except Exception as e:
                        st.error(f"An error occurred during processing: {str(e)}")

# ==========================================
# TOOL 7: CLUB CONTACTS & OFFICIALS DIRECTORY
# ==========================================
elif app_mode == "Club Contacts Directory":
    st.title(PAGE_TITLES["club_contacts"])
    st.markdown("Search and filter official union contacts by Club, Team Tier, or Union-Wide Role with one-tap calling and messaging.")

    with st.sidebar:
        st.divider()
        with st.expander("📁 Contacts Master Configuration", expanded=False):
            f_contacts = st.text_input("Contacts Excel File", value="2026 Season Club Contacts.xlsx", key="contacts_filepath")

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

        all_clubs = sorted([c for c in df_contacts['Club'].unique() if c and c.lower() != 'nan'])
        
        tier_hierarchy = [
            "All Roles & Officials",
            "Club Official",
            "1st XI", "2nd XI", "3rd XI", "4th XI", "5th XI", "6th XI",
            "Women's 1st XI", "Women's 2nd XI", "Women's 3rd XI",
            "1st Midweek XI", "2nd Midweek XI",
            "Boys Youth", "Girls Youth", "Indoor Cricket"
        ]
        present_tiers = [t for t in tier_hierarchy if t == "All Roles & Officials" or t in df_contacts['Team Tier'].unique()]

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
                if selected_tier != "All Roles & Officials":
                    filtered_view = club_matches[club_matches['Team Tier'] == selected_tier]
                else:
                    filtered_view = club_matches

                st.divider()
                # Using custom font size instead of st.subheader
                st.markdown(f"<div style='font-size: {UI_CLUB_HEADER_SIZE}; font-weight: bold; padding-top: 1rem; padding-bottom: 1rem;'>📌 {selected_club} — {selected_tier}</div>", unsafe_allow_html=True)

                if filtered_view.empty:
                    st.info(f"No contact records found for {selected_club} under {selected_tier}.")
                else:
                    cols = st.columns(min(len(filtered_view), 3) if len(filtered_view) > 0 else 1)
                    for idx, (_, row) in enumerate(filtered_view.iterrows()):
                        target_col = cols[idx % len(cols)]
                        with target_col:
                            with st.container(border=True):
                                role_title = row.get('Role', 'Club Official')
                                official_name = row.get('Name', 'Not Listed')
                                phone_val = row.get('Phone', '')
                                email_val = row.get('Email', '')

                                st.markdown(f"<div style='font-size: {UI_ROLE_TITLE_SIZE}; font-weight: bold; margin-bottom: 0.25rem;'>{role_title}</div>", unsafe_allow_html=True)
                                st.markdown(f"<div style='font-size: {UI_OFFICIAL_NAME_SIZE}; margin-bottom: 0.25rem;'>👤 <b>{official_name}</b></div>", unsafe_allow_html=True)
                                st.markdown(f"**Category:** `{row.get('Team Tier', 'General')}`")
                                st.markdown(format_tel_link(phone_val), unsafe_allow_html=True)
                                st.markdown(format_mail_link(email_val), unsafe_allow_html=True)

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
                        display_table.to_html(escape=False, index=False, justify='left'), 
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
                        matched_rows[['Club', 'Role', 'Team Tier', 'Name', 'Direct Phone', 'Direct Email']].to_html(escape=False, index=False, justify='left'),
                        unsafe_allow_html=True
                    )