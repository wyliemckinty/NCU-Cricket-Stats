# ==========================================
# stats_app.py (Stats & Milestones Hub)
# ==========================================
import streamlit as st
import pandas as pd
import os
import io
import json
import re
from datetime import datetime
import warnings

# Import shared core engine
import importlib
import engine as eng
importlib.reload(eng)

warnings.filterwarnings('ignore')

# ==========================================
# STREAMLIT CACHED EXCEL LOADERS
# ==========================================
def get_excel_df(filepath):
    return eng.get_excel_df(filepath)

def get_excel_sheet_df(filepath, sheet_name=None, header='infer'):
    return eng.get_excel_sheet_df(filepath, sheet_name=sheet_name, header=header)

# Enable docx for Milestones Report
try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

MAIN_HEADER_SIZE = "28px"
CONFIG_FILE = "threshold_settings.json"

DEFAULT_THRESHOLDS = {
    # Men's Thresholds
    "p_runs": 400, "p_bmat": 10, "p_wick": 25, "p_mmat": 5,
    "s1_runs": 300, "s1_bmat": 10, "s1_wick": 20, "s1_mmat": 5,
    "s2_runs": 200, "s2_bmat": 5, "s2_wick": 15, "s2_mmat": 5,
    "s3_runs": 200, "s3_bmat": 5, "s3_wick": 15, "s3_mmat": 5,
    "t1_runs": 200, "t1_bmat": 5, "t1_wick": 15, "t1_mmat": 5,
    "t2_runs": 150, "t2_bmat": 5, "t2_wick": 10, "t2_mmat": 5,
    "t3_runs": 100, "t3_bmat": 3, "t3_wick": 5,  "t3_mmat": 3,
    "t4_runs": 50,  "t4_bmat": 3, "t4_wick": 3,  "t4_mmat": 3,
    **{f"j{i}_runs": 100 for i in list(range(1, 11)) + ['11a', '11b', '11c']},
    **{f"j{i}_bmat": 3 for i in list(range(1, 11)) + ['11a', '11b', '11c']},
    **{f"j{i}_wick": 5 for i in list(range(1, 11)) + ['11a', '11b', '11c']},
    **{f"j{i}_mmat": 3 for i in list(range(1, 11)) + ['11a', '11b', '11c']},
    # Women's Thresholds
    "wp_runs": 100, "wp_bmat": 5, "wp_wick": 10, "wp_mmat": 5,
    "ws1_runs": 100, "ws1_bmat": 5, "ws1_wick": 7, "ws1_mmat": 5,
    "wj1_runs": 0, "wj1_bmat": 0, "wj1_wick": 0, "wj1_mmat": 0,
    # Midweek Thresholds
    "mw_min_runs": 50, "mw_min_innings": 0, "mw_min_wickets": 5, "mw_min_bowl_innings": 0
}

# Threshold Persistence Helpers
def init_threshold_store():
    if "threshold_store" not in st.session_state:
        store = dict(DEFAULT_THRESHOLDS)
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    saved = json.load(f)
                    for k, v in saved.items():
                        if k in DEFAULT_THRESHOLDS: store[k] = int(v)
            except Exception: pass
        st.session_state["threshold_store"] = store

def get_threshold_val(key):
    init_threshold_store()
    if st.session_state.get("disable_thresholds", False): return 0
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

# Page Setup
st.set_page_config(page_title="NCU Stats Hub", page_icon="🏏", layout="wide")

st.markdown(f"""
<style>
    h1 {{ font-size: {MAIN_HEADER_SIZE} !important; font-weight: 700; }}
    div.stButton > button {{ white-space: nowrap !important; }}
    div.stButton > button[kind="primary"] {{ border-radius: 8px; padding: 0.5rem 1.5rem; }}
    div.stDownloadButton > button:first-child {{ background-color: #0066cc; color: white; border-radius: 8px; border: none; padding: 0.5rem 1.5rem; }}
    
    /* Highlight the league selection dropdowns */
    div.element-container:has(.league-dropdown-marker) + div.element-container div[data-baseweb="select"] > div {{
        background-color: #e0f2fe !important;
        border: 1px solid #7dd3fc !important;
    }}
</style>
""", unsafe_allow_html=True)

# Sidebar Navigation
with st.sidebar:
    st.title("🏏 NCU Stats Hub")
    st.header("🛠️ Navigation")
    app_mode = st.radio("Choose a module to run:", ["Bulk Averages Calculator", "League Milestones Report"])
    st.divider()

# ==========================================
# TOOL 1: BULK AVERAGES
# ==========================================
def filter_match_formats(batting_df, bowling_df, f_cup, domain, include_cup, include_t20, include_pathway=False):
    """
    Filters scorecard records according to Cup and T20 inclusion toggles
    with strict date matching and explicit League match protection.
    """
    if include_cup and include_t20 and include_pathway:
        return batting_df, bowling_df

    cup_match_set = set()  # Stores (team1, team2, date_YYYY-MM-DD)
    t20_match_set = set()  # Stores (team1, team2, date_YYYY-MM-DD)

    if f_cup and os.path.exists(f_cup):
        try:
            excel_file_cup = pd.ExcelFile(f_cup)
            target_sheet = excel_file_cup.sheet_names[0]
            for sheet in excel_file_cup.sheet_names:
                if domain.lower().replace("'", "") in sheet.lower().replace("'", ""):
                    target_sheet = sheet
                    break
            cup_df = pd.read_excel(f_cup, sheet_name=target_sheet, header=None)

            for _, row_data in cup_df.iterrows():
                match_str_raw = str(row_data[0]).strip()
                cup_name = str(row_data[1]).strip()
                if match_str_raw.lower() in ['match string', 'match group', 'match', 'nan']:
                    continue

                parts = match_str_raw.rsplit(' - ', 1)
                rest = parts[0].strip()
                d_str = parts[1].strip() if len(parts) == 2 else ""

                # Clean ordinal dates (e.g., 24th -> 24)
                clean_d = re.sub(r'(?<=\d)(st|nd|rd|th)\b', '', d_str, flags=re.IGNORECASE).strip()
                dt = pd.to_datetime(clean_d, dayfirst=True, errors='coerce')

                if ' v ' in rest and pd.notna(dt):
                    t_a, rem = rest.split(' v ', 1)
                    t_b = rem.rsplit(', ', 1)[0] if ', ' in rem else rem
                    teams = sorted([t_a.strip().lower(), t_b.strip().lower()])
                    date_key = dt.strftime('%Y-%m-%d')
                    key = (teams[0], teams[1], date_key)

                    is_t20_comp = any(kw in cup_name.lower() or kw in match_str_raw.lower() for kw in ['t20', 'twenty20', 'lvs'])
                    if is_t20_comp:
                        t20_match_set.add(key)
                    else:
                        cup_match_set.add(key)
        except Exception:
            pass

    def classify_match(grp_str):
        grp_lower = str(grp_str).lower()

        # 1. Protect explicit League matches first
        is_explicit_league = any(kw in grp_lower for kw in [
            'premier league', 'senior league', 'junior league',
            'mercury premier', 'mercury senior', 'mercury junior',
            'section 1', 'section 2', 'section 3', 'section 4'
        ])

        # 2. Check explicit T20 match markers
        is_explicit_t20 = any(kw in grp_lower for kw in [
            'lvs t20', 'twenty20', 't20 cup', 't20 trophy', 't20 bowl', 't20 shield'
        ])
        if is_explicit_t20:
            return 't20'

        # 3. Check explicit Cup competition names
        cup_specific_kws = [
            'gallagher challenge cup', 'gallagher challenge plate',
            'junior cup', 'intermediate cup', 'lindsay cup',
            'minor qualifying cup', 'development cup', 'irish senior cup',
            'irish cup', 'national cup', 'ulster plate'
        ]
        if any(kw in grp_lower for kw in cup_specific_kws):
            return 'cup'

        # 4. Check date-strict match against Cup Fixtures Master
        if ' v ' in grp_str:
            parts = str(grp_str).rsplit(' - ', 1)
            rest = parts[0].strip()
            d_str = parts[1].strip() if len(parts) == 2 else ""
            clean_d = re.sub(r'(?<=\d)(st|nd|rd|th)\b', '', d_str, flags=re.IGNORECASE).strip()
            dt = pd.to_datetime(clean_d, dayfirst=True, errors='coerce')

            if pd.notna(dt):
                date_key = dt.strftime('%Y-%m-%d')
                t_a, rem = rest.split(' v ', 1)
                t_b = rem.rsplit(', ', 1)[0] if ', ' in rem else rem
                teams = sorted([t_a.strip().lower(), t_b.strip().lower()])
                key = (teams[0], teams[1], date_key)

                if key in t20_match_set:
                    return 't20'
                if key in cup_match_set:
                    return 'cup'

        # 5. Fallback checks only if NOT an explicit league fixture
        if not is_explicit_league:
            if 't20' in grp_lower:
                return 't20'
            if any(kw in grp_lower for kw in ['challenge cup', 'cup', 'trophy', 'plate', 'shield', 'bowl', 'vase']):
                return 'cup'

        return 'league'

    def should_keep(grp):
        grp_lower = str(grp).lower()
        if not include_pathway and 'pathway' in grp_lower:
            return False
            
        m_type = classify_match(grp)
        if m_type == 't20' and not include_t20:
            return False
        if m_type == 'cup' and not include_cup:
            return False
        return True

    filtered_batting = batting_df[batting_df['Group'].apply(should_keep)].copy()
    filtered_bowling = bowling_df[bowling_df['Group'].apply(should_keep)].copy()
    return filtered_batting, filtered_bowling
    
if app_mode == "Bulk Averages Calculator":
    init_threshold_store()
    st.title("📊 League Bulk Averages Calculator")
    
    col_dom, col_save, col_reset = st.columns([2, 1, 1])
    with col_dom:
        domain = st.radio("League Domain:", ["Men's", "Women's", "Midweek"], horizontal=True, label_visibility="collapsed")
    with col_save:
        if st.button("💾 Save Thresholds", use_container_width=True):
            save_threshold_settings()
            st.toast("Saved custom thresholds as new defaults!", icon="✅")
    with col_reset:
        if st.button("🔄 Reset Defaults", use_container_width=True):
            reset_threshold_settings()
            st.toast("Restored factory default thresholds!", icon="♻️")
    
    st.markdown("#### Minimum Thresholds")
    if domain == "Midweek":
        st.markdown("**Midweek Overall Qualifiers** *(Groups are fixed to 20 Runs / 2 Wickets)*")
        col1, col2, col3, col4 = st.columns(4)
        with col1: mw_min_runs = st.number_input("Minimum Runs", min_value=0, value=get_threshold_val("mw_min_runs"), key="mw_min_runs")
        with col2: mw_min_innings = st.number_input("Min Bat Innings", min_value=0, value=get_threshold_val("mw_min_innings"), key="mw_min_innings")
        with col3: mw_min_wickets = st.number_input("Minimum Wickets", min_value=0, value=get_threshold_val("mw_min_wickets"), key="mw_min_wickets")
        with col4: mw_min_bowl_innings = st.number_input("Min Bowl Innings", min_value=0, value=get_threshold_val("mw_min_bowl_innings"), key="mw_min_bowl_innings")
        
    elif domain == "Women's":
        st.info("💡 **Women's Tiered Rules Active:** Below are your automated target thresholds.")
        with st.container(border=True):
            c_title, c_drop = st.columns([1.5, 1])
            with c_title: st.markdown("🏏 **Women's Leagues (Individual)**")
            
            women_leagues = ['Premier League', 'Senior League Section 1', 'Junior League Section 1']
            with c_drop:
                selected_wl = st.selectbox("Select Section", women_leagues, label_visibility="collapsed")
                
            w_id = 'wp' if selected_wl == 'Premier League' else ('ws1' if 'Senior' in selected_wl else 'wj1')
            c1, c2, c3, c4 = st.columns(4)
            with c1: st.number_input("Min Runs", min_value=0, value=get_threshold_val(f"{w_id}_runs"), key=f"{w_id}_runs")
            with c2: st.number_input("Min Bat Innings", min_value=0, value=get_threshold_val(f"{w_id}_bmat"), key=f"{w_id}_bmat")
            with c3: st.number_input("Min Wickets", min_value=0, value=get_threshold_val(f"{w_id}_wick"), key=f"{w_id}_wick")
            with c4: st.number_input("Min Bowl Innings", min_value=0, value=get_threshold_val(f"{w_id}_mmat"), key=f"{w_id}_mmat")
            
    else:  # Men's
        st.info("💡 **Men's Tiered Rules Active:** Below are your automated target thresholds.")
        with st.container(border=True):
            c_title, c_drop = st.columns([1.5, 1])
            with c_title: st.markdown("🏆 **Premier & Senior Leagues (Individual)**")
            
            senior_leagues = ['Premier League', 'Section 1', 'Section 2', 'Section 3']
            with c_drop:
                selected_sl = st.selectbox("Select Section", senior_leagues, label_visibility="collapsed")
                
            s_id = 'p' if selected_sl == 'Premier League' else f"s{selected_sl[-1]}"
            c1, c2, c3, c4 = st.columns(4)
            with c1: st.number_input("Min Runs", min_value=0, value=get_threshold_val(f"{s_id}_runs"), key=f"{s_id}_runs")
            with c2: st.number_input("Min Bat Innings", min_value=0, value=get_threshold_val(f"{s_id}_bmat"), key=f"{s_id}_bmat")
            with c3: st.number_input("Min Wickets", min_value=0, value=get_threshold_val(f"{s_id}_wick"), key=f"{s_id}_wick")
            with c4: st.number_input("Min Bowl Innings", min_value=0, value=get_threshold_val(f"{s_id}_mmat"), key=f"{s_id}_mmat")
            
        with st.container(border=True):
            c_title, c_drop = st.columns([1.5, 1])
            with c_title: st.markdown("🏏 **Junior Leagues (Individual)**")
            
            junior_leagues = [str(i) for i in range(1, 11)] + ['11a', '11b', '11c']
            with c_drop:
                selected_jl = st.selectbox("Select Section", [f"Section {j}" for j in junior_leagues], label_visibility="collapsed")
                
            j_id = selected_jl.replace("Section ", "")
            c1, c2, c3, c4 = st.columns(4)
            with c1: st.number_input("Min Runs", min_value=0, value=get_threshold_val(f"j{j_id}_runs"), key=f"j{j_id}_runs")
            with c2: st.number_input("Min Bat Innings", min_value=0, value=get_threshold_val(f"j{j_id}_bmat"), key=f"j{j_id}_bmat")
            with c3: st.number_input("Min Wickets", min_value=0, value=get_threshold_val(f"j{j_id}_wick"), key=f"j{j_id}_wick")
            with c4: st.number_input("Min Bowl Innings", min_value=0, value=get_threshold_val(f"j{j_id}_mmat"), key=f"j{j_id}_mmat")

    st.markdown("#### Configuration & Sorting")
    with st.container(border=True):
        colA, colB = st.columns(2)
        with colA: bat_sort_pref = st.selectbox("Batting Sort", ["Runs", "Average", "Strike Rate"], index=0)
        with colB: bowl_sort_pref = st.selectbox("Bowling Sort", ["Wickets", "Average", "Economy", "Strike Rate"], index=0)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        include_cup, include_t20, include_pathway, include_irish = True, True, False, False
        
        # Calculate how many toggles we need to show
        toggle_cols = []
        if domain in ["Men's", "Women's"]: toggle_cols.append('cup')
        if domain == "Men's": toggle_cols.extend(['t20', 'pathway', 'irish'])
        toggle_cols.append('zero')
        
        cols = st.columns(len(toggle_cols))
        idx = 0
        if 'cup' in toggle_cols:
            with cols[idx]: include_cup = st.toggle("Include Cup", value=True)
            idx += 1
        if 't20' in toggle_cols:
            with cols[idx]: include_t20 = st.toggle("Include T20", value=True)
            idx += 1
        if 'pathway' in toggle_cols:
            with cols[idx]: include_pathway = st.toggle("Include Pathway", value=False)
            idx += 1
        if 'irish' in toggle_cols:
            with cols[idx]: include_irish = st.toggle("Include Irish", value=False)
            idx += 1
        if 'zero' in toggle_cols:
            with cols[idx]: disable_thresholds = st.toggle("Thresholds to 0", value=False, key="disable_thresholds", on_change=toggle_zero_thresholds)

    with st.sidebar:
        c_files = eng.DEFAULT_FILES[domain]
        with st.expander("📂 File Path Configurations", expanded=False):
            f_reg = st.text_input("Official Registry (Excel)", value=c_files["reg"], key=f"avg_reg_{domain}")
            f_alias = st.text_input("Aliases Master (Excel)", value=c_files["alias"], key=f"avg_alias_{domain}")
            f_unreg = st.text_input("Unregistered Players Map", value=c_files.get("unreg", ""), key=f"avg_unreg_{domain}")
            f_secondary = st.text_input("Secondary Team Map", value=c_files.get("secondary", ""), key=f"avg_sec_{domain}") 
            f_league = st.text_input("League Structure (Excel)", value=c_files["league"], key=f"avg_league_{domain}")
            f_bat = st.text_input("Batting Stats (Excel)", value=c_files["bat"], key=f"avg_bat_{domain}")
            f_bowl = st.text_input("Bowling Stats (Excel)", value=c_files["bowl"], key=f"avg_bowl_{domain}")
            f_cup = st.text_input("Cup Master (Excel)", value="NCU_Cup_Fixtures.xlsx", key=f"avg_cup_{domain}")
            if include_irish:
                f_irish_bat = st.text_input("Irish Batting Stats (Excel)", value="Irish Competitions 2026 Batting stats.xlsx", key="avg_irish_bat")
                f_irish_bowl = st.text_input("Irish Bowling Stats (Excel)", value="Irish Competitions 2026 Bowling stats.xlsx", key="avg_irish_bowl")

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚀 Process Averages", type="primary"):
        files_to_check = [f_reg, f_alias, f_league, f_bat, f_bowl, f_cup]
        if domain == "Men's" and include_irish: files_to_check.extend([f_irish_bat, f_irish_bowl])
            
        missing_files = [f for f in files_to_check if not os.path.exists(f)]
        if missing_files:
            st.error(f"Cannot find the following files:\n\n" + "\n".join([f"- {f}" for f in missing_files]))
        else:
            with st.spinner(f"Running {domain} Averages Engine..."):
                try:
                    
                    # Load datasets
                    reg_players = pd.read_excel(f_reg)
                    aliases = pd.read_excel(f_alias)
                    league_structure = pd.read_excel(f_league)
                    batting = pd.read_excel(f_bat)
                    bowling = pd.read_excel(f_bowl)
                    
                    if domain == "Men's" and include_irish:
                        if os.path.exists(f_irish_bat): batting = pd.concat([batting, pd.read_excel(f_irish_bat)], ignore_index=True)
                        if os.path.exists(f_irish_bowl): bowling = pd.concat([bowling, pd.read_excel(f_irish_bowl)], ignore_index=True)

                    # --- ADD THIS MATCH FORMAT FILTER ---
                    if domain in ["Men's", "Women's"]:
                        batting, bowling = filter_match_formats(
                            batting, bowling, f_cup, domain, include_cup, include_t20, include_pathway
                        )
                    # -----------------------------------

                    unreg_df = pd.read_excel(f_unreg) if os.path.exists(f_unreg) else None
                    sec_df = pd.read_excel(f_secondary) if os.path.exists(f_secondary) else None
                    
                    alias_map = eng.build_alias_map(aliases, domain)
                    secondary_map = eng.build_secondary_team_map(sec_df, alias_map) 
                    league_dict, team_keys, original_league_order = eng.build_league_dict(league_structure)
                    player_club_map = eng.build_player_club_map(reg_players, alias_map, domain, unreg_map_df=unreg_df)
                    player_club_map = eng.infer_unregistered_player_clubs(batting, bowling, player_club_map, min_matches=2)
                    
                    batting['Cleaned Name'] = batting.apply(lambda r: eng.cleanse_name_contextual(r['Name'], r, alias_map, player_club_map), axis=1)
                    bowling['Cleaned Name'] = bowling.apply(lambda r: eng.cleanse_name_contextual(r['Bowler'], r, alias_map, player_club_map), axis=1)
                    
                    batting_avgs, bowling_avgs = eng.calculate_averages(
                        batting, bowling, player_club_map, team_keys, league_dict, domain,
                        bat_sort_pref, bowl_sort_pref, secondary_map=secondary_map,
                        alias_map=alias_map, _cache_version=datetime.now().timestamp()
                    )                   
                    display_league_order = []
                    for raw_league in original_league_order:
                        if domain == "Midweek": display_league_order.append(str(raw_league))
                        else:
                            league_str = str(raw_league)
                            target_words = ['premier', 'senior league 1', 'senior league 2', 'senior league 3'] if domain == "Men's" else ['premier', 'senior league 1', 'senior league 2', 'senior league 3', 'senior']
                            if any(word in league_str.lower() for word in target_words): display_league_order.append(league_str.replace('NCU', 'Mercury'))
                            else: display_league_order.append(league_str)
                    
                    unique_leagues = sorted(
                        list(set(batting_avgs['League'].unique()).union(set(bowling_avgs['League'].unique()))), 
                        key=lambda x: eng.custom_league_sort(x, domain, display_league_order)
                    )

                    output_buffer = io.BytesIO()
                    with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
                        for league in unique_leagues:
                            league_bat = batting_avgs[batting_avgs['League'] == league].drop(columns=['League'])
                            league_bowl = bowling_avgs[bowling_avgs['League'] == league].drop(columns=['League'])
                            tab_prefix = league.replace("League", "Lge").replace("Midweek", "MW").replace("Group", "Grp").replace("Overall", "Ovr").replace("Unassigned", "Non NCU Players").strip()[:24].strip()
                            
                            name_lower = str(league).lower()
                            if domain == "Midweek":
                                mw_r = st.session_state.get("mw_min_runs", get_threshold_val("mw_min_runs"))
                                mw_i = st.session_state.get("mw_min_innings", get_threshold_val("mw_min_innings"))
                                mw_w = st.session_state.get("mw_min_wickets", get_threshold_val("mw_min_wickets"))
                                mw_bowl_i = st.session_state.get("mw_min_bowl_innings", get_threshold_val("mw_min_bowl_innings"))
                                bat_thresh, bat_match_thresh = (mw_r, mw_i) if "Overall" in league else (20, 0)
                                bowl_thresh, bowl_match_thresh = (mw_w, mw_bowl_i) if "Overall" in league else (2, 0)
                            elif domain == "Women's":
                                if "premier" in name_lower:
                                    bat_thresh = st.session_state.get("wp_runs", get_threshold_val("wp_runs"))
                                    bat_match_thresh = st.session_state.get("wp_bmat", get_threshold_val("wp_bmat"))
                                    bowl_thresh = st.session_state.get("wp_wick", get_threshold_val("wp_wick"))
                                    bowl_match_thresh = st.session_state.get("wp_mmat", get_threshold_val("wp_mmat"))
                                elif "senior" in name_lower:
                                    bat_thresh = st.session_state.get("ws1_runs", get_threshold_val("ws1_runs"))
                                    bat_match_thresh = st.session_state.get("ws1_bmat", get_threshold_val("ws1_bmat"))
                                    bowl_thresh = st.session_state.get("ws1_wick", get_threshold_val("ws1_wick"))
                                    bowl_match_thresh = st.session_state.get("ws1_mmat", get_threshold_val("ws1_mmat"))
                                else:
                                    bat_thresh = st.session_state.get("wj1_runs", get_threshold_val("wj1_runs"))
                                    bat_match_thresh = st.session_state.get("wj1_bmat", get_threshold_val("wj1_bmat"))
                                    bowl_thresh = st.session_state.get("wj1_wick", get_threshold_val("wj1_wick"))
                                    bowl_match_thresh = st.session_state.get("wj1_mmat", get_threshold_val("wj1_mmat"))
                            else:  # Men's
                                import re
                                if "premier" in name_lower:
                                    bat_thresh = st.session_state.get("p_runs", get_threshold_val("p_runs"))
                                    bat_match_thresh = st.session_state.get("p_bmat", get_threshold_val("p_bmat"))
                                    bowl_thresh = st.session_state.get("p_wick", get_threshold_val("p_wick"))
                                    bowl_match_thresh = st.session_state.get("p_mmat", get_threshold_val("p_mmat"))
                                elif "senior league 1" in name_lower or "senior 1" in name_lower:
                                    bat_thresh = st.session_state.get("s1_runs", get_threshold_val("s1_runs"))
                                    bat_match_thresh = st.session_state.get("s1_bmat", get_threshold_val("s1_bmat"))
                                    bowl_thresh = st.session_state.get("s1_wick", get_threshold_val("s1_wick"))
                                    bowl_match_thresh = st.session_state.get("s1_mmat", get_threshold_val("s1_mmat"))
                                elif "senior league 2" in name_lower or "senior 2" in name_lower:
                                    bat_thresh = st.session_state.get("s2_runs", get_threshold_val("s2_runs"))
                                    bat_match_thresh = st.session_state.get("s2_bmat", get_threshold_val("s2_bmat"))
                                    bowl_thresh = st.session_state.get("s2_wick", get_threshold_val("s2_wick"))
                                    bowl_match_thresh = st.session_state.get("s2_mmat", get_threshold_val("s2_mmat"))
                                elif "senior league 3" in name_lower or "senior 3" in name_lower:
                                    bat_thresh = st.session_state.get("s3_runs", get_threshold_val("s3_runs"))
                                    bat_match_thresh = st.session_state.get("s3_bmat", get_threshold_val("s3_bmat"))
                                    bowl_thresh = st.session_state.get("s3_wick", get_threshold_val("s3_wick"))
                                    bowl_match_thresh = st.session_state.get("s3_mmat", get_threshold_val("s3_mmat"))
                                else:
                                    j_match = re.search(r'junior league (11a|11b|11c|\d+)', name_lower)
                                    if j_match:
                                        j_id = j_match.group(1)
                                        bat_thresh = st.session_state.get(f"j{j_id}_runs", get_threshold_val(f"j{j_id}_runs"))
                                        bat_match_thresh = st.session_state.get(f"j{j_id}_bmat", get_threshold_val(f"j{j_id}_bmat"))
                                        bowl_thresh = st.session_state.get(f"j{j_id}_wick", get_threshold_val(f"j{j_id}_wick"))
                                        bowl_match_thresh = st.session_state.get(f"j{j_id}_mmat", get_threshold_val(f"j{j_id}_mmat"))
                                    else:
                                        bat_thresh = st.session_state.get("t3_runs", get_threshold_val("t3_runs"))
                                        bat_match_thresh = st.session_state.get("t3_bmat", get_threshold_val("t3_bmat"))
                                        bowl_thresh = st.session_state.get("t3_wick", get_threshold_val("t3_wick"))
                                        bowl_match_thresh = st.session_state.get("t3_mmat", get_threshold_val("t3_mmat"))

                            bat_qual_col = 'Innings'
                            bat_qual_label = 'innings'

                            if disable_thresholds: bat_thresh, bat_match_thresh, bowl_thresh, bowl_match_thresh = 0, 0, 0, 0

                            fielding_df = pd.DataFrame()
                            wk_df = pd.DataFrame()
                            if not league_bat.empty:
                                if 'Catches as Keeper' in league_bat.columns or 'Stumpings' in league_bat.columns:
                                    wk_data = league_bat.copy()
                                    wk_data['Catches as Keeper'] = pd.to_numeric(wk_data.get('Catches as Keeper', 0), errors='coerce').fillna(0).astype(int)
                                    wk_data['Stumpings'] = pd.to_numeric(wk_data.get('Stumpings', 0), errors='coerce').fillna(0).astype(int)
                                    wk_data['Total'] = wk_data['Catches as Keeper'] + wk_data['Stumpings']
                                    wk_data = wk_data[wk_data['Total'] > 0]
                                    if not wk_data.empty:
                                        wk_data = wk_data[['Player', 'Team', 'Catches as Keeper', 'Stumpings', 'Total']]
                                        wk_data.rename(columns={'Player': 'Name', 'Team': 'Club', 'Catches as Keeper': 'Catches'}, inplace=True)
                                        wk_data = wk_data.sort_values(by=['Total', 'Catches'], ascending=[False, False])
                                        wk_data.insert(0, 'Position', range(1, len(wk_data) + 1))
                                        wk_df = wk_data
                                
                                if 'Catches' in league_bat.columns:
                                    f_data = league_bat.copy()
                                    f_data['Catches'] = pd.to_numeric(f_data['Catches'], errors='coerce').fillna(0).astype(int)
                                    f_data = f_data[f_data['Catches'] > 0]
                                    if not f_data.empty:
                                        f_data = f_data[['Player', 'Team', 'Matches', 'Catches']]
                                        f_data.rename(columns={'Player': 'Name', 'Team': 'Club', 'Matches': 'M'}, inplace=True)
                                        f_data = f_data.sort_values(by=['Catches', 'M'], ascending=[False, True])
                                        f_data.insert(0, 'Position', range(1, len(f_data) + 1))
                                        fielding_df = f_data

                            if not league_bat.empty:
                                league_bat_filtered = league_bat[(league_bat['Runs'] >= bat_thresh) & (league_bat[bat_qual_col] >= bat_match_thresh)].copy()
                                if not league_bat_filtered.empty:
                                    league_bat_filtered.insert(0, 'Position', range(1, len(league_bat_filtered) + 1))
                                    drop_cols = [c for c in ['Catches', 'Catches as Keeper', 'Stumpings'] if c in league_bat_filtered.columns]
                                    league_bat_filtered = league_bat_filtered.drop(columns=drop_cols)
                                    eng.format_excel_sheet(writer, league_bat_filtered, f"{tab_prefix} Bat", min_label=f"Min {bat_thresh} runs, {bat_match_thresh} {bat_qual_label}")
                            if not league_bowl.empty:
                                league_bowl = league_bowl[(league_bowl['Wickets'] >= bowl_thresh) & (league_bowl['Innings'] >= bowl_match_thresh)]
                                if not league_bowl.empty:
                                    league_bowl.insert(0, 'Position', range(1, len(league_bowl) + 1))
                                    eng.format_excel_sheet(writer, league_bowl, f"{tab_prefix} Bowl", min_label=f"Min {bowl_thresh} wickets, {bowl_match_thresh} innings")
                            
                            if not wk_df.empty:
                                eng.format_excel_sheet(writer, wk_df, f"{tab_prefix} WK")
                            
                            if not fielding_df.empty:
                                eng.format_excel_sheet(writer, fielding_df, f"{tab_prefix} Field")
                    
                    st.success("✅ Averages calculated successfully!")
                    prefix = domain.replace("'", "")
                    file_out_name = f"{prefix}_Season_Averages_All_Leagues_{datetime.now().strftime('%Y-%m-%d_%H%M%S')}.xlsx"
                    st.download_button("📥 Download Output Excel File", data=output_buffer.getvalue(), file_name=file_out_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
                except Exception as e:
                    import traceback
                    st.error(f"An error occurred during processing: {str(e)}\n\n```\n{traceback.format_exc()}\n```")

# ==========================================
# TOOL 2: LEAGUE MILESTONES REPORT
# ==========================================
elif app_mode == "League Milestones Report":
    st.title("🏆 League Milestones Report")
    st.markdown("Generate a formatted Word document reporting all Centurions (100+ runs) and top Wicket hauls.")
    
    if not DOCX_AVAILABLE:
        st.error("The `python-docx` library is not installed. Please run `pip install python-docx` to use this feature.")
    else:
        st.subheader("Select League Domain")
        domain = st.radio("Choose the dataset domain:", ["Men's", "Women's"], horizontal=True)

        with st.sidebar:
            c_files = eng.DEFAULT_FILES[domain]
            with st.expander("📁 File Path Configurations", expanded=False):
                f_reg = st.text_input("Official Registry (Excel)", value=c_files["reg"], key=f"ms_reg_{domain}")
                f_alias = st.text_input("Aliases Master (Excel)", value=c_files["alias"], key=f"ms_alias_{domain}")
                f_league = st.text_input("League Structure (Excel)", value=c_files["league"], key=f"ms_league_{domain}")
                f_bat = st.text_input("Batting Stats (Excel)", value=c_files["bat"], key=f"ms_bat_{domain}")
                f_bowl = st.text_input("Bowling Stats (Excel)", value=c_files["bowl"], key=f"ms_bowl_{domain}")
                f_cup = st.text_input("Cup Master (Excel)", value="NCU_Cup_Fixtures.xlsx", key=f"ms_cup_{domain}")
        
        st.divider()
        if st.button("📄 Generate Milestones Word Doc", type="primary"):
            files_to_check = [f_reg, f_alias, f_league, f_bat, f_bowl]
            missing_files = [f for f in files_to_check if not os.path.exists(f)]
            
            if missing_files:
                st.error("Cannot find the following files:\n\n" + "\n".join([f"- {f}" for f in missing_files]))
            else:
                with st.spinner("Extracting milestones and building Word document..."):
                    try:
                        doc_io = eng.generate_milestones_report(domain, f_reg, f_alias, f_league, f_bat, f_bowl, f_cup)
                        domain_label = "Open" if domain == "Men's" else "Women"
                        st.success("✅ Milestones report generated successfully!")
                        st.download_button(
                            label="📥 Download Milestones Report",
                            data=doc_io.getvalue(),
                            file_name=f"Mercury_{domain_label}_Milestones_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary"
                        )
                    except Exception as e:
                        st.error(f"An error occurred: {str(e)}")