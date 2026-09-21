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
from typing import Tuple, List, Dict, Any, Optional, Set

# Import shared core engine
import engine as eng

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

    /* Metric Cards Styling & Responsive Text Sizing */
    [data-testid="stMetric"] {{
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 12px 16px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
    }}
    [data-testid="stMetricLabel"] {{
        font-size: 0.88rem !important;
        font-weight: 600 !important;
        color: #475569 !important;
        white-space: normal !important;
    }}
    [data-testid="stMetricValue"],
    [data-testid="stMetricValue"] > div,
    [data-testid="stMetricValue"] span,
    [data-testid="stMetricValue"] * {{
        font-size: 1.25rem !important;
        font-weight: 700 !important;
        color: #1F4E78 !important;
        white-space: normal !important;
        word-break: break-word !important;
        text-overflow: unset !important;
        overflow: visible !important;
        line-height: 1.3 !important;
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
    for st.dataframe across the stats application.
    Centered columns have center-balanced header labels mirroring centered cells,
    while left-aligned columns remain strictly left-aligned.
    """
    return {
        "Rank": st.column_config.Column(center_col_label("Rank", 10), alignment="center", width="small"),
        "Player": st.column_config.TextColumn("Player", alignment="left", width="medium"),
        "Bowler": st.column_config.TextColumn("Bowler", alignment="left", width="medium"),
        "Club": st.column_config.TextColumn("Club Name", alignment="left", width="medium"),
        "Innings": st.column_config.NumberColumn(center_col_label("Innings", 12), format="%d", alignment="center", width="small"),
        "Not Outs": st.column_config.NumberColumn(center_col_label("Not Outs", 12), format="%d", alignment="center", width="small"),
        "Runs": st.column_config.NumberColumn(center_col_label("Runs", 10), format="%d", alignment="center", width="small"),
        "High Score": st.column_config.Column(center_col_label("High Score", 14), alignment="center", width="small"),
        "Average": st.column_config.Column(center_col_label("Average", 12), alignment="center", width="small"),
        "50s": st.column_config.NumberColumn(center_col_label("50s", 10), format="%d", alignment="center", width="small"),
        "100s": st.column_config.NumberColumn(center_col_label("100s", 10), format="%d", alignment="center", width="small"),
        "Overs": st.column_config.Column(center_col_label("Overs", 10), alignment="center", width="small"),
        "Maidens": st.column_config.NumberColumn(center_col_label("Maidens", 12), format="%d", alignment="center", width="small"),
        "Wickets": st.column_config.NumberColumn(center_col_label("Wickets", 12), format="%d", alignment="center", width="small"),
        "Economy": st.column_config.Column(center_col_label("Economy", 12), alignment="center", width="small"),
        "Best Bowling": st.column_config.Column(center_col_label("Best Bowling", 16), alignment="center", width="medium"),
        "Matches": st.column_config.NumberColumn(center_col_label("Matches", 12), format="%d", alignment="center", width="small"),
        "Catches": st.column_config.NumberColumn(center_col_label("Catches", 12), format="%d", alignment="center", width="small"),
        "Stumpings": st.column_config.NumberColumn(center_col_label("Stumpings", 14), format="%d", alignment="center", width="small"),
        "Run Outs": st.column_config.NumberColumn(center_col_label("Run Outs", 12), format="%d", alignment="center", width="small"),
        "Total Dismissals": st.column_config.NumberColumn(center_col_label("Total Dismissals", 18), format="%d", alignment="center", width="medium"),
    }

# Sidebar Navigation
with st.sidebar:
    st.title("🏏 NCU Stats Hub")
    if os.environ.get("TEST_MODE", "0") == "1":
        st.warning("⚠️ **TEST MODE**\n\nUsing sample data from `test_data/`. No production files are being read or written.")
    st.header("🛠️ Navigation")
    app_mode = st.radio("Choose a module to run:", ["2026 Season Summary Dashboard", "Bulk Averages Calculator", "League Milestones Report"])
    st.divider()

# ==========================================
# TOOL 1: BULK AVERAGES
# ==========================================
# Competition filtering routines migrated to engine.py
get_cup_and_t20_match_sets = eng.get_cup_and_t20_match_sets
classify_match_type = eng.classify_match_type
filter_match_formats = eng.filter_match_formats

@st.cache_data(show_spinner="Computing season averages...")
def compute_season_averages_cached(domain, include_irish, include_cup, include_t20, include_pathway, bat_sort_pref, bowl_sort_pref, file_signatures):
    f_reg, f_alias, f_id_map, f_league, f_bat, f_bowl, f_irish_bat, f_irish_bowl, f_cup, f_unreg, f_secondary = [fs[0] for fs in file_signatures]
    
    reg_players = get_excel_df(f_reg)
    aliases = get_excel_df(f_alias)
    id_map_df = get_excel_df(f_id_map) if f_id_map and os.path.exists(f_id_map) else None
    id_map = eng.build_id_map(id_map_df)
    league_structure = get_excel_df(f_league)
    batting = get_excel_df(f_bat).copy()
    bowling = get_excel_df(f_bowl).copy()
    
    if domain == "Men's" and include_irish:
        if f_irish_bat and os.path.exists(f_irish_bat): batting = pd.concat([batting, get_excel_df(f_irish_bat)], ignore_index=True)
        if f_irish_bowl and os.path.exists(f_irish_bowl): bowling = pd.concat([bowling, get_excel_df(f_irish_bowl)], ignore_index=True)

    if domain in ["Men's", "Women's"]:
        batting, bowling = filter_match_formats(
            batting, bowling, f_cup, domain, include_cup, include_t20, include_pathway
        )

    unreg_df = get_excel_df(f_unreg) if f_unreg and os.path.exists(f_unreg) else None
    sec_df = get_excel_df(f_secondary) if f_secondary and os.path.exists(f_secondary) else None
    
    alias_map = eng.build_alias_map(aliases, domain)
    secondary_map = eng.build_secondary_team_map(sec_df, alias_map) 
    league_dict, team_keys, original_league_order = eng.build_league_dict(league_structure)
    player_club_map = eng.build_player_club_map(reg_players, alias_map, domain, unreg_map_df=unreg_df, id_map_df=id_map_df, secondary_map=secondary_map)
    player_club_map = eng.infer_unregistered_player_clubs(batting, bowling, player_club_map, min_matches=2)
    
    bat_res = batting.apply(lambda r: eng.resolve_player_from_row(r, r['Name'], id_map, alias_map, player_club_map, prefer_nv_play_name=True), axis=1)
    batting['Cleaned Name'] = [res[0] for res in bat_res]
    batting['Sport80_ID'] = [res[1] for res in bat_res]

    bowl_res = bowling.apply(lambda r: eng.resolve_player_from_row(r, r['Bowler'], id_map, alias_map, player_club_map, prefer_nv_play_name=True), axis=1)
    bowling['Cleaned Name'] = [res[0] for res in bowl_res]
    bowling['Sport80_ID'] = [res[1] for res in bowl_res]
    
    intra_team_map = eng.get_cached_intra_club_map()
    
    batting_avgs, bowling_avgs = eng.calculate_averages(
        batting, bowling, player_club_map, team_keys, league_dict, domain,
        bat_sort_pref, bowl_sort_pref, secondary_map=secondary_map,
        alias_map=alias_map, intra_team_map=intra_team_map
    )
    return batting_avgs, bowling_avgs, original_league_order


@st.cache_data(show_spinner="Loading and preparing club statistics...")
def load_club_summary_data(domain: str, file_signatures: tuple) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Loads, cleans, and pre-resolves batting and bowling records for the Club Summary dashboard tab.

    Inputs:
        domain: Competition domain ("Men's", "Women's", or "Midweek").
        file_signatures: Tuple of (filepath, modification_time) pairs to guarantee automatic cache invalidation.

    Outputs:
        Tuple[pd.DataFrame, pd.DataFrame]: (resolved_batting_df, resolved_bowling_df) with 'Cleaned Name',
            'Sport80_ID', 'Team', 'Club', and 'Match_Type' columns populated.
    """
    f_reg, f_alias, f_id_map, f_unreg, f_secondary, f_bat, f_bowl, f_cup = [fs[0] for fs in file_signatures]

    reg_players = get_excel_df(f_reg) if f_reg and os.path.exists(f_reg) else pd.DataFrame()
    aliases = get_excel_df(f_alias) if f_alias and os.path.exists(f_alias) else pd.DataFrame()
    id_map_df = get_excel_df(f_id_map) if f_id_map and os.path.exists(f_id_map) else None
    id_map = eng.build_id_map(id_map_df)
    unreg_df = get_excel_df(f_unreg) if f_unreg and os.path.exists(f_unreg) else None
    sec_df = get_excel_df(f_secondary) if f_secondary and os.path.exists(f_secondary) else None

    alias_map = eng.build_alias_map(aliases, domain)
    sec_map = eng.build_secondary_team_map(sec_df, alias_map)
    pcm = eng.build_player_club_map(reg_players, alias_map, domain, unreg_map_df=unreg_df, id_map_df=id_map_df, secondary_map=sec_map)

    batting = get_excel_df(f_bat).copy() if f_bat and os.path.exists(f_bat) else pd.DataFrame()
    bowling = get_excel_df(f_bowl).copy() if f_bowl and os.path.exists(f_bowl) else pd.DataFrame()

    cup_match_set, t20_match_set = get_cup_and_t20_match_sets(f_cup, domain)
    intra_map = eng.get_cached_intra_club_map()

    if not batting.empty:
        for col in ['Matches', 'Innings', 'Not Outs', 'Runs', 'Balls', 'Fours', 'Sixes', '50s', '100s', 'Catches', 'Catches as Keeper', 'Stumpings', 'Run Outs']:
            if col in batting.columns:
                batting[col] = pd.to_numeric(batting[col], errors='coerce').fillna(0)
        bat_res = batting.apply(lambda r: eng.resolve_player_from_row(r, r['Name'], id_map, alias_map, pcm, prefer_nv_play_name=True), axis=1)
        batting['Cleaned Name'] = [res[0] for res in bat_res]
        batting['Sport80_ID'] = [res[1] for res in bat_res]
        batting['Team'] = batting.apply(lambda r: eng.determine_player_team_for_row(r, pcm, domain, sec_map, alias_map=alias_map, intra_team_map=intra_map), axis=1)
        batting['Club'] = batting['Team'].apply(eng.extract_base_club_name)
        batting['Match_Type'] = batting['Group'].apply(lambda g: classify_match_type(g, cup_match_set, t20_match_set, domain))

    if not bowling.empty:
        for col in ['Matches', 'Innings', 'Overs', 'Maidens', 'Runs', 'Wickets', 'Balls']:
            if col in bowling.columns:
                bowling[col] = pd.to_numeric(bowling[col], errors='coerce').fillna(0)
        bowl_res = bowling.apply(lambda r: eng.resolve_player_from_row(r, r['Bowler'], id_map, alias_map, pcm, prefer_nv_play_name=True), axis=1)
        bowling['Cleaned Name'] = [res[0] for res in bowl_res]
        bowling['Sport80_ID'] = [res[1] for res in bowl_res]
        bowling['Team'] = bowling.apply(lambda r: eng.determine_player_team_for_row(r, pcm, domain, sec_map, alias_map=alias_map, intra_team_map=intra_map), axis=1)
        bowling['Club'] = bowling['Team'].apply(eng.extract_base_club_name)
        bowling['Match_Type'] = bowling['Group'].apply(lambda g: classify_match_type(g, cup_match_set, t20_match_set, domain))

    return batting, bowling


def render_club_averages_summary_tab() -> None:
    """
    Renders the interactive Club Averages & Summary dashboard tab.

    Features:
        - Domain radio selector (Men's, Women's, Midweek).
        - Dynamic competition selectbox based on active domain.
        - Dynamic club selectbox isolating clubs for active domain.
        - Three live summary metric cards (Top Individual Score, Best Bowling Analysis, Total Active Roster).
        - Parallel Top 5 Batsmen and Top 5 Bowlers leaderboards filtered dynamically in-memory.
    """
    st.markdown("""
    <div style="background: linear-gradient(135deg, #1F4E78 0%, #15375B 100%); color: #FFFFFF; padding: 14px 20px; border-radius: 8px; margin-bottom: 16px; border-left: 6px solid #D4AF37;">
        <h3 style="margin: 0; color: #FFFFFF; font-size: 1.35rem; font-weight: 700;">🏏 Club Averages & Performance Summary</h3>
        <p style="margin: 4px 0 0 0; color: #E0E7FF; font-size: 0.9rem;">Standalone performance metrics, records, and club leaderboards filtered dynamically in-memory</p>
    </div>
    """, unsafe_allow_html=True)

    # 1. Domain Selector
    active_domain = st.radio(
        "Active Domain:",
        options=["Men's", "Women's", "Midweek"],
        index=0,
        horizontal=True,
        key="club_summary_domain"
    )

    # 2. Dynamic Competition Format Options
    if active_domain == "Men's":
        comp_options = ["Combined", "League", "Cup", "T20"]
    elif active_domain == "Women's":
        comp_options = ["Combined", "League", "Cup"]
    else:  # Midweek
        comp_options = ["Midweek League"]

    # 3. Dynamic Club Options for Active Domain
    club_team_counts = eng.get_all_club_team_counts()
    if active_domain == "Men's":
        club_list = sorted(list(eng.NCU_ALL_37_CLUBS))
    elif active_domain == "Women's":
        club_list = sorted([c for c, d in club_team_counts.items() if d.get('women', 0) > 0])
    else:  # Midweek
        club_list = sorted([c for c, d in club_team_counts.items() if d.get('midweek', 0) > 0])

    if not club_list:
        club_list = sorted(list(eng.NCU_ALL_37_CLUBS))

    club_options = ["All Clubs"] + club_list

    # Controls Row
    c_left, c_right = st.columns(2)
    with c_left:
        selected_comp = st.selectbox(
            "Filter Competition:",
            options=comp_options,
            index=0,
            key=f"club_summary_comp_{active_domain}"
        )
    with c_right:
        selected_club = st.selectbox(
            "Select Club:",
            options=club_options,
            index=0,
            key=f"club_summary_club_{active_domain}"
        )

    # 4. Load Cached Performance Data
    c_files = eng.DEFAULT_FILES.get(active_domain, eng.DEFAULT_FILES["Men's"])
    f_reg = c_files.get("reg", "")
    f_alias = c_files.get("alias", "")
    f_id_map = c_files.get("id_map", "")
    f_unreg = c_files.get("unreg", "")
    f_secondary = c_files.get("secondary", "")
    f_bat = c_files.get("bat", "")
    f_bowl = c_files.get("bowl", "")
    f_cup = c_files.get("cup", "")

    file_list = [f_reg, f_alias, f_id_map, f_unreg, f_secondary, f_bat, f_bowl, f_cup]
    file_signatures = tuple(
        (f, os.path.getmtime(f) if (f and os.path.exists(f)) else 0)
        for f in file_list
    )

    batting_df, bowling_df = load_club_summary_data(active_domain, file_signatures)

    # 5. In-Memory Data Filtering Rule
    is_all_clubs = (selected_club == "All Clubs")
    top_limit = 20 if is_all_clubs else 10
    club_label = "All Clubs" if is_all_clubs else selected_club

    if is_all_clubs:
        c_bat = batting_df.copy() if not batting_df.empty else pd.DataFrame()
        c_bowl = bowling_df.copy() if not bowling_df.empty else pd.DataFrame()
    else:
        c_bat = batting_df[batting_df['Club'] == selected_club].copy() if not batting_df.empty and 'Club' in batting_df.columns else pd.DataFrame()
        c_bowl = bowling_df[bowling_df['Club'] == selected_club].copy() if not bowling_df.empty and 'Club' in bowling_df.columns else pd.DataFrame()

    if selected_comp != "Combined":
        if not c_bat.empty and 'Match_Type' in c_bat.columns:
            c_bat = c_bat[c_bat['Match_Type'] == selected_comp]
        if not c_bowl.empty and 'Match_Type' in c_bowl.columns:
            c_bowl = c_bowl[c_bowl['Match_Type'] == selected_comp]

    # 6. Live Summary Metric Cards
    is_midweek = (active_domain == "Midweek")
    bat_metric_label = "Most 30+ Innings" if is_midweek else "Top Individual Score"
    top_score_val = "N/A"
    top_score_help = "Most individual innings reaching the 30-run retirement threshold in this competition." if is_midweek else "Highest individual innings score in this competition."
    if not c_bat.empty:
        c_bat['HS_Num'] = c_bat['High Score'].apply(eng.parse_high_score_numeric)
        c_bat['Runs_Numeric'] = pd.to_numeric(c_bat['Runs'], errors='coerce').fillna(0)

        if is_midweek:
            c_bat['Is_30_Plus'] = (c_bat['Runs_Numeric'] >= 30).astype(int)
            p_30_agg = c_bat.groupby('Cleaned Name').agg(
                Count_30=('Is_30_Plus', 'sum'),
                Total_Runs=('Runs_Numeric', 'sum')
            ).reset_index()

            player_primary_club = c_bat.groupby('Cleaned Name')['Club'].first().to_dict()
            top_30_players = p_30_agg.sort_values(by=['Count_30', 'Total_Runs'], ascending=[False, False])
            if not top_30_players.empty:
                top_row = top_30_players.iloc[0]
                cnt = int(top_row['Count_30'])
                p_name = str(top_row['Cleaned Name'])
                p_club = player_primary_club.get(p_name, '')
                if cnt > 0:
                    top_score_val = f"{p_name} ({cnt} x 30+)"
                    if is_all_clubs and p_club:
                        top_score_help = f"Most individual innings reaching the 30-run retirement threshold in this competition: {p_name} ({p_club}) - {cnt} times."
                    else:
                        top_score_help = f"Most individual innings reaching the 30-run retirement threshold recorded by a batsman from {selected_club}: {p_name} - {cnt} times."
                else:
                    top_score_val = "None (0 x 30+)"
                    top_score_help = f"No batsman from {selected_club} reached the 30-run retirement threshold in this competition."
        else:
            top_bat_row = c_bat.sort_values(by=['HS_Num', 'Runs_Numeric'], ascending=[False, False]).iloc[0]
            hs_raw = top_bat_row.get('High Score')
            if pd.isna(hs_raw) or str(hs_raw).lower() == 'nan':
                hs_num = top_bat_row.get('HS_Num', 0)
                hs = str(int(hs_num) if (pd.notna(hs_num) and not pd.isna(hs_num)) else 0)
            else:
                hs = str(hs_raw)

            p_name = str(top_bat_row.get('Cleaned Name', top_bat_row.get('Name', 'Unknown')))
            top_score_val = f"{p_name} ({hs})"
            p_club = str(top_bat_row.get('Club', ''))
            if is_all_clubs and p_club:
                top_score_help = f"Highest individual innings score in this competition: {p_name} ({p_club}) - {hs}."
            else:
                top_score_help = f"Highest individual innings score recorded by a batsman from {selected_club}."

    best_bowl_val = "N/A"
    best_bowl_help = "Best single-innings bowling performance in this competition."
    if not c_bowl.empty:
        parsed_bb = c_bowl['Best Bowling in an Innings'].apply(eng.bb_sort_key)
        c_bowl['BB_Wickets'] = [p[0] for p in parsed_bb]
        c_bowl['BB_Runs'] = [-p[1] for p in parsed_bb]
        c_bowl['Wickets_Num'] = pd.to_numeric(c_bowl['Wickets'], errors='coerce').fillna(0)
        c_bowl['Runs_Num'] = pd.to_numeric(c_bowl['Runs'], errors='coerce').fillna(0)
        top_bowl_row = c_bowl.sort_values(by=['BB_Wickets', 'BB_Runs', 'Wickets_Num'], ascending=[False, True, False]).iloc[0]
        bb_raw = top_bowl_row.get('Best Bowling in an Innings')
        if pd.isna(bb_raw) or str(bb_raw).lower() == 'nan':
            w_num = top_bowl_row.get('BB_Wickets', 0)
            r_num = top_bowl_row.get('BB_Runs', 0)
            w_int = int(w_num) if (pd.notna(w_num) and not pd.isna(w_num)) else 0
            r_int = int(r_num) if (pd.notna(r_num) and not pd.isna(r_num)) else 0
            bb_fig = f"{w_int}-{r_int}"
        else:
            bb_fig = str(bb_raw)

        b_name = str(top_bowl_row.get('Cleaned Name', top_bowl_row.get('Bowler', 'Unknown')))
        best_bowl_val = f"{b_name} ({bb_fig})"
        b_club = str(top_bowl_row.get('Club', ''))
        if is_all_clubs and b_club:
            best_bowl_help = f"Best single-innings bowling performance in this competition: {b_name} ({b_club}) - {bb_fig}."
        else:
            best_bowl_help = f"Best single-innings bowling performance (wickets-runs) achieved by a bowler from {selected_club}."

    bat_players = set(c_bat['Cleaned Name'].dropna().unique()) if not c_bat.empty and 'Cleaned Name' in c_bat.columns else set()
    bowl_players = set(c_bowl['Cleaned Name'].dropna().unique()) if not c_bowl.empty and 'Cleaned Name' in c_bowl.columns else set()
    active_roster = len(bat_players.union(bowl_players))
    roster_help = "Total unique players across all clubs who made at least one appearance in this competition." if is_all_clubs else f"Unique players who made at least one batting or bowling appearance for {selected_club}."

    st.markdown("<br>", unsafe_allow_html=True)
    m1, m2, m3 = st.columns(3)
    with m1:
        st.metric(
            label=bat_metric_label,
            value=top_score_val,
            help=top_score_help
        )
    with m2:
        st.metric(
            label="Best Bowling Analysis",
            value=best_bowl_val,
            help=best_bowl_help
        )
    with m3:
        st.metric(
            label="Total Active Roster",
            value=str(active_roster),
            help=roster_help
        )

    st.divider()

    # 7. Batting Leaderboard
    st.markdown(f"#### 🏏 Top {top_limit} Batsmen ({club_label})")
    if not c_bat.empty:
        for col in ['Innings', 'Not Outs', 'Runs', 'Balls', '50s', '100s']:
            if col in c_bat.columns:
                c_bat[col] = pd.to_numeric(c_bat[col], errors='coerce').fillna(0)

        # Pre-compute high score string and primary club per player
        high_scores = {}
        player_clubs = {}
        for p_name, p_grp in c_bat.groupby('Cleaned Name'):
            top_r = p_grp.sort_values(by=['HS_Num', 'Runs_Numeric'], ascending=[False, False]).iloc[0]
            hs_val = top_r.get('High Score')
            if pd.isna(hs_val) or str(hs_val).lower() == 'nan':
                hs_num = top_r.get('HS_Num', 0)
                hs_val = int(hs_num) if (pd.notna(hs_num) and not pd.isna(hs_num)) else 0
            high_scores[p_name] = str(hs_val)
            player_clubs[p_name] = str(top_r.get('Club', '-'))

        bat_agg = c_bat.groupby('Cleaned Name').agg({
            'Innings': 'sum',
            'Not Outs': 'sum',
            'Runs': 'sum',
            'Balls': 'sum',
            '50s': 'sum',
            '100s': 'sum'
        }).reset_index()

        bat_agg['Average'] = bat_agg.apply(
            lambda r: eng.calculate_batting_average(r['Runs'], r['Innings'], r['Not Outs']),
            axis=1
        )
        bat_agg['High Score'] = bat_agg['Cleaned Name'].map(high_scores).fillna("-")
        bat_agg['Club'] = bat_agg['Cleaned Name'].map(player_clubs).fillna("-")
        top_batsmen = bat_agg.sort_values(by=['Runs'], ascending=[False]).head(top_limit).reset_index(drop=True)
        top_batsmen.insert(0, 'Rank', range(1, len(top_batsmen) + 1))
        top_batsmen.rename(columns={'Cleaned Name': 'Player'}, inplace=True)

        if is_all_clubs:
            disp_cols_bat = ['Rank', 'Player', 'Club', 'Innings', 'Not Outs', 'Runs', 'High Score', 'Average', '50s', '100s']
        else:
            disp_cols_bat = ['Rank', 'Player', 'Innings', 'Not Outs', 'Runs', 'High Score', 'Average', '50s', '100s']
        st.dataframe(
            top_batsmen[disp_cols_bat],
            width="stretch",
            hide_index=True,
            column_config=get_standard_column_config()
        )
    else:
        st.info(f"No batting records found for {club_label} ({selected_comp}).")

    st.divider()

    # 8. Bowling Leaderboard
    st.markdown(f"#### 🎯 Top {top_limit} Bowlers ({club_label})")
    if not c_bowl.empty:
        for col in ['Innings', 'Balls', 'Maidens', 'Runs', 'Wickets']:
            if col in c_bowl.columns:
                c_bowl[col] = pd.to_numeric(c_bowl[col], errors='coerce').fillna(0)

        # Pre-compute best bowling analysis and primary club per player
        best_figs = {}
        bowler_clubs = {}
        for p_name, p_grp in c_bowl.groupby('Cleaned Name'):
            top_b = p_grp.sort_values(by=['BB_Wickets', 'BB_Runs', 'Wickets_Num'], ascending=[False, True, False]).iloc[0]
            bb_val = top_b.get('Best Bowling in an Innings')
            if pd.isna(bb_val) or str(bb_val).lower() == 'nan':
                w_num = top_b.get('BB_Wickets', 0)
                r_num = top_b.get('BB_Runs', 0)
                w_int = int(w_num) if (pd.notna(w_num) and not pd.isna(w_num)) else 0
                r_int = int(r_num) if (pd.notna(r_num) and not pd.isna(r_num)) else 0
                bb_val = f"{w_int}-{r_int}"
            best_figs[p_name] = str(bb_val)
            bowler_clubs[p_name] = str(top_b.get('Club', '-'))

        bowl_agg = c_bowl.groupby('Cleaned Name').agg({
            'Innings': 'sum',
            'Balls': 'sum',
            'Maidens': 'sum',
            'Runs': 'sum',
            'Wickets': 'sum'
        }).reset_index()

        bowl_agg['Overs'] = bowl_agg['Balls'].apply(lambda b: f"{int(b // 6)}.{int(b % 6)}")
        bowl_agg['Average'] = bowl_agg.apply(
            lambda r: eng.calculate_bowling_average(r['Runs'], r['Wickets']),
            axis=1
        )
        bowl_agg['Economy'] = bowl_agg.apply(
            lambda r: eng.calculate_economy_rate(r['Runs'], r['Balls']),
            axis=1
        )
        bowl_agg['Best Bowling'] = bowl_agg['Cleaned Name'].map(best_figs).fillna("-")
        bowl_agg['Club'] = bowl_agg['Cleaned Name'].map(bowler_clubs).fillna("-")
        top_bowlers = bowl_agg.sort_values(by=['Wickets', 'Runs'], ascending=[False, True]).head(top_limit).reset_index(drop=True)
        top_bowlers.insert(0, 'Rank', range(1, len(top_bowlers) + 1))
        top_bowlers.rename(columns={'Cleaned Name': 'Bowler'}, inplace=True)

        if is_all_clubs:
            disp_cols_bowl = ['Rank', 'Bowler', 'Club', 'Innings', 'Overs', 'Maidens', 'Runs', 'Wickets', 'Average', 'Economy', 'Best Bowling']
        else:
            disp_cols_bowl = ['Rank', 'Bowler', 'Innings', 'Overs', 'Maidens', 'Runs', 'Wickets', 'Average', 'Economy', 'Best Bowling']
        st.dataframe(
            top_bowlers[disp_cols_bowl],
            width="stretch",
            hide_index=True,
            column_config=get_standard_column_config()
        )
    else:
        st.info(f"No bowling records found for {club_label} ({selected_comp}).")

    st.divider()

    # 9. Wicket Keeper Leaderboard
    st.markdown(f"#### 🧤 Top {top_limit} Wicket Keepers ({club_label})")
    if not c_bat.empty:
        for col in ['Matches', 'Catches as Keeper', 'Stumpings']:
            if col in c_bat.columns:
                c_bat[col] = pd.to_numeric(c_bat[col], errors='coerce').fillna(0)
            else:
                c_bat[col] = 0

        player_clubs_wk = {p_name: str(p_grp.iloc[0].get('Club', '-')) for p_name, p_grp in c_bat.groupby('Cleaned Name')}

        wk_agg = c_bat.groupby('Cleaned Name').agg({
            'Matches': 'sum',
            'Catches as Keeper': 'sum',
            'Stumpings': 'sum'
        }).reset_index()
        wk_agg['Total Dismissals'] = (wk_agg['Catches as Keeper'] + wk_agg['Stumpings']).astype(int)
        wk_agg['Catches as Keeper'] = wk_agg['Catches as Keeper'].astype(int)
        wk_agg['Stumpings'] = wk_agg['Stumpings'].astype(int)
        wk_agg['Matches'] = wk_agg['Matches'].astype(int)
        wk_agg['Club'] = wk_agg['Cleaned Name'].map(player_clubs_wk).fillna("-")

        qual_wk = wk_agg[wk_agg['Total Dismissals'] > 0].copy()
        if not qual_wk.empty:
            top_wk = qual_wk.sort_values(by=['Total Dismissals', 'Stumpings', 'Catches as Keeper'], ascending=[False, False, False]).head(top_limit).reset_index(drop=True)
            top_wk.insert(0, 'Rank', range(1, len(top_wk) + 1))
            top_wk.rename(columns={'Cleaned Name': 'Player', 'Catches as Keeper': 'Catches'}, inplace=True)
            if is_all_clubs:
                disp_cols_wk = ['Rank', 'Player', 'Club', 'Matches', 'Catches', 'Stumpings', 'Total Dismissals']
            else:
                disp_cols_wk = ['Rank', 'Player', 'Matches', 'Catches', 'Stumpings', 'Total Dismissals']
            st.dataframe(
                top_wk[disp_cols_wk],
                width="stretch",
                hide_index=True,
                column_config=get_standard_column_config()
            )
        else:
            st.info(f"No wicket keeping dismissals recorded for {club_label} ({selected_comp}).")
    else:
        st.info(f"No wicket keeping records found for {club_label} ({selected_comp}).")

    st.divider()

    # 10. Fielding Leaderboard
    st.markdown(f"#### 🛡️ Top {top_limit} Fielders ({club_label})")
    if not c_bat.empty:
        for col in ['Matches', 'Catches', 'Run Outs']:
            if col in c_bat.columns:
                c_bat[col] = pd.to_numeric(c_bat[col], errors='coerce').fillna(0)
            else:
                c_bat[col] = 0

        player_clubs_fld = {p_name: str(p_grp.iloc[0].get('Club', '-')) for p_name, p_grp in c_bat.groupby('Cleaned Name')}

        fld_agg = c_bat.groupby('Cleaned Name').agg({
            'Matches': 'sum',
            'Catches': 'sum',
            'Run Outs': 'sum'
        }).reset_index()
        fld_agg['Total Dismissals'] = (fld_agg['Catches'] + fld_agg['Run Outs']).astype(int)
        fld_agg['Catches'] = fld_agg['Catches'].astype(int)
        fld_agg['Run Outs'] = fld_agg['Run Outs'].astype(int)
        fld_agg['Matches'] = fld_agg['Matches'].astype(int)
        fld_agg['Club'] = fld_agg['Cleaned Name'].map(player_clubs_fld).fillna("-")

        qual_fld = fld_agg[fld_agg['Total Dismissals'] > 0].copy()
        if not qual_fld.empty:
            top_fld = qual_fld.sort_values(by=['Total Dismissals', 'Catches', 'Run Outs'], ascending=[False, False, False]).head(top_limit).reset_index(drop=True)
            top_fld.insert(0, 'Rank', range(1, len(top_fld) + 1))
            top_fld.rename(columns={'Cleaned Name': 'Player'}, inplace=True)
            if is_all_clubs:
                disp_cols_fld = ['Rank', 'Player', 'Club', 'Matches', 'Catches', 'Run Outs', 'Total Dismissals']
            else:
                disp_cols_fld = ['Rank', 'Player', 'Matches', 'Catches', 'Run Outs', 'Total Dismissals']
            st.dataframe(
                top_fld[disp_cols_fld],
                width="stretch",
                hide_index=True,
                column_config=get_standard_column_config()
            )
        else:
            st.info(f"No fielding dismissals recorded for {club_label} ({selected_comp}).")
    else:
        st.info(f"No fielding records found for {club_label} ({selected_comp}).")


if app_mode == "Bulk Averages Calculator":
    init_threshold_store()
    st.title("📊 League Bulk Averages Calculator")
    
    col_dom, col_save, col_reset = st.columns([2, 1, 1])
    with col_dom:
        domain = st.radio("League Domain:", ["Men's", "Women's", "Midweek"], horizontal=True, label_visibility="collapsed")
    with col_save:
        if st.button("💾 Save Thresholds", width="stretch"):
            save_threshold_settings()
            st.toast("Saved custom thresholds as new defaults!", icon="✅")
    with col_reset:
        if st.button("🔄 Reset Defaults", width="stretch"):
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
            f_id_map = st.text_input("ID Mapping Master (Excel)", value=c_files.get("id_map", ""), key=f"avg_id_map_{domain}")
            f_unreg = st.text_input("Unregistered Players Map", value=c_files.get("unreg", ""), key=f"avg_unreg_{domain}")
            f_secondary = st.text_input("Secondary Team Map", value=c_files.get("secondary", ""), key=f"avg_sec_{domain}") 
            f_league = st.text_input("League Structure (Excel)", value=c_files["league"], key=f"avg_league_{domain}")
            f_bat = st.text_input("Batting Stats (Excel)", value=c_files["bat"], key=f"avg_bat_{domain}")
            f_bowl = st.text_input("Bowling Stats (Excel)", value=c_files["bowl"], key=f"avg_bowl_{domain}")
            f_cup = st.text_input("Cup Master (Excel)", value=c_files.get("cup", eng.DEFAULT_CUP_FILE), key=f"avg_cup_{domain}")
            if include_irish:
                f_irish_bat = st.text_input("Irish Batting Stats (Excel)", value=c_files.get("irish_bat", eng.DEFAULT_IRISH_BAT_FILE), key="avg_irish_bat")
                f_irish_bowl = st.text_input("Irish Bowling Stats (Excel)", value=c_files.get("irish_bowl", eng.DEFAULT_IRISH_BOWL_FILE), key="avg_irish_bowl")

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚀 Process Averages", type="primary"):
        files_to_check = [f_reg, f_alias, f_league, f_bat, f_bowl, f_cup]
        if f_id_map: files_to_check.append(f_id_map)
        if domain == "Men's" and include_irish: files_to_check.extend([f_irish_bat, f_irish_bowl])
            
        missing_files = [f for f in files_to_check if not os.path.exists(f)]
        if missing_files:
            st.error(f"Cannot find the following files:\n\n" + "\n".join([f"- {f}" for f in missing_files]))
        else:
            with st.spinner(f"Running {domain} Averages Engine..."):
                try:
                    irish_bat_path = f_irish_bat if (domain == "Men's" and include_irish) else None
                    irish_bowl_path = f_irish_bowl if (domain == "Men's" and include_irish) else None
                    file_list = [f_reg, f_alias, f_id_map, f_league, f_bat, f_bowl, irish_bat_path, irish_bowl_path, f_cup, f_unreg, f_secondary]
                    file_signatures = tuple(
                        (f, os.path.getmtime(f) if (f and os.path.exists(f)) else 0)
                        for f in file_list
                    )
                    batting_avgs, bowling_avgs, original_league_order = compute_season_averages_cached(
                        domain, include_irish, include_cup, include_t20, include_pathway,
                        bat_sort_pref, bowl_sort_pref, file_signatures
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
                f_id_map = st.text_input("ID Mapping Master (Excel)", value=c_files.get("id_map", ""), key=f"ms_id_map_{domain}")
                f_league = st.text_input("League Structure (Excel)", value=c_files["league"], key=f"ms_league_{domain}")
                f_bat = st.text_input("Batting Stats (Excel)", value=c_files["bat"], key=f"ms_bat_{domain}")
                f_bowl = st.text_input("Bowling Stats (Excel)", value=c_files["bowl"], key=f"ms_bowl_{domain}")
                f_cup = st.text_input("Cup Master (Excel)", value=c_files.get("cup", eng.DEFAULT_CUP_FILE), key=f"ms_cup_{domain}")
                f_secondary = st.text_input("Secondary Team Map (Excel)", value=c_files.get("secondary", os.path.join("test_data", "5. Secondary_Team_Map.xlsx") if eng._TEST_MODE else "5. Secondary_Team_Map.xlsx"), key=f"ms_secondary_{domain}")
        
        st.divider()
        if st.button("📄 Generate Milestones Word Doc", type="primary"):
            files_to_check = [f_reg, f_alias, f_league, f_bat, f_bowl]
            if f_id_map: files_to_check.append(f_id_map)
            missing_files = [f for f in files_to_check if not os.path.exists(f)]
            
            if missing_files:
                st.error("Cannot find the following files:\n\n" + "\n".join([f"- {f}" for f in missing_files]))
            else:
                with st.spinner("Extracting milestones and building Word document..."):
                    try:
                        doc_io = eng.generate_milestones_report(domain, f_reg, f_alias, f_league, f_bat, f_bowl, f_cup, f_id_map=f_id_map, f_secondary=f_secondary)
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

# ==========================================
# TOOL 3: 2026 SEASON SUMMARY DASHBOARD
# ==========================================
elif app_mode == "2026 Season Summary Dashboard":
    tab_overview, tab_club = st.tabs(["📊 Season Summary", "🏏 Club Averages & Summary"])
    with tab_overview:
        eng.render_season_summary_dashboard()
    with tab_club:
        render_club_averages_summary_tab()