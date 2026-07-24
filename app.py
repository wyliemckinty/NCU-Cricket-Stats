# ==========================================
# app.py
# ==========================================
import streamlit as st
import pandas as pd
import os
import io
import zipfile
from datetime import datetime, timedelta

import engine as eng

try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False
    
import warnings
warnings.filterwarnings('ignore')

# ==========================================
# USER CONFIGURATIONS
# ==========================================
MAIN_HEADER_SIZE = "28px" 

PAGE_TITLES = {
    "bulk_averages": "📊 League Bulk Averages Calculator",
    "player_doc": "📄 Player Word Doc Generator",
    "reg_checks": "🛡️ Weekend Registration and Starring Checks",
    "midweek_checks": "🛡️ Midweek Registration & Starring Check",
    "starring_reports": "🚨 Club Starring & Inactivity Exporter",
    "fines_generator": "💸 Club Fines Generator",
    "unregistered_fines": "💸 Unregistered Player Fines Generator"
}

# ==========================================
# PAGE CONFIGURATION & CUSTOM CSS STYLING
# ==========================================
st.set_page_config(page_title="NCU Cricket Hub", page_icon="🏏", layout="wide")

st.markdown(f"""
<style>
    h1 {{
        font-size: {MAIN_HEADER_SIZE} !important;
        font-weight: 700;
    }}
    div.stButton > button[kind="primary"] {{
        border-radius: 8px;
        padding: 0.5rem 1.5rem;
        transition: background-color 0.3s ease;
    }}
    div.stDownloadButton > button:first-child {{
        background-color: #0066cc;
        color: white;
        border-radius: 8px;
        border: none;
        padding: 0.5rem 1.5rem;
        transition: background-color 0.3s ease;
    }}
    div.stDownloadButton > button:first-child:hover {{
        background-color: #0052a3;
        color: white;
    }}
    [data-testid="stMetricValue"] {{
        font-size: 1.8rem;
        font-weight: 700;
    }}
    [data-testid="metric-container"] {{
        background-color: rgba(250, 250, 250, 0.1);
        border: 1px solid rgba(128, 128, 128, 0.2);
        padding: 15px;
        border-radius: 10px;
    }}
</style>
""", unsafe_allow_html=True)

st.sidebar.title("🏏 NCU Cricket Hub")
app_mode = st.sidebar.selectbox("Select Tool", [
    "Bulk Averages Calculator", 
    "Player Word Doc Generator", 
    "Registration Checks",
    "Midweek Registration & Starring Check",
    "Starring & Inactivity Reports",
    "Club Fines Generator",
    "Unregistered Player Fines Generator"
])

# ==========================================
# TOOL 1: BULK AVERAGES
# ==========================================
if app_mode == "Bulk Averages Calculator":
    st.title(PAGE_TITLES["bulk_averages"])
    st.markdown("Generate full season averages for Men's, Women's, or Midweek leagues.")

    st.subheader("Select League Domain")
    domain = st.radio("Choose the ruleset and default files to apply:", ["Men's", "Women's", "Midweek"], horizontal=True)

    st.divider() 
    st.subheader("Set Minimum Thresholds")
    
    if domain == "Midweek":
        st.markdown("**Midweek Overall Qualifiers** *(Groups are fixed to 20 Runs / 2 Wickets)*")
        col1, col2, col3 = st.columns(3)
        with col1: mw_min_runs = st.number_input("Minimum Runs", min_value=0, value=50)
        with col2: mw_min_innings = st.number_input("Minimum Innings", min_value=0, value=0)
        with col3: mw_min_wickets = st.number_input("Minimum Wickets", min_value=0, value=5)
        
    elif domain == "Women's":
        st.info("💡 **Women's Tiered Rules Active:** Below are your automated target thresholds. Customize any values before exporting.")
        with st.container(border=True):
            st.markdown("🏆 **Premier League and Senior League Section 1**")
            c1, c2, c3, c4 = st.columns(4)
            with c1: w1_runs = st.number_input("Min Runs", min_value=0, value=100, key="w1_runs")
            with c2: w1_bmat = st.number_input("Min Bat Matches", min_value=0, value=5, key="w1_bmat")
            with c3: w1_wick = st.number_input("Min Wickets", min_value=0, value=10, key="w1_wick")
            with c4: w1_mmat = st.number_input("Min Bowl Matches", min_value=0, value=5, key="w1_mmat")
            
        with st.container(border=True):
            st.markdown("🏏 **Junior League Sections 1**")
            c1, c2, c3, c4 = st.columns(4)
            with c1: w2_runs = st.number_input("Min Runs", min_value=0, value=25, key="w2_runs")
            with c2: w2_bmat = st.number_input("Min Bat Matches", min_value=0, value=2, key="w2_bmat")
            with c3: w2_wick = st.number_input("Min Wickets", min_value=0, value=2, key="w2_wick")
            with c4: w2_mmat = st.number_input("Min Bowl Matches", min_value=0, value=2, key="w2_mmat")
            
    else:  # Men's
        st.info("💡 **Men's Tiered Rules Active:** Below are your automated target thresholds. Customize any values before exporting.")
        with st.container(border=True):
            st.markdown("🏆 **Premier League and Section 1**")
            c1, c2, c3, c4 = st.columns(4)
            with c1: t1_runs = st.number_input("Min Runs", min_value=0, value=200, key="t1_runs")
            with c2: t1_bmat = st.number_input("Min Bat Matches", min_value=0, value=5, key="t1_bmat")
            with c3: t1_wick = st.number_input("Min Wickets", min_value=0, value=15, key="t1_wick")
            with c4: t1_mmat = st.number_input("Min Bowl Matches", min_value=0, value=5, key="t1_mmat")
            
        with st.container(border=True):
            st.markdown("🛡️ **Senior League Sections 2 and 3**")
            c1, c2, c3, c4 = st.columns(4)
            with c1: t2_runs = st.number_input("Min Runs", min_value=0, value=150, key="t2_runs")
            with c2: t2_bmat = st.number_input("Min Bat Matches", min_value=0, value=5, key="t2_bmat")
            with c3: t2_wick = st.number_input("Min Wickets", min_value=0, value=10, key="t2_wick")
            with c4: t2_mmat = st.number_input("Min Bowl Matches", min_value=0, value=5, key="t2_mmat")
            
        with st.container(border=True):
            st.markdown("🏏 **Junior League Sections 1 to 10**")
            c1, c2, c3, c4 = st.columns(4)
            with c1: t3_runs = st.number_input("Min Runs", min_value=0, value=100, key="t3_runs")
            with c2: t3_bmat = st.number_input("Min Bat Matches", min_value=0, value=3, key="t3_bmat")
            with c3: t3_wick = st.number_input("Min Wickets", min_value=0, value=5, key="t3_wick")
            with c4: t3_mmat = st.number_input("Min Bowl Matches", min_value=0, value=3, key="t3_mmat")
            
        with st.container(border=True):
            st.markdown("🌱 **Junior League Sections 11a to 11b**")
            c1, c2, c3, c4 = st.columns(4)
            with c1: t4_runs = st.number_input("Min Runs", min_value=0, value=50, key="t4_runs")
            with c2: t4_bmat = st.number_input("Min Bat Matches", min_value=0, value=3, key="t4_bmat")
            with c3: t4_wick = st.number_input("Min Wickets", min_value=0, value=3, key="t4_wick")
            with c4: t4_mmat = st.number_input("Min Bowl Matches", min_value=0, value=3, key="t4_mmat")

    st.divider() 
    st.subheader("Set Sorting Preferences")
    colX, colY = st.columns(2)
    with colX: bat_sort_pref = st.selectbox("Batting Primary Sort", ["Runs", "Average", "Strike Rate"], index=0)
    with colY: bowl_sort_pref = st.selectbox("Bowling Primary Sort", ["Wickets", "Average", "Economy", "Strike Rate"], index=0)

    # --- ADDED CODE: Match Inclusion & Threshold Settings ---
    st.divider()
    st.subheader("Match Inclusion & Threshold Settings")
    colA, colB, colC = st.columns(3)
    
    include_cup = True
    include_t20 = True
    disable_thresholds = False
    
    with colA:
        if domain in ["Men's", "Women's"]:
            include_cup = st.checkbox("Include Cup Matches", value=True)
    with colB:
        if domain == "Men's":
            include_t20 = st.checkbox("Include T20 Matches", value=True)
    with colC:
        disable_thresholds = st.checkbox("Set all target thresholds to 0", value=False)
    # --------------------------------------------------------

    with st.sidebar:
        st.divider() 
        c_files = eng.DEFAULT_FILES[domain]
        with st.expander("📁 File Path Configurations", expanded=False):
            st.markdown("*Verify or update the default local file paths below.*")
            f_reg = st.text_input("Official Registry (Excel)", value=c_files["reg"], key=f"avg_reg_{domain}")
            f_alias = st.text_input("Aliases Master (Excel)", value=c_files["alias"], key=f"avg_alias_{domain}")
            f_league = st.text_input("League Structure (Excel)", value=c_files["league"], key=f"avg_league_{domain}")
            f_bat = st.text_input("Batting Stats (Excel)", value=c_files["bat"], key=f"avg_bat_{domain}")
            f_bowl = st.text_input("Bowling Stats (Excel)", value=c_files["bowl"], key=f"avg_bowl_{domain}")
            f_cup = st.text_input("Cup Master (Excel)", value="NCU_Cup_Fixtures.xlsx", key=f"avg_cup_{domain}")
    
    include_irish = False
    if domain == "Men's":
        include_irish = st.checkbox("Include Irish Competitions in Averages?", value=False)
        
    if domain == "Men's" and include_irish:
        with st.sidebar:
            with st.expander("📁 File Path Configurations", expanded=False):
                f_irish_bat = st.text_input("Irish Batting Stats (Excel)", value="Irish Competitions 2026 Batting stats.xlsx", key="avg_irish_bat")
                f_irish_bowl = st.text_input("Irish Bowling Stats (Excel)", value="Irish Competitions 2026 Bowling stats.xlsx", key="avg_irish_bowl")

    st.subheader("Generate Averages")
    if st.button("🚀 Process Averages", type="primary"):
        files_to_check = [f_reg, f_alias, f_league, f_bat, f_bowl, f_cup]
        if domain == "Men's" and include_irish:
            files_to_check.extend([f_irish_bat, f_irish_bowl])
            
        missing_files = [f for f in files_to_check if not os.path.exists(f)]
        if missing_files:
            st.error(f"Cannot find the following files in this directory:\n\n" + "\n".join([f"- {f}" for f in missing_files]))
        else:
            with st.spinner(f"Running {domain} Averages Engine..."):
                try:
                    reg_players = pd.read_excel(f_reg)
                    aliases = pd.read_excel(f_alias)
                    league_structure = pd.read_excel(f_league)
                    
                    batting = pd.read_excel(f_bat)
                    bowling = pd.read_excel(f_bowl)
                    
                    if domain == "Men's" and include_irish:
                        if os.path.exists(f_irish_bat):
                            batting = pd.concat([batting, pd.read_excel(f_irish_bat)], ignore_index=True)
                        if os.path.exists(f_irish_bowl):
                            bowling = pd.concat([bowling, pd.read_excel(f_irish_bowl)], ignore_index=True)

                    # --- ADDED CODE: Advanced Cup/T20 Filtering ---
                    cup_match_dict = {}
                    if os.path.exists(f_cup):
                        try:
                            excel_file_cup = pd.ExcelFile(f_cup)
                            target_sheet = excel_file_cup.sheet_names[0]
                            for sheet in excel_file_cup.sheet_names:
                                if domain.lower().replace("'", "") in sheet.lower().replace("'", ""):
                                    target_sheet = sheet
                                    break
                            cup_df = pd.read_excel(f_cup, sheet_name=target_sheet, header=None)
                            
                            def local_parse(group_str):
                                try:
                                    group_str = str(group_str).strip()
                                    parts = group_str.rsplit(' - ', 1)
                                    date_str = parts[1].strip() if len(parts) == 2 else group_str
                                    rest = parts[0].strip() if len(parts) == 2 else group_str
                                    match_date = pd.to_datetime(date_str, dayfirst=True, errors='coerce')
                                    if pd.notna(match_date): match_date = match_date.normalize()
                                    if ' v ' in rest:
                                        team_a, remainder = rest.split(' v ', 1)
                                        team_b = remainder.rsplit(', ', 1)[0] if ', ' in remainder else (remainder.rsplit(' - ', 1)[0] if ' - ' in remainder else remainder)
                                    else:
                                        team_a, team_b = rest, "Unknown"
                                    return team_a.strip(), team_b.strip(), match_date
                                except: return None, None, None

                            for _, row_data in cup_df.iterrows():
                                match_str_raw = str(row_data[0]).strip()
                                cup_name = str(row_data[1]).strip()
                                if match_str_raw.lower() in ['match string', 'match group', 'match', 'nan']: continue
                                cleaned_match_str = eng.doc_format_cricket_names(match_str_raw, domain)
                                c_team_a, c_team_b, c_date = local_parse(cleaned_match_str)
                                if c_team_a and c_team_b:
                                    teams = sorted([str(c_team_a).lower(), str(c_team_b).lower()])
                                    if pd.notna(c_date):
                                        cup_match_dict[f"{teams[0]}_{teams[1]}_{c_date.strftime('%Y-%m-%d')}"] = cup_name
                                    else:
                                        cup_match_dict[f"{teams[0]}_{teams[1]}"] = cup_name
                        except Exception: pass

                    def is_target_match(grp_str, target_kws):
                        grp_str_clean = str(grp_str).lower()
                        # 1. Literal substring match in the Group string
                        if any(kw in grp_str_clean for kw in target_kws):
                            return True
                        # 2. Cross-reference NCU Cup Fixtures registry
                        if cup_match_dict:
                            c_team_a, c_team_b, c_date = local_parse(eng.doc_format_cricket_names(grp_str, domain))
                            if c_team_a and c_team_b:
                                teams = sorted([str(c_team_a).lower(), str(c_team_b).lower()])
                                comp = None
                                if pd.notna(c_date):
                                    comp = cup_match_dict.get(f"{teams[0]}_{teams[1]}_{c_date.strftime('%Y-%m-%d')}")
                                if not comp:
                                    comp = cup_match_dict.get(f"{teams[0]}_{teams[1]}")
                                if comp and any(kw in str(comp).lower() for kw in target_kws):
                                    return True
                        return False

                    cup_kws = ['cup', 'trophy', 'shield', 'plate', 'bowl', 'vase', 'challenge']
                    t20_kws = ['t20', 'twenty20']

                    if domain in ["Men's", "Women's"] and not include_cup:
                        if 'Group' in batting.columns:
                            batting = batting[~batting['Group'].apply(lambda x: is_target_match(x, cup_kws))]
                        if 'Group' in bowling.columns:
                            bowling = bowling[~bowling['Group'].apply(lambda x: is_target_match(x, cup_kws))]
                            
                    if domain == "Men's" and not include_t20:
                        if 'Group' in batting.columns:
                            batting = batting[~batting['Group'].apply(lambda x: is_target_match(x, t20_kws))]
                        if 'Group' in bowling.columns:
                            bowling = bowling[~bowling['Group'].apply(lambda x: is_target_match(x, t20_kws))]
                    # ------------------------------------------------------

                    alias_map = eng.build_alias_map(aliases, domain)
                    league_dict, team_keys, original_league_order = eng.build_league_dict(league_structure)
                    player_club_map = eng.build_player_club_map(reg_players, alias_map, domain)
                    
                    batting['Cleaned Name'] = batting.apply(lambda r: eng.cleanse_name_contextual(r['Name'], r, alias_map), axis=1)
                    bowling['Cleaned Name'] = bowling.apply(lambda r: eng.cleanse_name_contextual(r['Bowler'], r, alias_map), axis=1)
                    
                    batting_avgs, bowling_avgs = eng.calculate_averages(batting, bowling, player_club_map, team_keys, league_dict, domain, bat_sort_pref, bowl_sort_pref)
                    
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
                                bat_thresh, bat_match_thresh = (mw_min_runs, mw_min_innings) if "Overall" in league else (20, 0)
                                bowl_thresh, bowl_match_thresh = (mw_min_wickets, 0) if "Overall" in league else (2, 0)
                            elif domain == "Women's":
                                if "premier" in name_lower or "senior league 1" in name_lower or "senior 1" in name_lower or "senior league" in name_lower:
                                    bat_thresh, bat_match_thresh, bowl_thresh, bowl_match_thresh = w1_runs, w1_bmat, w1_wick, w1_mmat
                                else:
                                    bat_thresh, bat_match_thresh, bowl_thresh, bowl_match_thresh = w2_runs, w2_bmat, w2_wick, w2_mmat
                            else:  # Men's
                                if "premier" in name_lower or "senior league 1" in name_lower or "senior 1" in name_lower:
                                    bat_thresh, bat_match_thresh, bowl_thresh, bowl_match_thresh = t1_runs, t1_bmat, t1_wick, t1_mmat
                                elif "senior league 2" in name_lower or "senior league 3" in name_lower or "senior 2" in name_lower or "senior 3" in name_lower:
                                    bat_thresh, bat_match_thresh, bowl_thresh, bowl_match_thresh = t2_runs, t2_bmat, t2_wick, t2_mmat
                                elif "11a" in name_lower or "11b" in name_lower or "junior league 11" in name_lower:
                                    bat_thresh, bat_match_thresh, bowl_thresh, bowl_match_thresh = t4_runs, t4_bmat, t4_wick, t4_mmat
                                else:
                                    bat_thresh, bat_match_thresh, bowl_thresh, bowl_match_thresh = t3_runs, t3_bmat, t3_wick, t3_mmat

                            # --- ADDED CODE: Zero out thresholds ---
                            if disable_thresholds:
                                bat_thresh, bat_match_thresh, bowl_thresh, bowl_match_thresh = 0, 0, 0, 0
                            # ---------------------------------------

                            if not league_bat.empty:
                                league_bat = league_bat[(league_bat['Runs'] >= bat_thresh) & (league_bat['Matches'] >= bat_match_thresh)]
                                if not league_bat.empty:
                                    league_bat.insert(0, 'Position', range(1, len(league_bat) + 1))
                                    eng.format_excel_sheet(writer, league_bat, f"{tab_prefix} Bat", min_label=f"Min {bat_thresh} runs, {bat_match_thresh} matches")
                                
                            if not league_bowl.empty:
                                league_bowl = league_bowl[(league_bowl['Wickets'] >= bowl_thresh) & (league_bowl['Matches'] >= bowl_match_thresh)]
                                if not league_bowl.empty:
                                    league_bowl.insert(0, 'Position', range(1, len(league_bowl) + 1))
                                    eng.format_excel_sheet(writer, league_bowl, f"{tab_prefix} Bowl", min_label=f"Min {bowl_thresh} wickets, {bowl_match_thresh} matches")
                    
                    st.success("✅ Averages calculated successfully!")
                    st.subheader("👀 Computed Averages Preview (Top 50)")
                    tab_preview_bat, tab_preview_bowl = st.tabs(["🏏 Batting Preview", "🔮 Bowling Preview"])
                    with tab_preview_bat:
                        st.dataframe(batting_avgs.head(50), width="stretch", hide_index=True)
                    with tab_preview_bowl:
                        st.dataframe(bowling_avgs.head(50), width="stretch", hide_index=True)
                    
                    st.divider()
                    file_out_name = f"{domain.replace('''s''', '')}_Season_Averages_All_Leagues_{datetime.now().strftime('%Y-%m-%d_%H%M%S')}.xlsx"
                    st.download_button(
                        label="📥 Download Output Excel File",
                        data=output_buffer.getvalue(),
                        file_name=file_out_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                except Exception as e:
                    st.error(f"An error occurred during processing: {str(e)}")

# ==========================================
# TOOL 2: WORD DOC GENERATOR
# ==========================================
elif app_mode == "Player Word Doc Generator":
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
                st.markdown("*Verify or update the default local file paths below.*")
                f_reg = st.text_input("Official Registry (Excel)", value=c_files["reg"], key=f"doc_reg_{domain}")
                f_alias = st.text_input("Aliases Master (Excel)", value=c_files["alias"], key=f"doc_alias_{domain}")
                f_bat = st.text_input("Batting Stats (Excel)", value=c_files["bat"], key=f"doc_bat_{domain}")
                f_bowl = st.text_input("Bowling Stats (Excel)", value=c_files["bowl"], key=f"doc_bowl_{domain}")
        
        include_irish = False
        if domain == "Men's":
            include_irish = st.checkbox("Include Irish Competitions in Player Report?", value=False)
            
        if domain == "Men's" and include_irish:
            with st.sidebar:
                with st.expander("📁 File Path Configurations", expanded=False):
                    f_irish_bat = st.text_input("Irish Batting Stats (Excel)", value="Irish Competitions 2026 Batting stats.xlsx", key="doc_irish_bat")
                    f_irish_bowl = st.text_input("Irish Bowling Stats (Excel)", value="Irish Competitions 2026 Bowling stats.xlsx", key="doc_irish_bowl")

        # Reset search if domain is switched
        if 'doc_last_domain' not in st.session_state or st.session_state.doc_last_domain != domain:
            st.session_state.player_search_active = False
            st.session_state.doc_last_domain = domain

        with st.container(border=True):
            st.subheader("🔍 Search Player Database")
            search_query = st.text_input("Enter the player's full name or scorecard alias:", placeholder="e.g., Pawan Thakur")
            
            col_btn, _ = st.columns([1, 4])
            with col_btn:
                execute_search = st.button("🔍 Search Player", type="primary", width="stretch")

        # --- Session State Management ---
        if 'player_search_active' not in st.session_state:
            st.session_state.player_search_active = False

        if execute_search:
            st.session_state.player_search_active = True
            st.session_state.player_search_query = search_query
            st.session_state.data_loaded = False 

        if st.session_state.player_search_active:
            current_query = st.session_state.player_search_query
            
            files_to_check = [f_reg, f_alias, f_bat, f_bowl]
            if domain == "Men's" and include_irish:
                files_to_check.extend([f_irish_bat, f_irish_bowl])
                
            missing_files = [f for f in files_to_check if not os.path.exists(f)]
            
            if missing_files:
                st.error(f"Cannot find the following files:\n\n" + "\n".join([f"- {f}" for f in missing_files]))
            elif not current_query:
                st.warning("Please enter a player name to search.")
            else:
                if not st.session_state.get('data_loaded'):
                    with st.spinner("Searching datasets and building options..."):
                        reg_players = pd.read_excel(f_reg)
                        aliases = pd.read_excel(f_alias)
                        
                        batting = pd.read_excel(f_bat)
                        bowling = pd.read_excel(f_bowl)
                        
                        if domain == "Men's" and include_irish:
                            if os.path.exists(f_irish_bat):
                                batting = pd.concat([batting, pd.read_excel(f_irish_bat)], ignore_index=True)
                            if os.path.exists(f_irish_bowl):
                                bowling = pd.concat([bowling, pd.read_excel(f_irish_bowl)], ignore_index=True)

                        alias_map = eng.build_alias_map(aliases, domain)
                        player_club_map = eng.build_player_club_map(reg_players, alias_map, domain)
                        
                        def resolve_duplicates(row, name_col):
                            name = str(row[name_col])
                            row_team = str(row.get('Team', '')).lower()
                            match_grp = str(row.get('Group', '')).lower()
                            
                            if domain == "Men's" and name in eng.KNOWN_DUPLICATES:
                                for club in eng.KNOWN_DUPLICATES[name]:
                                    if club.lower() in row_team:
                                        return f"{name} ({club})"
                                for club in eng.KNOWN_DUPLICATES[name]:
                                    if club.lower() in match_grp:
                                        return f"{name} ({club})"
                            return name
                        
                        batting['Name'] = batting.apply(lambda r: eng.cleanse_name_contextual(r['Name'], r, alias_map), axis=1)
                        bowling['Bowler'] = bowling.apply(lambda r: eng.cleanse_name_contextual(r['Bowler'], r, alias_map), axis=1)
                        
                        batting['Name'] = batting.apply(lambda x: resolve_duplicates(x, 'Name'), axis=1)
                        bowling['Bowler'] = bowling.apply(lambda x: resolve_duplicates(x, 'Bowler'), axis=1)
                        
                        batting['Group'] = batting['Group'].apply(lambda x: eng.doc_format_cricket_names(x, domain))
                        bowling['Group'] = bowling['Group'].apply(lambda x: eng.doc_format_cricket_names(x, domain))

                        clean_q = current_query.strip().lower()

                        # 1. Search across Official Registry & Alias Master to find official target names
                        target_official_names = set()

                        # Check exact / substring matches in Aliases Master
                        if 'Input Name (Scorecard/Stats)' in aliases.columns and 'Official Registered Name' in aliases.columns:
                            alias_matches = aliases[
                                aliases['Input Name (Scorecard/Stats)'].astype(str).str.contains(clean_q, case=False, na=False) |
                                aliases['Official Registered Name'].astype(str).str.contains(clean_q, case=False, na=False)
                            ]
                            target_official_names.update(alias_matches['Official Registered Name'].dropna().astype(str).str.strip().tolist())
                        else:
                            for _, row in aliases.iterrows():
                                a_val, o_val = str(row.iloc[0]), str(row.iloc[1])
                                if clean_q in a_val.lower() or clean_q in o_val.lower():
                                    target_official_names.add(o_val.strip())

                        # Check exact / substring matches in Official Registry
                        if 'Full Name' in reg_players.columns:
                            reg_matches = reg_players[reg_players['Full Name'].astype(str).str.contains(clean_q, case=False, na=False)]
                            target_official_names.update(reg_matches['Full Name'].dropna().astype(str).str.strip().tolist())

                        # Check stats directly
                        bat_direct = batting[batting['Name'].astype(str).str.contains(clean_q, case=False, na=False)]['Name'].unique().tolist()
                        bowl_direct = bowling[bowling['Bowler'].astype(str).str.contains(clean_q, case=False, na=False)]['Bowler'].unique().tolist()
                        target_official_names.update(bat_direct + bowl_direct)

                        # 2. Filter stats by the collected target official names
                        matched_batting_list = []
                        matched_bowling_list = []
                        
                        for off_name in target_official_names:
                            matched_batting_list.append(batting[batting['Name'].astype(str).str.contains(off_name, case=False, na=False)])
                            matched_bowling_list.append(bowling[bowling['Bowler'].astype(str).str.contains(off_name, case=False, na=False)])

                        # Concat without ignoring index to preserve original row IDs
                        matched_batting = pd.concat(matched_batting_list) if matched_batting_list else pd.DataFrame()
                        matched_bowling = pd.concat(matched_bowling_list) if matched_bowling_list else pd.DataFrame()

                        # Deduplicate rows caught by multiple overlapping partial string matches
                        if not matched_batting.empty:
                            matched_batting = matched_batting[~matched_batting.index.duplicated(keep='first')].reset_index(drop=True)
                            
                        if not matched_bowling.empty:
                            matched_bowling = matched_bowling[~matched_bowling.index.duplicated(keep='first')].reset_index(drop=True)

                        found_batters = matched_batting['Name'].dropna().unique().tolist() if not matched_batting.empty else []
                        found_bowlers = matched_bowling['Bowler'].dropna().unique().tolist() if not matched_bowling.empty else []
                        
                        raw_unique_players = list(set(found_batters + found_bowlers))
                        
                        def player_sort_key(name):
                            pure_name = name.split(' (')[0].strip()
                            club = player_club_map.get(name.lower(), "Unknown Club").lower()
                            parts = pure_name.split()
                            if len(parts) > 1:
                                surname = parts[-1].lower()
                                firstnames = " ".join(parts[:-1]).lower()
                            elif len(parts) == 1:
                                surname = parts[0].lower()
                                firstnames = ""
                            else:
                                surname, firstnames = "", ""
                            return (surname, firstnames, club)
                            
                        st.session_state.matched_batting = matched_batting
                        st.session_state.matched_bowling = matched_bowling
                        st.session_state.unique_players = sorted(raw_unique_players, key=player_sort_key)
                        st.session_state.reg_players = reg_players
                        st.session_state.aliases_df = aliases
                        st.session_state.player_club_map = player_club_map
                        st.session_state.data_loaded = True

                matched_batting = st.session_state.matched_batting
                matched_bowling = st.session_state.matched_bowling
                unique_players = st.session_state.unique_players
                reg_players = st.session_state.reg_players
                aliases_df = st.session_state.aliases_df

                if matched_batting.empty and matched_bowling.empty:
                    st.error(f"No statistics found for '{current_query}'. Please try another name.")
                else:
                    def get_club_for_player(name):
                        if '(' in name and ')' in name:
                            return name.split('(')[-1].replace(')', '').strip()
                        club = st.session_state.player_club_map.get(name.lower(), None)
                        if club and str(club).lower() not in ['nan', 'none', '', 'unknown club']:
                            return str(club).replace(" Cricket Club", "").replace(" CC", "").strip()
                        return "Unknown Club"

                    def format_player_display(name):
                        pure_registered_name = name.split(' (')[0].strip()
                        club_clean = get_club_for_player(name)
                        
                        p_aliases = eng.get_player_aliases(pure_registered_name, aliases_df)
                        
                        if p_aliases:
                            alias_str = " / ".join(p_aliases)
                            return f"{pure_registered_name} / {alias_str} ({club_clean})"
                        return f"{pure_registered_name} ({club_clean})"

                    if len(unique_players) == 1:
                        active_player = unique_players[0]
                        pure_registered_name = active_player.split(' (')[0].strip()
                        p_aliases = eng.get_player_aliases(pure_registered_name, aliases_df)
                        
                        display_lbl = format_player_display(active_player)
                        st.success(f"Found Match: {display_lbl}")
                        
                        p_bat = matched_batting[matched_batting['Name'] == active_player]
                        p_bowl = matched_bowling[matched_bowling['Bowler'] == active_player]
                        
                        doc_io, filename = eng.generate_single_player_doc(
                            active_player, p_bat, p_bowl, reg_players, domain, aliases_list=p_aliases
                        )
                        
                        st.download_button(
                            label="📥 Download Player Word Document",
                            data=doc_io.getvalue(),
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary"
                        )
                    else:
                        st.warning(f"Multiple players match '{current_query}'. Please select the players you want to generate reports for.")
                        
                        select_all = st.checkbox("Select all players")
                        
                        selected_players = st.multiselect(
                            "Select players:",
                            options=unique_players,
                            default=unique_players if select_all else [],
                            format_func=format_player_display
                        )
                        
                        if selected_players:
                            if len(selected_players) == 1:
                                active_player = selected_players[0]
                                pure_registered_name = active_player.split(' (')[0].strip()
                                p_aliases = eng.get_player_aliases(pure_registered_name, aliases_df)
                                
                                p_bat = matched_batting[matched_batting['Name'] == active_player]
                                p_bowl = matched_bowling[matched_bowling['Bowler'] == active_player]
                                
                                doc_io, filename = eng.generate_single_player_doc(
                                    active_player, p_bat, p_bowl, reg_players, domain, aliases_list=p_aliases
                                )
                                
                                st.download_button(
                                    label=f"📥 Download Report for {format_player_display(active_player)}",
                                    data=doc_io.getvalue(),
                                    file_name=filename,
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    type="primary",
                                    key="dl_single_multi"
                                )
                            else:
                                zip_buffer = io.BytesIO()
                                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                                    for active_player in selected_players:
                                        pure_registered_name = active_player.split(' (')[0].strip()
                                        p_aliases = eng.get_player_aliases(pure_registered_name, aliases_df)
                                        
                                        p_bat = matched_batting[matched_batting['Name'] == active_player]
                                        p_bowl = matched_bowling[matched_bowling['Bowler'] == active_player]
                                        
                                        doc_io, filename = eng.generate_single_player_doc(
                                            active_player, p_bat, p_bowl, reg_players, domain, aliases_list=p_aliases
                                        )
                                        zip_file.writestr(filename, doc_io.getvalue())
                                        
                                st.download_button(
                                    label=f"📦 Download Reports for {len(selected_players)} Players (ZIP)",
                                    data=zip_buffer.getvalue(),
                                    file_name=f"Player_Reports_{current_query.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.zip",
                                    mime="application/zip",
                                    type="primary",
                                    key="dl_zip_multi"
                                )
                        else:
                            st.info("Please select at least one player to generate a report.")
                            
# ==========================================
# TOOL 3: WEEKEND REGISTRATION CHECKS
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
                st.markdown("*Verify or update the default local file paths below.*")
                f_reg = st.text_input("Official Registry (Excel)", value=c_files["reg"], key=f"reg_check_reg_{domain}")
                f_alias = st.text_input("Aliases Master (Excel)", value=c_files["alias"], key=f"reg_check_alias_{domain}")
                f_starring = st.text_input("Starring Master (Excel)", value=c_files["starring"], key=f"reg_check_starring_{domain}")
                f_cup = st.text_input("Cup Master (Excel)", value="NCU_Cup_Fixtures.xlsx", key=f"reg_check_cup_{domain}")
                f_league = st.text_input("League Structure (Excel)", value=c_files["league"], key=f"reg_check_league_{domain}")
                f_bat = st.text_input("Batting Stats (Excel)", value=c_files["bat"], key=f"reg_check_bat_{domain}")
                f_bowl = st.text_input("Bowling Stats (Excel)", value=c_files["bowl"], key=f"reg_check_bowl_{domain}")
        
        include_irish = False
        if domain == "Men's":
            include_irish = st.checkbox("Include Irish Competitions in Audit?", value=False)
            
        if domain == "Men's" and include_irish:
            with st.sidebar:
                with st.expander("📁 File Path Configurations", expanded=False):
                    f_irish_bat = st.text_input("Irish Batting Stats (Excel)", value="Irish Competitions 2026 Batting stats.xlsx", key="reg_irish_bat")
                    f_irish_bowl = st.text_input("Irish Bowling Stats (Excel)", value="Irish Competitions 2026 Bowling stats.xlsx", key="reg_irish_bowl")

        st.subheader("Run Audit Engine")
        if st.button("🚀 Execute Security Audit", type="primary"):
            files_to_check = [f for f in [f_reg, f_alias, f_league, f_bat, f_bowl, f_cup] if f]
            if f_starring: files_to_check.append(f_starring)
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
                            excel_io, doc_io = eng.run_registration_audit(domain, start_ts, end_ts, f_reg, f_alias, f_starring, f_league, f_bat, f_bowl, f_irish_bat, f_irish_bowl, f_cup)
                        else:
                            excel_io, doc_io = eng.run_registration_audit(domain, start_ts, end_ts, f_reg, f_alias, f_starring, f_league, f_bat, f_bowl, f_cup=f_cup)
                        
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
                        with m_col1:
                            st.metric(label="⚠️ Unregistered Match Appearances", value=unreg_count, delta=f"{unreg_count} Flagged", delta_color="inverse")
                        with m_col2:
                            st.metric(label="ℹ️ Deemed Registered Records", value=deemed_count, delta=f"{deemed_count} Tracked", delta_color="off")
                        with m_col3:
                            st.metric(label="🚨 Starring Violations", value=star_count, delta=f"{star_count} Flagged", delta_color="inverse")

                        if unreg_count > 0 or deemed_count > 0 or star_count > 0:
                            st.subheader("📋 Audit Report Previews")
                            if unreg_count > 0:
                                Guide = st.expander("⚠️ Unregistered Matches")
                                with Guide: st.dataframe(df_unreg, width="stretch", hide_index=True)
                            if deemed_count > 0:
                                with st.expander("ℹ️ Deemed Registered Players"):
                                    st.dataframe(df_deemed, width="stretch", hide_index=True)
                            if star_count > 0:
                                with st.expander("🚨 Starring Violations"):
                                    st.dataframe(df_starring_viols, width="stretch", hide_index=True)

                        st.divider()
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                            date_str = f"{start_ts.strftime('%d-%m-%Y')}_to_{end_ts.strftime('%d-%m-%Y')}"
                            prefix = domain.replace("'", "")
                            
                            zip_file.writestr(f"{prefix}_Audit_Database_{date_str}.xlsx", excel_io.getvalue())
                            zip_file.writestr(f"{prefix}_Audit_Report_{date_str}.docx", doc_io.getvalue())
                                
                        st.download_button(
                            label="📦 Download Audit Results (ZIP)",
                            data=zip_buffer.getvalue(),
                            file_name=f"{domain.replace('''s''', '')}_Registration_Audit_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                            mime="application/zip",
                            type="primary"
                        )
                    except Exception as e:
                        st.error(f"An error occurred during processing: {str(e)}")

# ==========================================
# TOOL 4: MIDWEEK REGISTRATION CHECKS
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
            
        with st.sidebar:
            st.divider() 
            with st.expander("📁 File Path Configurations", expanded=False):
                st.markdown("*Verify or update the default local file paths below.*")
                f_reg = st.text_input("Official Registry (Excel)", value=eng.DEFAULT_FILES["Midweek"]["reg"], key="mw_check_reg")
                f_alias = st.text_input("Aliases Master (Excel)", value=eng.DEFAULT_FILES["Midweek"]["alias"], key="mw_check_alias")
                f_starring = st.text_input("Men's Starring Master (Excel)", value=eng.DEFAULT_FILES["Men's"]["starring"], key="mw_check_starring")
                f_weekend_league = st.text_input("Weekend League Structure (Excel)", value=eng.DEFAULT_FILES["Men's"]["league"], key="mw_check_wknd_league")
                f_midweek_league = st.text_input("Midweek League Structure (Excel)", value=eng.DEFAULT_FILES["Midweek"]["league"], key="mw_check_mw_league")
                f_bat = st.text_input("Midweek Batting Stats (Excel)", value=eng.DEFAULT_FILES["Midweek"]["bat"], key="mw_check_bat")
                f_bowl = st.text_input("Midweek Bowling Stats (Excel)", value=eng.DEFAULT_FILES["Midweek"]["bowl"], key="mw_check_bowl")

        st.subheader("Run Midweek Audit Engine")
        if st.button("🚀 Execute Midweek Audit", type="primary"):
            files_to_check = [f_reg, f_alias, f_starring, f_weekend_league, f_midweek_league, f_bat, f_bowl]
                
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
                        
                        excel_io, doc_io = eng.run_midweek_registration_audit(start_ts, end_ts, f_reg, f_alias, f_starring, f_weekend_league, f_midweek_league, f_bat, f_bowl)
                        
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
                        with m_col1:
                            st.metric(label="⚠️ Unregistered Midweek Players", value=unreg_count, delta=f"{unreg_count} Flagged", delta_color="inverse")
                        with m_col2:
                            st.metric(label="ℹ️ Deemed Registered Players", value=deemed_count, delta=f"{deemed_count} Tracked", delta_color="off")
                        with m_col3:
                            st.metric(label="🚨 Midweek Starring Ceiling Violations", value=star_count, delta=f"{star_count} Flagged", delta_color="inverse")

                        if unreg_count > 0 or deemed_count > 0 or star_count > 0:
                            st.subheader("📋 Audit Report Previews")
                            if unreg_count > 0:
                                with st.expander("⚠️ Unregistered Midweek Matches"):
                                    st.dataframe(df_unreg, width="stretch", hide_index=True)
                            if deemed_count > 0:
                                with st.expander("ℹ️ Deemed Registered Players"):
                                    st.dataframe(df_deemed, width="stretch", hide_index=True)
                            if star_count > 0:
                                with st.expander("🚨 Midweek Ceiling Violations (Junior 3 & Above Starred players)"):
                                    st.dataframe(df_starring_viols, width="stretch", hide_index=True)

                        st.divider()
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                            date_str = f"{start_ts.strftime('%d-%m-%Y')}_to_{end_ts.strftime('%d-%m-%Y')}"
                            
                            zip_file.writestr(f"Midweek_Audit_Database_{date_str}.xlsx", excel_io.getvalue())
                            zip_file.writestr(f"Midweek_Audit_Report_{date_str}.docx", doc_io.getvalue())
                                
                        st.download_button(
                            label="📦 Download Audit Results (ZIP)",
                            data=zip_buffer.getvalue(),
                            file_name=f"Midweek_Registration_Audit_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                            mime="application/zip",
                            type="primary"
                        )
                    except Exception as e:
                        st.error(f"An error occurred during processing: {str(e)}")

# ==========================================
# TOOL 5: STARRING & INACTIVITY REPORTS
# ==========================================
elif app_mode == "Starring & Inactivity Reports":
    st.title(PAGE_TITLES["starring_reports"])
    st.markdown("Generate club-by-club Excel files highlighting inactive starred players (Red/Yellow) and tracking international exemptions (Green).")
    
    st.subheader("Select League Domain")
    domain = st.radio("Choose the dataset domain to audit:", ["Men's", "Women's"], horizontal=True)

    include_irish = False
    if domain == "Men's":
        include_irish = st.checkbox("Include Irish Competitions in Inactivity Reports?", value=False, key="star_include_irish")

    with st.sidebar:
        st.divider() 
        c_files = eng.DEFAULT_FILES[domain]
        with st.expander("📁 File Path Configurations", expanded=False):
            st.markdown("*Verify or update the default local file paths below.*")
            f_reg = st.text_input("Official Registry (Excel)", value=c_files["reg"], key=f"star_reg_{domain}")
            f_alias = st.text_input("Aliases Master (Excel)", value=c_files["alias"], key=f"star_alias_{domain}")
            f_starring = st.text_input("Starring Master (Excel)", value=c_files["starring"], key=f"star_starring_{domain}")
            f_bat = st.text_input("Batting Stats (Excel)", value=c_files["bat"], key=f"star_bat_{domain}")
            f_bowl = st.text_input("Bowling Stats (Excel)", value=c_files["bowl"], key=f"star_bowl_{domain}")
            
            if domain == "Men's" and include_irish:
                f_irish_bat = st.text_input("Irish Batting Stats (Excel)", value="Irish Competitions 2026 Batting stats.xlsx", key="star_irish_bat")
                f_irish_bowl = st.text_input("Irish Bowling Stats (Excel)", value="Irish Competitions 2026 Bowling stats.xlsx", key="star_irish_bowl")

    st.subheader("Generate Reports")
    if st.button("📦 Process All Clubs & Download ZIP", type="primary"):
        files_to_check = [f_reg, f_alias, f_starring, f_bat, f_bowl]
        if domain == "Men's" and include_irish:
            files_to_check.extend([f_irish_bat, f_irish_bowl])
            
        missing_files = [f for f in files_to_check if not os.path.exists(f)]
        
        if missing_files:
            st.error(f"Cannot find the following files:\n\n" + "\n".join([f"- {f}" for f in missing_files]))
        else:
            with st.spinner(f"Generating {domain} Starring Reports for all clubs..."):
                try:
                    if domain == "Men's" and include_irish:
                        zip_buffer = eng.generate_starring_inactivity_reports(
                            domain, f_reg, f_alias, f_starring, f_bat, f_bowl, f_irish_bat, f_irish_bowl
                        )
                    else:
                        zip_buffer = eng.generate_starring_inactivity_reports(
                            domain, f_reg, f_alias, f_starring, f_bat, f_bowl
                        )
                    
                    try:
                        zip_buffer.seek(0)
                        with zipfile.ZipFile(zip_buffer, 'r') as z_file:
                            name_list = z_file.namelist()
                            workbooks_count = sum(1 for item in name_list if item.startswith("NCU_Master_Audit_"))
                            has_unreg = "Unregistered_Starred_Players.xlsx" in name_list
                    except:
                        workbooks_count, has_unreg = 0, False

                    st.success("✅ Reports generated successfully!")
                    st.subheader("📊 Exporter Output Summary")
                    col_star1, col_star2 = st.columns(2)
                    with col_star1:
                        st.metric(label="Clubs Workbooks Created", value=workbooks_count)
                    with col_star2:
                        st.metric(label="Flagged Unregistered Starred Players List", value="Yes" if has_unreg else "No")
                    
                    st.divider()
                    st.download_button(
                        label="📥 Download Club Reports (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name=f"{domain.replace('''s''', '')}_Starring_Reports_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                        mime="application/zip",
                        type="primary"
                    )
                except Exception as e:
                    st.error(f"An error occurred during processing: {str(e)}")

# ==========================================
# TOOL 6: CLUB FINES GENERATOR
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
            
        with st.sidebar:
            st.divider() 
            c_files = eng.DEFAULT_FILES.get(domain, eng.DEFAULT_FILES["Men's"])
            with st.expander("📁 File Path Configurations", expanded=False):
                st.markdown("*Verify or update the default local file paths below.*")
                f_reg = st.text_input("Official Registry (Excel)", value=c_files["reg"], key=f"fines_reg_{domain}")
                f_alias = st.text_input("Aliases Master (Excel)", value=c_files["alias"], key=f"fines_alias_{domain}")
                f_bat = st.text_input("Batting Stats (Excel)", value=c_files["bat"], key=f"fines_bat_{domain}")
                f_bowl = st.text_input("Bowling Stats (Excel)", value=c_files["bowl"], key=f"fines_bowl_{domain}")
                
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
            include_irish = st.checkbox("Include Irish Competitions in Audit?", value=False, key="fines_irish_check")
            if include_irish:
                with st.sidebar:
                    with st.expander("📁 Irish File Path Configurations", expanded=False):
                        f_irish_bat = st.text_input("Irish Batting Stats (Excel)", value="Irish Competitions 2026 Batting stats.xlsx", key="fines_irish_bat")
                        f_irish_bowl = st.text_input("Irish Bowling Stats (Excel)", value="Irish Competitions 2026 Bowling stats.xlsx", key="fines_irish_bowl")
        
        st.divider()
        st.subheader("Forfeited Matches Data")
        st.markdown("Contains the teams fined for forfeiting matches.")
        
        col1, col2 = st.columns([1, 2])
        with col1:
            default_forfeit_path = "Team Fines for forfeiting matches 2026.xlsx"
            use_default_forfeit = st.checkbox(f"Use local '{default_forfeit_path}'", value=os.path.exists(default_forfeit_path))
        with col2:
            f_forfeit = st.file_uploader("Or Upload Forfeits Excel File", type=["xlsx"], key="fines_forfeit_upload")

        st.divider()
        if st.button("📄 Run Engine & Generate Fines Report", type="primary"):
            forfeit_path = f_forfeit if f_forfeit is not None else (default_forfeit_path if use_default_forfeit and os.path.exists(default_forfeit_path) else None)
            
            files_to_check = [f_reg, f_alias, f_bat, f_bowl]
            if domain != "Midweek":
                files_to_check.extend([f_starring, f_league])
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
                                audit_excel_io, _ = eng.run_registration_audit(domain, start_ts, end_ts, f_reg, f_alias, f_starring, f_league, f_bat, f_bowl, f_irish_bat, f_irish_bowl, f_cup)
                            else:
                                audit_excel_io, _ = eng.run_registration_audit(domain, start_ts, end_ts, f_reg, f_alias, f_starring, f_league, f_bat, f_bowl, f_cup=f_cup)
                        else:
                            audit_excel_io, _ = eng.run_midweek_registration_audit(start_ts, end_ts, f_reg, f_alias, f_starring, f_weekend_league, f_midweek_league, f_bat, f_bowl)
                            
                        audit_excel_io.seek(0)
                        
                        doc_io = eng.generate_club_fines_report(audit_excel_io, forfeit_path, start_ts, end_ts)
                        
                        st.success("✅ Fines report generated successfully!")
                        st.download_button(
                            label=f"📥 Download {domain} Fines Report (Word)",
                            data=doc_io.getvalue(),
                            file_name=f"NCU_{domain.replace('''s''', '')}_Fines_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary"
                        )  
                    except Exception as e:
                        st.error(f"An error occurred during processing: {str(e)}")

# ==========================================
# TOOL 7: UNREGISTERED FINES GENERATOR 17-07-2026
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
            
        with st.sidebar:
            st.divider() 
            c_files = eng.DEFAULT_FILES.get(domain, eng.DEFAULT_FILES["Men's"])
            with st.expander("📁 File Path Configurations", expanded=False):
                st.markdown("*Verify or update the default local file paths below.*")
                f_reg = st.text_input("Official Registry (Excel)", value=c_files["reg"], key=f"unreg_reg_{domain}")
                f_alias = st.text_input("Aliases Master (Excel)", value=c_files["alias"], key=f"unreg_alias_{domain}")
                f_bat = st.text_input("Batting Stats (Excel)", value=c_files["bat"], key=f"unreg_bat_{domain}")
                f_bowl = st.text_input("Bowling Stats (Excel)", value=c_files["bowl"], key=f"unreg_bowl_{domain}")
                
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
            include_irish = st.checkbox("Include Irish Competitions in Audit?", value=False, key="unreg_irish_check")
            if include_irish:
                with st.sidebar:
                    with st.expander("📁 Irish File Path Configurations", expanded=False):
                        f_irish_bat = st.text_input("Irish Batting Stats (Excel)", value="Irish Competitions 2026 Batting stats.xlsx", key="unreg_irish_bat")
                        f_irish_bowl = st.text_input("Irish Bowling Stats (Excel)", value="Irish Competitions 2026 Bowling stats.xlsx", key="unreg_irish_bowl")
        
        st.divider()
        if st.button("📄 Run Engine & Generate Unregistered Report", type="primary"):
            
            # Require all the same data sets used in the normal Registration Checks tools
            files_to_check = [f_reg, f_alias, f_bat, f_bowl]
            if domain != "Midweek":
                files_to_check.extend([f_starring, f_league])
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
                        
                        # 1. Run the appropriate audit engine to find violations automatically
                        if domain != "Midweek":
                            if domain == "Men's" and include_irish:
                                audit_excel_io, _ = eng.run_registration_audit(domain, start_ts, end_ts, f_reg, f_alias, f_starring, f_league, f_bat, f_bowl, f_irish_bat, f_irish_bowl, f_cup)
                            else:
                                audit_excel_io, _ = eng.run_registration_audit(domain, start_ts, end_ts, f_reg, f_alias, f_starring, f_league, f_bat, f_bowl, f_cup=f_cup)
                        else:
                            audit_excel_io, _ = eng.run_midweek_registration_audit(start_ts, end_ts, f_reg, f_alias, f_starring, f_weekend_league, f_midweek_league, f_bat, f_bowl)
                            
                        # 2. Reset the buffer so pandas can read it
                        audit_excel_io.seek(0)
                        
                        # 3. Pass the generated audit directly into the Unregistered Fines Generator
                        doc_io = eng.generate_unregistered_fines_only(audit_excel_io)
                        
                        st.success("✅ Unregistered Fines report generated successfully!")
                        st.download_button(
                            label=f"📥 Download {domain} Unregistered Fines Report (Word)",
                            data=doc_io.getvalue(),
                            file_name=f"NCU_{domain.replace('''s''', '')}_Unreg_Fines_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary"
                        )
                    except Exception as e:
                        st.error(f"An error occurred during processing: {str(e)}")