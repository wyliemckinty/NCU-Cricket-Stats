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
    "starring_reports": "🚨 Club Starring & Inactivity Exporter"
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
    "Starring & Inactivity Reports"
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
        files_to_check = [f_reg, f_alias, f_league, f_bat, f_bowl]
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

                    alias_map = eng.build_alias_map(aliases, domain)
                    league_dict, team_keys, original_league_order = eng.build_league_dict(league_structure)
                    player_club_map = eng.build_player_club_map(reg_players, alias_map, domain)
                    
                    # Apply row-based contextual name cleansing
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
            search_query = st.text_input("Enter the player's full name or scorecard alias:", placeholder="e.g., Jay Venus")
            
            col_btn, _ = st.columns([1, 4])
            with col_btn:
                execute_search = st.button("🔍 Search Player", type="primary", width="stretch")

        # --- Session State Management ---
        if 'player_search_active' not in st.session_state:
            st.session_state.player_search_active = False

        # When the user clicks the search button, activate the search state and reset the data cache
        if execute_search:
            st.session_state.player_search_active = True
            st.session_state.player_search_query = search_query
            st.session_state.data_loaded = False 

        # If a search is active, render the results and selection UI
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
                # Cache the data processing so interacting with the multiselect is instant
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

                        search_term = alias_map.get(current_query.lower(), current_query)
                        
                        if search_term.lower() in ['callum weir', 'john weir']:
                            matched_batting = batting[batting['Name'].astype(str).str.contains('Callum Weir|John Weir', case=False, na=False)]
                            matched_bowling = bowling[bowling['Bowler'].astype(str).str.contains('Callum Weir|John Weir', case=False, na=False)]
                        else:
                            matched_batting = batting[batting['Name'].astype(str).str.contains(search_term, case=False, na=False)]
                            matched_bowling = bowling[bowling['Bowler'].astype(str).str.contains(search_term, case=False, na=False)]

                        found_batters = matched_batting['Name'].dropna().unique().tolist()
                        found_bowlers = matched_bowling['Bowler'].dropna().unique().tolist()
                        
                        raw_unique_players = list(set(found_batters + found_bowlers))
                        
                        # Custom sort key: Surname, First Name, Club
                        def player_sort_key(name):
                            # Remove any bracketed club names added by resolve_duplicates
                            pure_name = name.split(' (')[0].strip()
                            
                            # Get the club name for the third sorting tier
                            club = player_club_map.get(name.lower(), "Unknown Club").lower()
                            
                            # Split into surname and first name(s)
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
                            
                        # Store processed data in session state
                        st.session_state.matched_batting = matched_batting
                        st.session_state.matched_bowling = matched_bowling
                        st.session_state.unique_players = sorted(raw_unique_players, key=player_sort_key)
                        st.session_state.reg_players = reg_players
                        st.session_state.player_club_map = player_club_map
                        st.session_state.data_loaded = True
 
                # Retrieve cached variables
                matched_batting = st.session_state.matched_batting
                matched_bowling = st.session_state.matched_bowling
                unique_players = st.session_state.unique_players
                reg_players = st.session_state.reg_players

                if matched_batting.empty and matched_bowling.empty:
                    st.error(f"No statistics found for '{current_query}'. Please try another name.")
                else:
                    if len(unique_players) == 1:
                        st.success(f"Found Match: {unique_players[0]}")
                        active_player = unique_players[0]
                        p_bat = matched_batting[matched_batting['Name'] == active_player]
                        p_bowl = matched_bowling[matched_bowling['Bowler'] == active_player]
                        
                        doc_io, filename = eng.generate_single_player_doc(active_player, p_bat, p_bowl, reg_players, domain)
                        
                        st.download_button(
                            label="📥 Download Player Word Document",
                            data=doc_io.getvalue(),
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary"
                        )
                    else:
                        st.warning(f"Multiple players match '{current_query}'. Please select the players you want to generate reports for.")
                        
                        # Checkbox to trigger 'Select All'
                        select_all = st.checkbox("Select all players")
                        
                        # Format function to append club names dynamically based on actual match data
                        def format_player_display(name):
                            # If resolve_duplicates already appended a club, return as is
                            if '(' in name and ')' in name:
                                return name
                            
                            import re
                            # Fetch the actual scorecard rows for this specific player
                            p_bat = st.session_state.matched_batting[st.session_state.matched_batting['Name'] == name]
                            p_bowl = st.session_state.matched_bowling[st.session_state.matched_bowling['Bowler'] == name]
                            
                            # 1. Try to get the team directly from the scorecard's "Team" column
                            teams_found = []
                            if 'Team' in p_bat.columns:
                                teams_found.extend(p_bat['Team'].dropna().tolist())
                            if 'Team' in p_bowl.columns:
                                teams_found.extend(p_bowl['Team'].dropna().tolist())
                                
                            if teams_found:
                                # Find the team they played for most often in the search results
                                most_common_team = max(set(teams_found), key=teams_found.count)
                                club_clean = str(most_common_team).replace(" Cricket Club", "").replace(" CC", "").strip()
                                # Strip away "1st XI", "2nd", "3", etc. to leave just the club name
                                club_clean = re.sub(r'\s+\d(st|nd|rd|th)?\s*XI?$', '', club_clean, flags=re.IGNORECASE).strip()
                                club_clean = re.sub(r'\s+\d$', '', club_clean).strip()
                                return f"{name} ({club_clean})"

                            # 2. Fallback: calculate frequency from the match 'Group' strings
                            all_groups = pd.concat([p_bat['Group'], p_bowl['Group']]).dropna().tolist()
                            if all_groups:
                                team_frequency = {}
                                for grp in all_groups:
                                    if ' v ' in grp:
                                        t1, t2 = grp.split(' v ')[0].strip(), grp.split(' v ')[1].split(',')[0].strip()
                                        team_frequency[t1] = team_frequency.get(t1, 0) + 1
                                        team_frequency[t2] = team_frequency.get(t2, 0) + 1
                                
                                if team_frequency:
                                    top_teams = sorted(team_frequency.items(), key=lambda x: x[1], reverse=True)
                                    inferred_club = str(top_teams[0][0]).replace(" Cricket Club", "").replace(" CC", "").strip()
                                    inferred_club = re.sub(r'\s+\d(st|nd|rd|th)?\s*XI?$', '', inferred_club, flags=re.IGNORECASE).strip()
                                    inferred_club = re.sub(r'\s+\d$', '', inferred_club).strip()
                                    return f"{name} ({inferred_club})"
                            
                            # 3. Final Fallback: The master registry map
                            club = st.session_state.player_club_map.get(name.lower(), "Unknown Club")
                            club_clean = str(club).replace(" Cricket Club", "").replace(" CC", "").strip()
                            return f"{name} ({club_clean})"
                      
                        # UI for selecting players
                        selected_players = st.multiselect(
                            "Select players:",
                            options=unique_players,
                            default=unique_players if select_all else [],
                            format_func=format_player_display
                        )
                        
                        if selected_players:
                            if len(selected_players) == 1:
                                active_player = selected_players[0]
                                p_bat = matched_batting[matched_batting['Name'] == active_player]
                                p_bowl = matched_bowling[matched_bowling['Bowler'] == active_player]
                                
                                doc_io, filename = eng.generate_single_player_doc(active_player, p_bat, p_bowl, reg_players, domain)
                                
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
                                        p_bat = matched_batting[matched_batting['Name'] == active_player]
                                        p_bowl = matched_bowling[matched_bowling['Bowler'] == active_player]
                                        
                                        doc_io, filename = eng.generate_single_player_doc(active_player, p_bat, p_bowl, reg_players, domain)
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

    # NEW: Add the checkbox toggle for Irish competitions (Men's only)
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
            
            # NEW: Show the Irish file paths in the sidebar if checkbox is selected
            if domain == "Men's" and include_irish:
                f_irish_bat = st.text_input("Irish Batting Stats (Excel)", value="Irish Competitions 2026 Batting stats.xlsx", key="star_irish_bat")
                f_irish_bowl = st.text_input("Irish Bowling Stats (Excel)", value="Irish Competitions 2026 Bowling stats.xlsx", key="star_irish_bowl")

    st.subheader("Generate Reports")
    if st.button("📦 Process All Clubs & Download ZIP", type="primary"):
        # Compile list of files to check based on your checkbox choice
        files_to_check = [f_reg, f_alias, f_starring, f_bat, f_bowl]
        if domain == "Men's" and include_irish:
            files_to_check.extend([f_irish_bat, f_irish_bowl])
            
        missing_files = [f for f in files_to_check if not os.path.exists(f)]
        
        if missing_files:
            st.error(f"Cannot find the following files:\n\n" + "\n".join([f"- {f}" for f in missing_files]))
        else:
            with st.spinner(f"Generating {domain} Starring Reports for all clubs..."):
                try:
                    # NEW: Pass Irish files to the engine if the checkbox is checked
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