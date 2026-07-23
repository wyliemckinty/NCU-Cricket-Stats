# ==========================================
# engine.py
# ==========================================
import pandas as pd
import numpy as np
import os
import re
import io
import zipfile
from datetime import datetime, timedelta
from thefuzz import process, fuzz
import warnings

warnings.filterwarnings('ignore')

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
except ImportError:
    pass

try:
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    pass

# ==========================================
# DEFAULT FILE NAME MAPPING & REGISTRIES
# ==========================================
DEFAULT_FILES = {
    "Men's": {
        "reg": "1. NCU_Registered_Players.xlsx",
        "alias": "2. NCU_Validated_Aliases_Master.xlsx",
        "starring": "3. NCU Complete -Men's- Starring List from 1st June.xlsx",
        "league": "2026 Season League Structure for Gemini AI.xlsx",
        "bat": "NV Play NCU League and Saturday Cup batting stats for season.xlsx",
        "bowl": "NV Play NCU League and Saturday Cup bowling stats for season.xlsx"
    },
    "Women's": {
        "reg": "1. NCU_Registered_Players.xlsx",
        "alias": "12. NCU_Validated_Women's Aliases_Master.xlsx",
        "starring": "13. NCU Complete Women's Starring List from 1st June.xlsx",
        "league": "2026 Season League Structure Women for Gemini AI.xlsx",
        "bat": "NV Play Women's Fixtures batting stats for season.xlsx",
        "bowl": "NV Play Women's Fixtures bowling stats for season.xlsx"
    },
    "Midweek": {
        "reg": "1. NCU_Registered_Players.xlsx",
        "alias": "2. NCU_Validated_Aliases_Master.xlsx",
        "starring": "", 
        "league": "2026 Season Midweek League Structure for Gemini AI.xlsx",
        "bat": "NV Play Midweek League batting stats for season.xlsx",
        "bowl": "NV Play Midweek League bowling stats for season.xlsx"
    }
}

KNOWN_DUPLICATES = {
    'Adam Gardner': ['North Down', 'Carrickfergus'],
    'Adam Mcmaster': ['Templepatrick', 'Ballymena'],
    'Angus Bell': ['Donacloney Mill', 'Lisburn'],
    'David Millar': ['Instonians', 'Ballymena'],
    'Edward Wilson': ['Larne', 'Ballymena'],
    'Grace Wilson': ['Lisburn', 'North Down'],
    'Harry Jackson': ['Cregagh', 'Ards & Donaghadee'],
    'Harsh Shah': ['Laurelvale', 'Cooke Collegians'],
    'Jack Elliott': ['Ballymena', 'Lisburn'],
    'Jack Smyth': ['Templepatrick', 'Saintfield'],
    'Katie Hunter': ['Waringstown', 'Saintfield'],
    'Luke Marshall': ['CSNI', 'Ballymena'],
    'Mirza Baig': ['Dunmurry', 'Dungannon'],
    'Thomas Hamill': ['Waringstown', 'Muckamore'],                        
    'Dylan McCann': ['Lisburn', 'Carrickfergus'],
    'James Shannon': ['Holywood', 'Saintfield'],
    'James Atkinson': ['Holywood 1881', 'Armagh', 'Lisburn'],
    'David Kane': ['Dungannon', 'Templepatrick'],
    'Andrew Holmes': ['CIYMS', 'CSNI'],
    'Timothy Scott': ['Saintfield', 'Ballymena'],
    'Gareth Thompson': ['CSNI', 'Lurgan'],
    'Harry Thompson': ['Derriaghy', 'Lurgan'],
    'Joshua Wilson': ['Armagh', 'Muckamore'],
    'Noah Kelly': ['Cregagh', 'Derriaghy'],
    'Adam Orr': ['Holywood 1881', 'Templepatrick'],
    'Jonathan Bell': ['Derriaghy', 'CSNI'],
    'David Jones': ['Instonians', 'Downpatrick'],
    'Anoop Joseph': ['Lurgan', 'Downpatrick'],
    'Robert Hall': ['Lisburn', 'Laurelvale']
}

# ==========================================
# UNIFIED ENGINE FUNCTIONS 
# ==========================================
def build_alias_map(aliases, domain):
    alias_map = {}
    if 'Input Name (Scorecard/Stats)' in aliases.columns and 'Official Registered Name' in aliases.columns:
        aliases_deduped = aliases.drop_duplicates(subset=['Input Name (Scorecard/Stats)'], keep='last')
        for idx, row in aliases_deduped.iterrows():
            alias_val = str(row['Input Name (Scorecard/Stats)']).replace('‡', '').strip().lower() 
            official_val = str(row['Official Registered Name']).replace('‡', '').strip()
            if alias_val != 'nan':
                alias_map[alias_val] = official_val
    else:
        for idx, row in aliases.iterrows():
            alias_val = str(row.iloc[0]).replace('‡', '').strip().lower() 
            official_val = str(row.iloc[1]).replace('‡', '').strip()
            if alias_val != 'nan':
                alias_map[alias_val] = official_val
                
    return alias_map
    
def get_alias_used_for_player(official_name, search_input, alias_map):
    """
    If the search input was an alias mapped to official_name, 
    return that alias string in original casing or title case.
    """
    if not search_input:
        return None
    
    clean_search = search_input.strip().lower()
    
    # Check if search_input is a key in the alias_map pointing to official_name
    mapped = alias_map.get(clean_search)
    if mapped and mapped.lower() == official_name.lower() and clean_search != official_name.lower():
        return search_input.strip().title()
        
    return None

def cleanse_name(name, alias_map):
    original_name = str(name).replace('‡', '').strip()
    return alias_map.get(original_name.lower(), original_name)

def cleanse_name_contextual(name, row, alias_map):
    original_name = str(name).replace('‡', '').strip()
    original_name_lower = original_name.lower()
    group_lower = str(row.get('Group', '')).lower()
    row_team = str(row.get('Team', '')).lower()
    
    # 1. Callum Weir / John Weir Contextual Override
    if original_name_lower == 'callum weir':
        if 'derriaghy' in group_lower or 'derriaghy' in row_team:
            return 'John Weir'
        elif 'cliftonville' in group_lower or 'cliftonville academy' in group_lower or 'cliftonville' in row_team:
            return 'Callum Weir'
            
    # 2. General known duplicates matched against row metadata
    for dup_name, clubs in KNOWN_DUPLICATES.items():
        if original_name_lower == dup_name.lower():
            # Check the row's scorecard 'Team' column first (solves head-to-head collisions)
            for club in clubs:
                if club.lower() in row_team:
                    return f"{original_name} ({club})"
            # Fallback to matchup context if scorecard team is missing
            for club in clubs:
                if club.lower() in group_lower:
                    return f"{original_name} ({club})"
                    
    return alias_map.get(original_name_lower, original_name)

def build_player_club_map(reg_players, alias_map, domain):
    club_map = {}
    if reg_players is None or reg_players.empty: return club_map
    
    # 1. Forcefully compute the Full Name to guarantee exact matches
    if 'First Name' in reg_players.columns and 'Last Name' in reg_players.columns:
        reg_players['_computed_name'] = reg_players['First Name'].astype(str).str.strip() + ' ' + reg_players['Last Name'].astype(str).str.strip()
    elif 'First Name' in reg_players.columns and 'Surname' in reg_players.columns:
        reg_players['_computed_name'] = reg_players['First Name'].astype(str).str.strip() + ' ' + reg_players['Surname'].astype(str).str.strip()
    elif 'Full Name' in reg_players.columns:
        reg_players['_computed_name'] = reg_players['Full Name'].astype(str).str.replace('‡', '', regex=False).str.strip()
    
    # Gather name columns to check
    name_cols = []
    if '_computed_name' in reg_players.columns:
        name_cols.append('_computed_name')
    name_cols.extend([c for c in reg_players.columns if 'name' in str(c).lower() and c != '_computed_name'])
    
    if not name_cols:
        name_cols = [reg_players.columns[0]]
        
    # Women's Cara Murray override
    if domain == "Women's":
        for c in name_cols:
            cara_murray = reg_players[reg_players[c].astype(str).str.lower() == 'cara murray']
            if not cara_murray.empty:
                cara_waringstown = cara_murray.copy()
                cara_waringstown['Individual Membership Primary Club'] = 'Waringstown Cricket Club'
                reg_players = pd.concat([reg_players, cara_waringstown], ignore_index=True)
                break
    
    for _, r in reg_players.iterrows():
        reg_club = None
        
        # STRICT TARGET: Explicitly look for the exact column name verified in the Registry
        if 'Individual Membership Primary Club' in reg_players.columns and pd.notna(r['Individual Membership Primary Club']):
            val = str(r['Individual Membership Primary Club']).strip()
            if val.lower() != 'nan' and val != '':
                reg_club = val
        
        # Fallback keyword scan if the exact column is missing
        if not reg_club:
            for keyword in ['Primary Club', 'Transfer', 'Wylie', 'Club']:
                cols = [c for c in reg_players.columns if keyword in str(c)]
                for col in cols:
                    if pd.notna(r[col]) and str(r[col]).strip() and str(r[col]).lower() != 'nan':
                        reg_club = str(r[col]).strip()
                        break
                if reg_club: break

        # Map the found club to the player's name and alias variations
        if reg_club:
            for c in name_cols:
                if pd.notna(r[c]):
                    r_name = str(r[c]).replace('‡', '').strip().lower()
                    norm_name = re.sub(r'\s+', ' ', r_name)
                    if norm_name and norm_name != 'nan':
                        mapped_name = alias_map.get(norm_name, norm_name)
                        club_map[mapped_name] = reg_club
                        club_map[norm_name] = reg_club
                        
    return club_map

def build_league_dict(league_structure):
    league_dict = {}
    original_league_order = [] 
    
    team_col = next((col for col in league_structure.columns if 'team' in str(col).lower() or 'club' in str(col).lower()), league_structure.columns[0])
    league_col = next((col for col in league_structure.columns if 'league' in str(col).lower() or 'division' in str(col).lower()), league_structure.columns[1])

    for _, row in league_structure.iterrows():
        raw_team = str(row[team_col]).strip()
        raw_league = str(row[league_col]).strip()
        if raw_team and raw_team.lower() != 'nan' and raw_league and raw_league.lower() != 'nan':
            league_dict[raw_team] = raw_league
            if raw_league not in original_league_order:
                original_league_order.append(raw_league)
                
    return league_dict, list(league_dict.keys()), original_league_order

def extract_xi(team_str):
    match = re.search(r'((?:mw\d?|\d(?:st|nd|rd|th))\s*xi)', str(team_str).lower())
    if match: return match.group(1).replace(' ', '')
    return None

def clean_team_for_compare(t, domain):
    t = str(t).lower()
    t = re.sub(r'\bcc\b|\bcricket club\b', '', t)
    t = t.replace('1881', '').replace('ciyms', 'ci').replace('nimacc', 'nima').replace('dungannnon', 'dungannon')
    t = re.sub(r'(?i)drumaness\s+super\s*kings', 'drumaness', t)
    
    if domain == "Women's":
        t = re.sub(r'\bwomen\'s\b|\bwomens\b|\bwomen\b', '', t)
        
    t = re.sub(r'\b1st\s*xi\b|\b1st\b', '1', t)
    t = re.sub(r'\b2nd\s*xi\b|\b2nd\b', '2', t)
    t = re.sub(r'\b3rd\s*xi\b|\b3rd\b', '3', t)
    t = re.sub(r'\b4th\s*xi\b|\b4th\b', '4', t)
    t = re.sub(r'\b5th\s*xi\b|\b5th\b', '5', t)
    t = re.sub(r'\b6th\s*xi\b|\b6th\b', '6', t)
    t = re.sub(r'\bxi\b', '', t)
    
    return " ".join(t.split())

def get_team_league(team_name, team_keys, league_dict, domain):
    clean_search_name = re.sub(r',.*', '', str(team_name)).strip()
    for k in team_keys:
        if k.lower() == clean_search_name.lower(): return league_dict[k]
    team_clean = clean_team_for_compare(clean_search_name, domain)
    for k in team_keys:
        if clean_team_for_compare(k, domain) == team_clean: return league_dict[k]
    if not extract_xi(clean_search_name):
        bare_fallback = f"{clean_search_name} 1st XI"
        fallback_clean = clean_team_for_compare(bare_fallback, domain)
        for k in team_keys:
            if clean_team_for_compare(k, domain) == fallback_clean: return league_dict[k]
    team_xi = extract_xi(team_clean)
    best_match, best_score = None, 0
    for k in team_keys:
        k_clean = clean_team_for_compare(k, domain)
        k_xi = extract_xi(k_clean)
        if team_xi == k_xi or team_xi is None or k_xi is None:
            score = fuzz.token_sort_ratio(team_clean, k_clean)
            if score > best_score and score >= 75:
                best_score, best_match = score, k
    return league_dict.get(best_match)

def format_display_team(team_str, domain):
    if pd.isna(team_str): return "Unknown"
    c = str(team_str)
    if domain == "Midweek":
        c = re.sub(r'(?i)\bmw1\s*xi\b', ' 1', c)
        c = re.sub(r'(?i)\bmw2\s*xi\b', ' 2', c)
        c = re.sub(r'(?i)\bmw\s*xi\b', '', c)
        c = c.replace(' 1st XI', ' 1').replace(' 1st', ' 1')
    elif domain == "Women's":
        c = re.sub(r'(?i)\bwomen\'s\b|\bwomens\b|\bwomen\b', '', c)
        c = c.replace(' 1st XI', '').replace(' 1st', '')
    else:
        c = c.replace(' 1st XI', '').replace(' 1st', '')
    c = c.replace(' 2nd XI', ' 2').replace(' 3rd XI', ' 3')
    c = c.replace(' 4th XI', ' 4').replace(' 5th XI', ' 5').replace(' 6th XI', ' 6')
    c = c.replace(' XI', '')
    c = re.sub(r'\s+', ' ', c).strip()
    if domain != "Midweek" and c.endswith(' 1'):
        c = c[:-2].strip()
    if 'Holywood' in c and '1881' not in c:
        c = c.replace('Holywood', 'Holywood 1881')
    if domain == "Midweek" and not c.endswith(" MW"):
        c = f"{c} MW"
    return c

def extract_teams_from_group(group_str):
    try:
        parts = str(group_str).strip().rsplit(' - ', 1)
        rest = parts[0].strip()
        if ' v ' in rest:
            t1, remainder = rest.split(' v ', 1)
            t2 = remainder.rsplit(', ', 1)[0] if ', ' in remainder else remainder.rsplit(' - ', 1)[0]
            return t1.strip(), t2.strip()
        return rest, "Unknown"
    except: return "Unknown", "Unknown"

def clean_club_for_matching(club_str):
    if pd.isna(club_str): return ""
    c = str(club_str).lower()
    c = re.sub(r'\bcricket club\b|\bcc\b', '', c)
    c = c.replace('ciyms', 'ci')
    c = re.sub(r'(?i)northern\s+ireland\s+malayali\s+association', 'nimacc', c)
    c = re.sub(r'(?i)\bnima\s*cc\b|\bnimacc\b|\bnima\b', 'nimacc', c)
    c = re.sub(r'(?i)belfast\s+international\s+sports\s+club|belfast\s+b\.i\.s\.c\.', 'bisc', c)
    c = re.sub(r'(?i)civil\s+service\s+north\s+of\s+ireland|civil\s+service\s+north', 'csni', c)
    c = re.sub(r'(?i)drumaness\s+super\s*kings', 'drumaness', c)
    c = re.sub(r'(?i)donaghcloney', 'donacloney', c)
    return " ".join(c.split())

def determine_player_team_for_row(row, player_club_map, domain):
    player = str(row['Cleaned Name']).strip()
    t1, t2 = extract_teams_from_group(row['Group'])
    
    # 1. Check the row's scorecard 'Team' column directly to sidestep duplicate maps
    row_team = str(row.get('Team', '')).strip().lower()
    if row_team and row_team != 'nan':
        clean_row_t = clean_team_for_compare(row_team, domain)
        clean_t1 = clean_team_for_compare(t1, domain)
        clean_t2 = clean_team_for_compare(t2, domain)
        if clean_row_t in clean_t1 or clean_t1 in clean_row_t: return t1
        if clean_row_t in clean_t2 or clean_t2 in clean_row_t: return t2

    if domain != "Women's":
        if player.lower() in ['neil brand', 'sandeep singh']:
            return t1 if 'muckamore' in t1.lower() else (t2 if 'muckamore' in t2.lower() else t1)
        if player.lower() == 'james shannon':
            if 'holywood' in t1.lower() or 'holywood' in t2.lower():
                return t1 if 'holywood' in t1.lower() else t2
            elif 'saintfield' in t1.lower() or 'saintfield' in t2.lower():
                return t1 if 'saintfield' in t1.lower() else t2
                
    reg_club = player_club_map.get(player.lower())
    if reg_club and reg_club.lower() != 'nan':
        clean_reg = clean_club_for_matching(reg_club)
        clean_t1 = clean_club_for_matching(t1)
        clean_t2 = clean_club_for_matching(t2)
        if clean_reg in clean_t1 or clean_t1 in clean_reg: return t1
        if clean_reg in clean_t2 or clean_t2 in clean_reg: return t2
        reg_tokens = set(clean_reg.split())
        stop_words = {'1st', '2nd', '3rd', '4th', '5th', '6th', '7th', '8th', 'xi', '1', '2', '3', '4', '5', '6', '7', '8', 'women', 'womens', "women's", 'mw', 'mw1', 'mw2', 'mw3'}
        t1_tokens_clean = set(clean_t1.split()) - stop_words
        t2_tokens_clean = set(clean_t2.split()) - stop_words
        if len(reg_tokens.intersection(t1_tokens_clean)) > len(reg_tokens.intersection(t2_tokens_clean)): return t1
        elif len(reg_tokens.intersection(t2_tokens_clean)) > len(reg_tokens.intersection(t1_tokens_clean)): return t2
    return t1

def parse_high_score(scores_series):
    best_score = 0
    is_not_out = False
    for hs in scores_series.dropna().astype(str):
        if hs.lower() == 'nan': continue
        val = hs.replace('*', '').replace('.0', '')
        try:
            val_int = int(val)
            if val_int > best_score:
                best_score = val_int
                is_not_out = '*' in hs
            elif val_int == best_score and '*' in hs:
                is_not_out = True
        except ValueError: pass
    if best_score == 0 and not is_not_out: return "0"
    return f"{best_score}*" if is_not_out else str(best_score)

def calculate_averages(batting_df, bowling_df, player_club_map, team_keys, league_dict, domain, bat_sort="Runs", bowl_sort="Wickets"):
    for col in ['Matches', 'Innings', 'Not Outs', 'Runs', 'Balls', 'Fours', 'Sixes']:
        if col in batting_df.columns: batting_df[col] = pd.to_numeric(batting_df[col], errors='coerce').fillna(0)
    for col in ['Innings', 'Balls', 'Maidens', 'Runs', 'Wickets']:
        if col in bowling_df.columns: bowling_df[col] = pd.to_numeric(bowling_df[col], errors='coerce').fillna(0)
            
    batting_df['Team Played For'] = batting_df.apply(lambda r: determine_player_team_for_row(r, player_club_map, domain), axis=1)
    bowling_df['Team Played For'] = bowling_df.apply(lambda r: determine_player_team_for_row(r, player_club_map, domain), axis=1)
    
    def apply_league(t):
        league = get_team_league(t, team_keys, league_dict, domain)
        if not league: return "Unassigned"
        if domain == "Midweek": return str(league)
        league_str = str(league)
        target_words = ['premier', 'senior league 1', 'senior league 2', 'senior league 3'] if domain == "Men's" else ['premier', 'senior league 1', 'senior league 2', 'senior league 3', 'senior']
        if any(word in league_str.lower() for word in target_words): return league_str.replace('NCU', 'Mercury')
        return league_str

    batting_df['League'] = batting_df['Team Played For'].apply(apply_league)
    bowling_df['League'] = bowling_df['Team Played For'].apply(apply_league)

    def group_batting(df_to_group):
        grouped = df_to_group.groupby(['League', 'Cleaned Name']).agg({
            'Name': lambda x: x.value_counts().index[0] if not x.empty else "Unknown",
            'Team Played For': lambda x: format_display_team(x.value_counts().index[0], domain) if not x.empty else "Unknown",
            'Matches': 'sum', 'Innings': 'sum', 'Not Outs': 'sum', 'Runs': 'sum',
            'Balls': 'sum', 'Fours': 'sum', 'Sixes': 'sum', 'High Score': parse_high_score
        }).reset_index()
        grouped.rename(columns={'Team Played For': 'Team', 'Name': 'Player'}, inplace=True)
        grouped.drop(columns=['Cleaned Name'], inplace=True)
        outs = grouped['Innings'] - grouped['Not Outs']
        grouped['Average'] = np.where(outs > 0, grouped['Runs'] / outs, np.nan)
        grouped['Strike Rate'] = np.where(grouped['Balls'] > 0, (grouped['Runs'] / grouped['Balls']) * 100, np.nan)
        return grouped[['League', 'Player', 'Team', 'Matches', 'Innings', 'Not Outs', 'Runs', 'Balls', 'Fours', 'Sixes', 'High Score', 'Average', 'Strike Rate']]

    def group_bowling(df_to_group, bat_avgs):
        sorted_df = df_to_group.sort_values(by=['Wickets', 'Runs'], ascending=[False, True])
        best_spells = sorted_df.drop_duplicates(subset=['League', 'Cleaned Name'])
        best_spells['Best Bowling'] = best_spells['Wickets'].fillna(0).astype(int).astype(str) + '-' + best_spells['Runs'].fillna(0).astype(int).astype(str)
        bbi_series = best_spells.set_index(['League', 'Cleaned Name'])['Best Bowling']
        
        grouped = df_to_group.groupby(['League', 'Cleaned Name']).agg({
            'Bowler': lambda x: x.value_counts().index[0] if not x.empty else "Unknown",
            'Team Played For': lambda x: format_display_team(x.value_counts().index[0], domain) if not x.empty else "Unknown",
            'Innings': 'sum', 'Balls': 'sum', 'Maidens': 'sum', 'Runs': 'sum', 'Wickets': 'sum'
        }).reset_index()
        grouped = grouped.merge(bbi_series, on=['League', 'Cleaned Name'], how='left')
        grouped.rename(columns={'Team Played For': 'Team', 'Bowler': 'Player'}, inplace=True)
        grouped.drop(columns=['Cleaned Name'], inplace=True)
        
        total_matches = bat_avgs[['League', 'Player', 'Matches']].rename(columns={'Matches': 'Total_Matches'})
        grouped = grouped.merge(total_matches, on=['League', 'Player'], how='left')
        grouped['Matches'] = grouped['Total_Matches'].fillna(grouped['Innings']).astype(int)
        grouped['Overs'] = (grouped['Balls'] // 6) + (grouped['Balls'] % 6) / 10
        grouped['Average'] = np.where(grouped['Wickets'] > 0, grouped['Runs'] / grouped['Wickets'], np.nan)
        grouped['Economy'] = np.where(grouped['Balls'] > 0, (grouped['Runs'] / grouped['Balls']) * 6, np.nan)
        grouped['Strike Rate'] = np.where(grouped['Wickets'] > 0, grouped['Balls'] / grouped['Wickets'], np.nan)
        return grouped[['League', 'Player', 'Team', 'Matches', 'Innings', 'Overs', 'Maidens', 'Runs', 'Wickets', 'Best Bowling', 'Average', 'Economy', 'Strike Rate']]

    if domain == "Midweek":
        leagues_bat = group_batting(batting_df)
        overall_bat_df = batting_df.copy()
        overall_bat_df['League'] = 'Overall Midweek'
        overall_bat = group_batting(overall_bat_df)
        batting_final = pd.concat([leagues_bat, overall_bat], ignore_index=True)
        
        leagues_bowl = group_bowling(bowling_df, batting_final)
        overall_bowl_df = bowling_df.copy()
        overall_bowl_df['League'] = 'Overall Midweek'
        overall_bowl = group_bowling(overall_bowl_df, batting_final)
        bowling_final = pd.concat([leagues_bowl, overall_bowl], ignore_index=True)
    else:
        batting_final = group_batting(batting_df)
        bowling_final = group_bowling(bowling_df, batting_final)
        
    if bat_sort == "Average": batting_final = batting_final.sort_values(by=['League', 'Average', 'Runs'], ascending=[True, False, False])
    elif bat_sort == "Strike Rate": batting_final = batting_final.sort_values(by=['League', 'Strike Rate', 'Runs'], ascending=[True, False, False])
    else: batting_final = batting_final.sort_values(by=['League', 'Runs', 'Average'], ascending=[True, False, False])

    if bowl_sort == "Average": bowling_final = bowling_final.sort_values(by=['League', 'Average', 'Wickets'], ascending=[True, True, False], na_position='last')
    elif bowl_sort == "Economy": bowling_final = bowling_final.sort_values(by=['League', 'Economy', 'Wickets'], ascending=[True, True, False], na_position='last')
    elif bowl_sort == "Strike Rate": bowling_final = bowling_final.sort_values(by=['League', 'Strike Rate', 'Wickets'], ascending=[True, True, False], na_position='last')
    else: bowling_final = bowling_final.sort_values(by=['League', 'Wickets', 'Average'], ascending=[True, False, True], na_position='last')
    
    if domain in ["Women's", "Midweek"]:
        batting_final['Average'] = batting_final['Average'].round(2)
        batting_final['Strike Rate'] = batting_final['Strike Rate'].round(2)
        bowling_final['Average'] = bowling_final['Average'].round(2)
        bowling_final['Economy'] = bowling_final['Economy'].round(2)
        bowling_final['Strike Rate'] = bowling_final['Strike Rate'].round(2)

    return batting_final, bowling_final

def custom_league_sort(league_name, domain, ordered_leagues=None):
    name_lower = str(league_name).lower()
    if domain == "Midweek" and 'overall' in name_lower: return (0, 0, '')
    if ordered_leagues and league_name in ordered_leagues: return (1, ordered_leagues.index(league_name), '')
    if domain == "Midweek":
        if 'group a' in name_lower: return (2, 1, league_name)
        elif 'group b' in name_lower: return (2, 2, league_name)
        elif 'group c' in name_lower: return (2, 3, league_name)
        return (3, 0, league_name)
    else:
        if 'premier' in name_lower: return (2, 0, '')
        elif 'senior' in name_lower or 'section' in name_lower:
            match = re.search(r'senior(?: league)? (\d+)', name_lower)
            return (3, int(match.group(1)) if match else 99, '')
        elif 'junior' in name_lower:
            match = re.search(r'junior(?: league)? (\d+)([a-z]?)', name_lower)
            return (4, int(match.group(1)) if match else 99, match.group(2) if match else '')
        return (5, 0, name_lower)

def format_excel_sheet(writer, df, sheet_name, min_label=None):
    safe_sheet_name = str(sheet_name).replace("League", "Lge").replace("Midweek", "MW").replace("Group", "Grp").replace("Overall", "Ovr").strip()[:31].strip()
    df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
    worksheet = writer.sheets[safe_sheet_name]
    workbook = writer.book
    
    left_header = workbook.add_format({'bold': True, 'bottom': 1, 'bg_color': '#D9D9D9', 'align': 'left'})
    center_header = workbook.add_format({'bold': True, 'bottom': 1, 'bg_color': '#D9D9D9', 'align': 'center'})
    bold_name, left_align, center_align = workbook.add_format({'bold': True}), workbook.add_format({'align': 'left'}), workbook.add_format({'align': 'center'})
    two_decimals = workbook.add_format({'num_format': '0.00', 'align': 'center'})
    
    for col_num, col_name in enumerate(df.columns):
        worksheet.write(0, col_num, col_name, left_header if col_name in ['Player', 'Team'] else center_header)
        col_width = max(max((len(str(x)) for x in df[col_name]), default=0), len(str(col_name))) + 2
        
        if col_name == 'Player': worksheet.set_column(col_num, col_num, col_width, bold_name)
        elif col_name == 'Team': worksheet.set_column(col_num, col_num, col_width, left_align)
        elif col_name in ['Average', 'Strike Rate', 'Economy']: worksheet.set_column(col_num, col_num, col_width, two_decimals)
        else: worksheet.set_column(col_num, col_num, col_width, center_align)
            
    if min_label:
        worksheet.write(len(df) + 2, 0, min_label, workbook.add_format({'italic': True, 'bold': True}))

# ==========================================
# WORD DOC GENERATOR FUNCTIONS
# ==========================================
def doc_format_cricket_names(text, domain):
    if pd.isna(text): return text
    text = str(text)
    text = text.replace('NCU', 'Mercury').replace('Mercury Pathway', 'NCU Pathway')
    text = text.replace('CIYMS', 'CI')
    text = re.sub(r'(?i)\bnima\s*cc\b|\bnimacc\b|\bnima\b', 'NIMACC', text)
    text = re.sub(r'(?i)belfast\s+international\s+sports\s+club|belfast\s+b\.i\.s\.c\.', 'BISC', text)
    text = re.sub(r'(?i)civil\s+service\s+north\s+of\s+ireland|civil\s+service\s+north', 'CSNI', text)
    text = re.sub(r'(?i)drumaness\s+super\s*kings', 'Drumaness Superkings', text)
    text = text.replace('Donaghcloney', 'Donacloney')
    if 'Holywood' in text and '1881' not in text: text = text.replace('Holywood', 'Holywood 1881')
    
    if domain == "Midweek": text = text.replace(' 1st XI', ' 1').replace(' 1st', ' 1')
    else: text = text.replace(' 1st XI', '').replace(' 1st', '')
        
    text = text.replace(' 2nd XI', ' 2').replace(' 3rd XI', ' 3')
    text = text.replace(' 4th XI', ' 4').replace(' 5th XI', ' 5').replace(' 6th XI', ' 6')
    text = text.replace(' XI', '') 
    text = re.sub(r'\s+', ' ', text).strip()
    
    if domain != "Midweek" and text.endswith(' 1'):
        text = text[:-2].strip()
    return text

def doc_get_player_team_from_match(match_str, base_club):
    if pd.isna(match_str) or ' v ' not in str(match_str): return "Unknown Team"
    parts = str(match_str).split(' v ')
    team1, team2 = parts[0].strip(), parts[1].split(',')[0].strip()
    
    def clean_for_matching(s):
        s = str(s).lower()
        s = re.sub(r'northern\s+ireland\s+malayali\s+association', 'nimacc', s)
        s = re.sub(r'\bnima\s*cc\b|\bnimacc\b|\bnima\b', 'nimacc', s)
        s = re.sub(r'belfast\s+international\s+sports\s+club|belfast\s+b\.i\.s\.c\.', 'bisc', s)
        s = re.sub(r'civil\s+service\s+north\s+of\s+ireland|civil\s+service\s+north', 'csni', s)
        s = re.sub(r'drumaness\s+super\s*kings', 'drumaness', s)
        s = re.sub(r'donaghcloney', 'donacloney', s)
        s = re.sub(r'\b(cricket club|club|teams|cc|1st|2nd|3rd|4th|5th|6th|1|2|3|4|5|6|xi)\b', '', s)
        return set(s.split())
        
    t1_words, t2_words, club_words = clean_for_matching(team1), clean_for_matching(team2), clean_for_matching(base_club)
    for target in ['nimacc', 'bisc', 'csni', 'drumaness', 'donacloney']:
        if target in club_words:
            if target in t1_words: return team1
            if target in t2_words: return team2
    if len(t1_words.intersection(club_words)) > len(t2_words.intersection(club_words)): return team1
    elif len(t2_words.intersection(club_words)) > len(t1_words.intersection(club_words)): return team2
    
    if str(base_club).lower() in team1.lower(): return team1
    if str(base_club).lower() in team2.lower(): return team2
    return "Other Team"

def doc_team_sort_key(team_name):
    words = team_name.split()
    if words and words[-1].isdigit(): return (team_name.rsplit(' ', 1)[0], int(words[-1]))
    return (team_name, 1)

def add_bullet_point(doc, text, level=1, space_after=0):
    style_name = 'List Bullet' if level == 1 else f'List Bullet {level}'
    try:
        p = doc.add_paragraph(style=style_name)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 0.9
        run = p.add_run(text)
        run.font.name, run.font.size = 'Calibri', Pt(11) 
    except KeyError:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 0.9
        p.paragraph_format.left_indent = Pt(18 * level)
        run = p.add_run(f"{'·' if level == 1 else 'o'}\t{text}")
        run.font.name, run.font.size = 'Calibri', Pt(10) 

def add_custom_heading(doc, text, level):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6) if level > 1 else Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 0.9
    if level == 1:
        p.style = doc.styles['Heading 1']
        run = p.add_run(text)
        run.font.name, run.font.size = 'Calibri', Pt(14)
    elif level == 2:
        p.style = doc.styles['Heading 2']
        run = p.add_run(text)
        run.font.name, run.font.size = 'Calibri', Pt(12)
    elif level == 3:
        p.style = doc.styles['Heading 3']
        run = p.add_run(text)
        run.font.name, run.font.size = 'Calibri', Pt(10)
    return p

# ==========================================
# ALIAS LOOKUP HELPERS
# ==========================================
def get_player_aliases(official_name, aliases):
    """
    Retrieves all known scorecard/match aliases mapped to an official registered name.
    """
    if aliases is None or aliases.empty:
        return []
    
    clean_official = official_name.strip().lower()
    found_aliases = []
    
    if 'Input Name (Scorecard/Stats)' in aliases.columns and 'Official Registered Name' in aliases.columns:
        match_rows = aliases[aliases['Official Registered Name'].astype(str).str.strip().str.lower() == clean_official]
        for val in match_rows['Input Name (Scorecard/Stats)'].dropna().unique():
            cleaned_val = str(val).replace('‡', '').strip()
            if cleaned_val and cleaned_val.lower() != clean_official and cleaned_val.lower() != 'nan':
                if cleaned_val not in found_aliases:
                    found_aliases.append(cleaned_val)
    else:
        for _, row in aliases.iterrows():
            alias_val = str(row.iloc[0]).replace('‡', '').strip()
            off_val = str(row.iloc[1]).replace('‡', '').strip()
            if off_val.lower() == clean_official and alias_val.lower() != clean_official and alias_val.lower() != 'nan':
                if alias_val not in found_aliases:
                    found_aliases.append(alias_val)
                    
    return found_aliases


def generate_single_player_doc(active_player, player_batting, player_bowling, reg_players_df, domain, aliases_list=None):
    club_name = "Unknown_Club"
    if active_player.lower() == 'neil brand' and domain != "Women's":
        club_name = 'Muckamore'
    else:
        reg_search_name = active_player.split(' (')[0] if ' (' in active_player else active_player
        reg_match = reg_players_df[reg_players_df['Full Name'].astype(str).str.lower() == reg_search_name.lower()]
        if not reg_match.empty:
            for keyword in ['Wylie', 'Transfer']:
                cols = [c for c in reg_match.columns if keyword in str(c)]
                if cols and len(reg_match[cols[0]].dropna().values) > 0 and str(reg_match[cols[0]].dropna().values[0]).strip() != '':
                    club_name = str(reg_match[cols[0]].dropna().values[0]).strip()
                    break
            if club_name == "Unknown_Club":
                primary_cols = [c for c in reg_match.columns if 'Primary Club' in str(c) and 'Wylie' not in str(c)]
                if primary_cols and len(reg_match[primary_cols[0]].dropna().values) > 0 and str(reg_match[primary_cols[0]].dropna().values[0]).strip() != '': 
                    club_name = str(reg_match[primary_cols[0]].dropna().values[0]).strip()
                    
    if club_name == "Unknown_Club":
        all_groups_fallback = pd.concat([player_batting['Group'], player_bowling['Group']]).dropna().tolist()
        team_frequency = {}
        for grp in all_groups_fallback:
            if ' v ' in grp:
                t1, t2 = grp.split(' v ')[0].strip(), grp.split(' v ')[1].split(',')[0].strip()
                t1, t2 = doc_format_cricket_names(t1, domain), doc_format_cricket_names(t2, domain)
                team_frequency[t1] = team_frequency.get(t1, 0) + 1
                team_frequency[t2] = team_frequency.get(t2, 0) + 1
        if team_frequency: club_name = sorted(team_frequency.items(), key=lambda x: x[1], reverse=True)[0][0]
    
    if ' (' in active_player: club_name = active_player.split(' (')[1].replace(')', '')
    club_name_clean = doc_format_cricket_names(club_name, domain)

    if not player_batting.empty: player_batting['Team'] = player_batting['Group'].apply(lambda x: doc_get_player_team_from_match(x, club_name_clean))
    if not player_bowling.empty: player_bowling['Team'] = player_bowling['Group'].apply(lambda x: doc_get_player_team_from_match(x, club_name_clean))
        
    unique_teams = set()
    if not player_batting.empty: unique_teams.update(player_batting['Team'].unique())
    if not player_bowling.empty: unique_teams.update(player_bowling['Team'].unique())
    unique_teams = sorted(list(unique_teams), key=doc_team_sort_key)

    all_groups = []
    if not player_batting.empty: all_groups.extend(player_batting['Group'].tolist())
    if not player_bowling.empty: all_groups.extend(player_bowling['Group'].tolist())
    unique_groups = list(dict.fromkeys(all_groups)) 
    
    def extract_match_date(grp_str):
        try:
            parts = str(grp_str).rsplit(' - ', 1)
            if len(parts) == 2:
                d_str = parts[1].strip()
                d_str = re.sub(r'(st|nd|rd|th)\b', '', d_str, flags=re.IGNORECASE)
                dt = pd.to_datetime(d_str, dayfirst=True, errors='coerce')
                if pd.notna(dt): return dt
        except: pass
        try:
            match = re.search(r'(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{1,2}\s+[A-Za-z]{3,}\s+\d{4})', str(grp_str))
            if match:
                dt = pd.to_datetime(match.group(1), dayfirst=True, errors='coerce')
                if pd.notna(dt): return dt
        except: pass
        return pd.Timestamp.min

    unique_groups.sort(key=extract_match_date)
    
    matches_by_team = {}
    for grp in unique_groups:
        team_played_for = doc_get_player_team_from_match(grp, club_name_clean)
        
        b_row = player_batting[player_batting['Group'] == grp] if not player_batting.empty else pd.DataFrame()
        if not b_row.empty and b_row.iloc[0]['Innings'] > 0:
            hs = b_row.iloc[0]['High Score']
            hs_str = str(hs).replace('.0', '') if pd.notna(hs) else str(int(b_row.iloc[0]['Runs']))
            bat_str = f"Batting: {hs_str} runs"
        else: bat_str = "Batting: Did not bat"
            
        bw_row = player_bowling[player_bowling['Group'] == grp] if not player_bowling.empty else pd.DataFrame()
        if not bw_row.empty and bw_row.iloc[0]['Innings'] > 0 and bw_row.iloc[0]['Overs'] > 0:
            o = bw_row.iloc[0]['Overs']
            o_str = str(o).replace('.0', '') if str(o).endswith('.0') else str(o)
            m = int(bw_row.iloc[0]['Maidens']) if pd.notna(bw_row.iloc[0]['Maidens']) else 0
            r = int(bw_row.iloc[0]['Runs']) if pd.notna(bw_row.iloc[0]['Runs']) else 0
            w = int(bw_row.iloc[0]['Wickets']) if pd.notna(bw_row.iloc[0]['Wickets']) else 0
            bowl_str = f"Bowling: {o_str}-{m}-{r}-{w}"
        else: bowl_str = "Bowling: Did not bowl"
            
        if team_played_for not in matches_by_team: matches_by_team[team_played_for] = []
        matches_by_team[team_played_for].append({'match': grp, 'bat_str': bat_str, 'bowl_str': bowl_str})

    doc = Document()
    style_normal = doc.styles['Normal']
    style_normal.font.name, style_normal.font.size = 'Calibri', Pt(11) 
    
    header_club_name = re.sub(r'(?i)\s*cricket club', '', club_name_clean).strip()
    display_player_name = active_player.split(' (')[0].title()
    
    # Format dual-name header for Word File
    if aliases_list:
        alias_str = " / ".join(aliases_list)
        heading_title = f"{display_player_name} (Registered Name) / {alias_str} (Match Name) - {header_club_name} - Season Summary\n"
    else:
        heading_title = f"{display_player_name} - {header_club_name} - Season Summary\n"
        
    add_custom_heading(doc, heading_title, level=1)
    
    add_custom_heading(doc, "Batting Statistics", level=2)

    batting_found = False
    for team in unique_teams:
        team_bat = player_batting[player_batting['Team'] == team] if not player_batting.empty else pd.DataFrame()
        if team_bat.empty or team_bat['Matches'].sum() == 0: continue
            
        batting_found = True
        team_bowl_for_matches = player_bowling[player_bowling['Team'] == team] if not player_bowling.empty else pd.DataFrame()
        bowl_m_raw = int(team_bowl_for_matches['Matches'].sum()) if not team_bowl_for_matches.empty else 0
        bat_matches = max(int(team_bat['Matches'].sum()), bowl_m_raw)
        
        bat_innings = int(team_bat['Innings'].sum())
        total_runs = int(team_bat['Runs'].sum())
        not_outs = int(team_bat['Not Outs'].sum()) if 'Not Outs' in team_bat.columns else 0
        balls_faced = int(team_bat['Balls'].sum()) if 'Balls' in team_bat.columns else 0
        fours, sixes = (int(team_bat['Fours'].sum()) if 'Fours' in team_bat.columns else 0), (int(team_bat['Sixes'].sum()) if 'Sixes' in team_bat.columns else 0)
        
        best_score, is_not_out = 0, False
        for hs in team_bat['High Score'].dropna().astype(str).tolist():
            if hs.lower() == 'nan': continue
            val = hs.replace('*', '').replace('.0', '')
            try:
                if int(val) > best_score: best_score, is_not_out = int(val), '*' in hs
                elif int(val) == best_score and '*' in hs: is_not_out = True
            except ValueError: pass
        
        high_score_str = f"{best_score}*" if is_not_out else str(best_score)
        if bat_innings == 0: high_score_str = "N/A"
        outs = bat_innings - not_outs
        bat_avg = f"{total_runs / outs:.2f}" if outs > 0 else "N/A"

        add_custom_heading(doc, team, level=3)
        add_bullet_point(doc, f"Matches: {bat_matches}")
        add_bullet_point(doc, f"Innings: {bat_innings}")
        add_bullet_point(doc, f"Total Runs: {total_runs}")
        add_bullet_point(doc, f"Highest Score: {high_score_str}")
        add_bullet_point(doc, f"Batting Average: {bat_avg}")
        add_bullet_point(doc, f"Balls Faced: {balls_faced}")
        add_bullet_point(doc, f"Boundaries: {fours} Fours, {sixes} Sixes\n")

    if not batting_found:
        p = doc.add_paragraph("No batting statistics recorded.\n")
        p.style.font.name, p.style.font.size = 'Calibri', Pt(11)
        
    add_custom_heading(doc, "Bowling Statistics", level=2)
    bowling_found = False
    for team in unique_teams:
        team_bowl = player_bowling[player_bowling['Team'] == team] if not player_bowling.empty else pd.DataFrame()
        if team_bowl.empty or team_bowl['Matches'].sum() == 0: continue
            
        bowling_found = True
        team_bat_for_matches = player_batting[player_batting['Team'] == team] if not player_batting.empty else pd.DataFrame()
        bat_m_raw = int(team_bat_for_matches['Matches'].sum()) if not team_bat_for_matches.empty else 0
        bowl_matches = max(bat_m_raw, int(team_bowl['Matches'].sum()))
        
        bowl_innings = int(team_bowl['Innings'].sum())
        overs, maidens = round(team_bowl['Overs'].sum(), 1), int(team_bowl['Maidens'].sum())
        bowl_runs, wickets = int(team_bowl['Runs'].sum()), int(team_bowl['Wickets'].sum())
        bowl_avg = f"{bowl_runs / wickets:.2f}" if wickets > 0 else "N/A"

        add_custom_heading(doc, team, level=3)
        add_bullet_point(doc, f"Matches: {bowl_matches}")
        add_bullet_point(doc, f"Innings: {bowl_innings}")
        add_bullet_point(doc, f"Overs: {overs}")
        add_bullet_point(doc, f"Maidens: {maidens}")
        add_bullet_point(doc, f"Runs Conceded: {bowl_runs}")
        add_bullet_point(doc, f"Wickets: {wickets}")
        add_bullet_point(doc, f"Bowling Average: {bowl_avg}\n")

    if not bowling_found:
        p = doc.add_paragraph("No bowling statistics recorded.\n")
        p.style.font.name, p.style.font.size = 'Calibri', Pt(11)

    add_custom_heading(doc, "Fielding Statistics", level=2)
    fielding_found = False
    for team in unique_teams:
        team_bat = player_batting[player_batting['Team'] == team] if not player_batting.empty else pd.DataFrame()
        if team_bat.empty or team_bat['Matches'].sum() == 0: continue
            
        fielding_found = True
        catches = int(team_bat['Catches'].sum()) if 'Catches' in team_bat.columns else 0
        catches_wk = int(team_bat['Catches as Keeper'].sum()) if 'Catches as Keeper' in team_bat.columns else 0
        stumpings = int(team_bat['Stumpings'].sum()) if 'Stumpings' in team_bat.columns else 0
        run_outs = int(team_bat['Run Outs'].sum()) if 'Run Outs' in team_bat.columns else 0

        add_custom_heading(doc, team, level=3)
        field_stats_list = [f"Catches: {catches}"]
        if catches_wk > 0: field_stats_list.append(f"Catches as Keeper: {catches_wk}")
        if stumpings > 0: field_stats_list.append(f"Stumpings: {stumpings}")
        if run_outs > 0: field_stats_list.append(f"Run Outs: {run_outs}")
        
        for i, stat in enumerate(field_stats_list):
            if i == len(field_stats_list) - 1: add_bullet_point(doc, stat + "\n")
            else: add_bullet_point(doc, stat)

    if not fielding_found:
        p = doc.add_paragraph("No fielding statistics recorded.\n")
        p.style.font.name, p.style.font.size = 'Calibri', Pt(11)

    add_custom_heading(doc, "Match Appearances", level=2)
    for team in sorted(matches_by_team.keys(), key=doc_team_sort_key):
        add_custom_heading(doc, team, level=3)
        for m in matches_by_team[team]:
            add_bullet_point(doc, m['match'], level=1)
            add_bullet_point(doc, m['bat_str'], level=2)
            add_bullet_point(doc, m['bowl_str'], level=2, space_after=4)
            
    doc_io = io.BytesIO()
    doc.save(doc_io)
    
    filename_player = active_player.split(' (')[0].title().replace(' ', '_')
    filename = f"{filename_player}_{club_name_clean.replace(' ', '_')}_{datetime.now().strftime('%Y-%m-%d_%H%M%S')}.docx"
    
    return doc_io, filename

# ==========================================
# REGISTRATION ENGINE SPECIFIC FUNCTIONS
# ==========================================
def get_ordinal_date(dt, include_year=True):
    if pd.isna(dt): return "N/A"
    day = dt.day
    if 11 <= (day % 100) <= 13: suffix = 'th'
    else: suffix = ['th', 'st', 'nd', 'rd', 'th'][min(day % 10, 4)]
    month = dt.strftime('%B')
    if include_year: return f"{month} {day}{suffix}, {dt.year}"
    return f"{month} {day}{suffix}"

def doc_formal_team_name(team_name):
    team_str = str(team_name)
    if 'holywood' in team_str.lower() and '1881' not in team_str.lower():
        return re.sub(r'(?i)(holywood)', r'Holywood 1881', team_str)
    return team_str

def match_sort_key(match_tuple):
    m_d = match_tuple[3]
    date_key = m_d if pd.notna(m_d) else datetime.min
    league = match_tuple[2]
    parts = re.split(r'(\d+)', league)
    parts = [int(p) if p.isdigit() else p for p in parts]
    return (date_key, parts, match_tuple[0])
    
def violation_sort_key(record, name_field):
    m_date = record.get('Match Date')
    date_val = m_date if pd.notna(m_date) else pd.Timestamp.min
    
    full_name = str(record.get(name_field, '')).strip()
    parts = full_name.split()
    
    if len(parts) > 1:
        surname = parts[-1].lower()
        firstnames = " ".join(parts[:-1]).lower()
    elif len(parts) == 1:
        surname = parts[0].lower()
        firstnames = ""
    else:
        surname = ""
        firstnames = ""
        
    return (date_val, surname, firstnames)

def doc_add_bullet(doc, text):
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(text)
    run.font.name, run.font.size = 'Calibri', Pt(11) 
    
def get_competition_sort_key(comp, domain="Men's"):
    comp_lower = str(comp).lower().strip()
    
    if "women's gallagher challenge plate" in comp_lower or "womens gallagher challenge plate" in comp_lower:
        return (1, 12, comp_lower)
    elif "women's gallagher challenge cup" in comp_lower or "womens gallagher challenge cup" in comp_lower:
        return (1, 11, comp_lower)
    elif "gallagher challenge cup" in comp_lower:
        return (1, 0, comp_lower)
    elif "lvs t20 cup" in comp_lower:
        return (1, 1, comp_lower)
    elif "lvs t20 trophy" in comp_lower:
        return (1, 2, comp_lower)
    elif "junior 1 t20 plate" in comp_lower:
        return (1, 6, comp_lower)
    elif "junior cup" in comp_lower:
        return (1, 3, comp_lower)
    elif "lvs t20 bowl" in comp_lower:
        return (1, 4, comp_lower)
    elif "lvs t20 shield" in comp_lower:
        return (1, 5, comp_lower)
    elif "intermediate cup" in comp_lower:
        return (1, 7, comp_lower)
    elif "lindsay" in comp_lower or "minor (lindsay)" in comp_lower:
        return (1, 8, comp_lower)
    elif "minor qualifying cup" in comp_lower:
        return (1, 9, comp_lower)
    elif "development cup" in comp_lower:
        return (1, 10, comp_lower)
    elif "irish cup" in comp_lower or "irish senior cup" in comp_lower:
        return (1, 13, comp_lower)
    elif "national cup" in comp_lower:
        return (1, 14, comp_lower)
    elif "ulster plate" in comp_lower:
        return (1, 15, comp_lower)
    
    is_cup_fallback = any(word in comp_lower for word in ['cup', 'trophy', 'shield', 'plate', 'bowl', 'vase', 'challenge', 'fixture', 'possible cup match', 'derby', 'knockout'])
    if is_cup_fallback:
        return (1, 100, comp_lower)
    
    is_womens = (domain == "Women's") or any(word in comp_lower for word in ['women', 'womens', "women's"])
    
    if is_womens:
        if 'premier' in comp_lower:
            sub_rank = 1
        elif 'senior' in comp_lower or 'section 1' in comp_lower:
            sub_rank = 2
        elif 'junior' in comp_lower:
            match = re.search(r'junior(?: league)?\s*(\d+)([a-z]?)', comp_lower)
            if match:
                num = int(match.group(1))
                let = match.group(2)
                let_val = (ord(let) - ord('a') + 1) * 0.01 if let else 0.0
                sub_rank = 10 + num + let_val
            else:
                sub_rank = 50
        else:
            sub_rank = 100
        return (3, sub_rank, comp_lower)
    
    if 'premier' in comp_lower:
        return (2, 0, comp_lower)
    elif 'senior league 1' in comp_lower or 'senior 1' in comp_lower or 'senior league section 1' in comp_lower:
        return (2, 1, comp_lower)
    elif 'senior league 2' in comp_lower or 'senior 2' in comp_lower or 'senior league section 2' in comp_lower:
        return (2, 2, comp_lower)
    elif 'senior league 3' in comp_lower or 'senior 3' in comp_lower or 'senior league section 3' in comp_lower:
        return (2, 3, comp_lower)
    elif 'junior' in comp_lower:
        match = re.search(r'junior(?: league)?\s*(\d+)([a-z]?)', comp_lower)
        if match:
            num = int(match.group(1))
            let = match.group(2)
            let_val = (ord(let) - ord('a') + 1) * 0.01 if let else 0.0
            sub_rank = 10 + num + let_val  
            return (2, sub_rank, comp_lower)
        else:
            return (2, 99, comp_lower)
    elif 'midweek' in comp_lower or 'group' in comp_lower:
        match = re.search(r'group\s*([a-z])', comp_lower)
        sub_rank = (ord(match.group(1)) - ord('a')) if match else 0
        return (2, 200 + sub_rank, comp_lower)
        
    return (4, 0, comp_lower)

def render_grouped_matches(doc, matches, domain):
    by_date_ts = {}
    for t_a, t_b, lg, m_d in matches:
        if pd.isna(m_d):
            continue
        m_d_normalized = m_d.normalize()
        if m_d_normalized not in by_date_ts:
            by_date_ts[m_d_normalized] = {}
        
        comp_name = str(lg).strip()
        if comp_name not in by_date_ts[m_d_normalized]:
            by_date_ts[m_d_normalized][comp_name] = []
        
        by_date_ts[m_d_normalized][comp_name].append((t_a, t_b))
        
    for dt in sorted(by_date_ts.keys()):
        date_header = f"{dt.day} {dt.strftime('%B %Y')}"
        
        p_date = doc.add_paragraph()
        p_date.paragraph_format.space_before = Pt(6)
        p_date.paragraph_format.space_after = Pt(2)
        p_date.paragraph_format.line_spacing = 0.9
        run_date = p_date.add_run(date_header)
        run_date.bold = True
        run_date.font.name = 'Calibri'
        run_date.font.size = Pt(11)
        
        comps = by_date_ts[dt]
        sorted_comps = sorted(comps.keys(), key=lambda c: get_competition_sort_key(c, domain))
        
        for comp in sorted_comps:
            try:
                p_comp = doc.add_paragraph(style='List Bullet')
            except KeyError:
                p_comp = doc.add_paragraph()
                p_comp.paragraph_format.left_indent = Pt(18)
                p_comp.add_run("• ")
                
            p_comp.paragraph_format.space_before = Pt(2)
            p_comp.paragraph_format.space_after = Pt(0)
            p_comp.paragraph_format.line_spacing = 0.9
            
            run_comp = p_comp.add_run(comp)
            run_comp.font.name = 'Calibri'
            run_comp.font.size = Pt(11)
            
            sorted_matches = sorted(comps[comp], key=lambda x: (str(x[0]).lower(), str(x[1]).lower()))
            
            for t_a, t_b in sorted_matches:
                try:
                    p_match = doc.add_paragraph(style='List Bullet 2')
                except KeyError:
                    p_match = doc.add_paragraph()
                    p_match.paragraph_format.left_indent = Pt(36)
                    p_match.add_run("o ")
                    
                p_match.paragraph_format.space_before = Pt(0)
                p_match.paragraph_format.space_after = Pt(0)
                p_match.paragraph_format.line_spacing = 0.9
                
                match_text = f"{doc_formal_team_name(t_a)} v {doc_formal_team_name(t_b)}"
                run_match = p_match.add_run(match_text)
                run_match.font.name = 'Calibri'
                run_match.font.size = Pt(11) 

def export_and_format_excel(df, writer, sheet_name):
    df.to_excel(writer, index=False, sheet_name=sheet_name)
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    header_format = workbook.add_format({'bold': True, 'bottom': 1})
    bold_name_format = workbook.add_format({'bold': True})
    
    for col_num, col_name in enumerate(df.columns):
        worksheet.write(0, col_num, col_name, header_format)
        series_len = max((len(str(x)) for x in df[col_name]), default=0)
        header_len = len(str(col_name))
        max_width = max(series_len, header_len) + 2
        if col_name in ['Stats Name (Cleaned)', 'Player (Cleaned)']: worksheet.set_column(col_num, col_num, max_width, bold_name_format)
        else: worksheet.set_column(col_num, col_num, max_width)

def run_registration_audit(domain, start_date, end_date, f_reg, f_alias, f_starring, f_league, f_bat, f_bowl, f_irish_bat=None, f_irish_bowl=None, f_cup=None):
    registered_players = pd.read_excel(f_reg)
    aliases = pd.read_excel(f_alias)
    league_structure = pd.read_excel(f_league)
    batting_stats = pd.read_excel(f_bat)
    bowling_stats = pd.read_excel(f_bowl)

    batting_stats['Is_Irish_Match'] = False
    bowling_stats['Is_Irish_Match'] = False
    
    if f_irish_bat and os.path.exists(f_irish_bat):
        irish_bat = pd.read_excel(f_irish_bat)
        irish_bat['Is_Irish_Match'] = True
        batting_stats = pd.concat([batting_stats, irish_bat], ignore_index=True)
        
    if f_irish_bowl and os.path.exists(f_irish_bowl):
        irish_bowl = pd.read_excel(f_irish_bowl)
        irish_bowl['Is_Irish_Match'] = True
        bowling_stats = pd.concat([bowling_stats, irish_bowl], ignore_index=True)

    def parse_match_group(group_str):
        try:
            group_str = str(group_str).strip()
            parts = group_str.rsplit(' - ', 1)
            date_str = parts[1].strip() if len(parts) == 2 else group_str
            rest = parts[0].strip() if len(parts) == 2 else group_str
            
            match_date = pd.to_datetime(date_str, dayfirst=True, errors='coerce')
            if pd.notna(match_date):
                match_date = match_date.normalize()
                
            if ' v ' in rest:
                team_a, remainder = rest.split(' v ', 1)
                team_b = remainder.rsplit(', ', 1)[0] if ', ' in remainder else (remainder.rsplit(' - ', 1)[0] if ' - ' in remainder else remainder)
            else:
                team_a, team_b = rest, "Unknown"
            return team_a.strip(), team_b.strip(), match_date
        except: return None, None, None

    cup_match_dict = {}
    if f_cup and os.path.exists(f_cup):
        try:
            excel_file = pd.ExcelFile(f_cup)
            target_sheet = excel_file.sheet_names[0]
            
            for sheet in excel_file.sheet_names:
                if domain.lower().replace("'", "") in sheet.lower().replace("'", ""):
                    target_sheet = sheet
                    break
                    
            cup_df = pd.read_excel(f_cup, sheet_name=target_sheet, header=None)
            
            for _, row_data in cup_df.iterrows():
                match_str_raw = str(row_data[0]).strip()
                cup_name = str(row_data[1]).strip()
                
                if match_str_raw.lower() in ['match string', 'match group', 'match', 'nan']:
                    continue
                
                cleaned_match_str = doc_format_cricket_names(match_str_raw, domain)
                c_team_a, c_team_b, c_date = parse_match_group(cleaned_match_str)
                
                if c_team_a and c_team_b:
                    teams = sorted([str(c_team_a).lower(), str(c_team_b).lower()])
                    if pd.notna(c_date):
                        robust_key = f"{teams[0]}_{teams[1]}_{c_date.strftime('%Y-%m-%d')}"
                        cup_match_dict[robust_key] = cup_name
                    else:
                        robust_key_no_date = f"{teams[0]}_{teams[1]}"
                        cup_match_dict[robust_key_no_date] = cup_name
        except Exception:
            pass

    reg_name_col = 'Full Name' if 'Full Name' in registered_players.columns else registered_players.columns[0]
    registered_players[reg_name_col] = registered_players[reg_name_col].astype(str).str.replace('‡', '', regex=False).str.strip()

    if domain == "Men's":
        registered_players.loc[registered_players[reg_name_col].str.lower() == 'matthew humphreys', 'Individual Membership Primary Club'] = 'Instonians Cricket Club'
        registered_players[reg_name_col] = registered_players[reg_name_col].str.replace('Hoffmeyr', 'Hofmeyr', regex=False)
        exclusions = ['mark adair', 'ben calitz']
        pronoun = "he"
    elif domain == "Women's":
        cara_murray = registered_players[registered_players[reg_name_col].str.lower() == 'cara murray']
        if not cara_murray.empty:
            cara_waringstown = cara_murray.copy()
            cara_waringstown['Individual Membership Primary Club'] = 'Waringstown Cricket Club'
            registered_players = pd.concat([registered_players, cara_waringstown], ignore_index=True)
        exclusions = ['cara murray']
        pronoun = "she"

    registered_players['Date Registered'] = pd.to_datetime(registered_players['Date Registered'], dayfirst=True, errors='coerce').dt.normalize()

    starring_df = pd.DataFrame(columns=['Rank', 'Surname', 'Forename', 'XI_Level', 'Club', 'Full Name'])
    if f_starring and os.path.exists(f_starring):
        starring_excel = pd.read_excel(f_starring, sheet_name=None, header=None)
        starring_data_list = []
        for club_name, df in starring_excel.items():
            try:
                while len(df.columns) < 5: df[len(df.columns)] = None
                df = df.iloc[:, [0, 1, 4]]
                df.columns = ['Rank', 'Surname', 'Forename']
                df['XI_Level'] = df['Rank'].apply(lambda x: str(x).strip() if 'XI' in str(x) else None).ffill()
                df['Is_Numeric'] = pd.to_numeric(df['Rank'], errors='coerce').notna()
                df = df[df['Is_Numeric']].copy().dropna(subset=['Surname'])
                if not df.empty:
                    df['Club'] = str(club_name).strip()
                    df['Forename'] = df['Forename'].fillna('')
                    df['Full Name'] = (df['Forename'].astype(str).str.strip() + ' ' + df['Surname'].astype(str).str.strip()).str.replace('‡', '', regex=False).str.strip()
                    starring_data_list.append(df)
            except Exception: pass
        if starring_data_list: starring_df = pd.concat(starring_data_list, ignore_index=True)

    alias_map = build_alias_map(aliases, domain)
    player_club_map = build_player_club_map(registered_players, alias_map, domain) 
    
    # Apply row-based contextual name cleansing
    batting_stats['Cleaned Name'] = batting_stats.apply(lambda r: cleanse_name_contextual(r['Name'], r, alias_map), axis=1)
    bowling_stats['Cleaned Name'] = bowling_stats.apply(lambda r: cleanse_name_contextual(r['Bowler'], r, alias_map), axis=1)
    
    batting_stats['Group'] = batting_stats['Group'].apply(lambda x: doc_format_cricket_names(x, domain))
    bowling_stats['Group'] = bowling_stats['Group'].apply(lambda x: doc_format_cricket_names(x, domain))

    batters = batting_stats[['Group', 'Cleaned Name', 'Name', 'Is_Irish_Match']].rename(columns={'Cleaned Name': 'Player', 'Name': 'Scorecard Name'})
    bowlers = bowling_stats[['Group', 'Cleaned Name', 'Bowler', 'Is_Irish_Match']].rename(columns={'Cleaned Name': 'Player', 'Bowler': 'Scorecard Name'})
    all_appearances = pd.concat([batters, bowlers]).drop_duplicates(subset=['Group', 'Player'])
    all_appearances[['Team A', 'Team B', 'Match Date']] = all_appearances['Group'].apply(lambda x: pd.Series(parse_match_group(x)))
    all_appearances = all_appearances.sort_values(by=['Match Date'])

    league_dict, team_keys, _ = build_league_dict(league_structure)
    def determine_league(t_a, t_b):
        league_a, league_b = get_team_league(t_a, team_keys, league_dict, domain), get_team_league(t_b, team_keys, league_dict, domain)
        return league_a if league_a and league_b and league_a == league_b else "possible cup match"

    official_names = registered_players[reg_name_col].dropna().unique()
    deemed_registered, unregistered_audit, starring_violations = [], [], []
    all_matches_in_range, violation_matches = set(), set()
    first_unreg_match_date, player_match_cache = {}, {}

    for idx, row in all_appearances.iterrows():
        player, scorecard_name, match_date = row['Player'], row['Scorecard Name'], row['Match Date']
        if pd.isna(match_date): continue
            
        team_a, team_b = row['Team A'], row['Team B']
        
        match_league = None
        teams = sorted([str(team_a).lower(), str(team_b).lower()])
        robust_key_date = f"{teams[0]}_{teams[1]}_{match_date.strftime('%Y-%m-%d')}"
        robust_key_no_date = f"{teams[0]}_{teams[1]}"
        
        if robust_key_date in cup_match_dict:
            match_league = cup_match_dict[robust_key_date]
        elif robust_key_no_date in cup_match_dict:
            match_league = cup_match_dict[robust_key_no_date]
            
        if not match_league:
            match_league = determine_league(team_a, team_b)
            
        in_date_range = (start_date <= match_date <= end_date)
        
        if in_date_range: all_matches_in_range.add((team_a, team_b, match_league, match_date))

        if row.get('Is_Irish_Match', False):
            reg_club = player_club_map.get(player.lower())
            if not reg_club or str(reg_club).lower() == 'nan':
                continue
                
            mock_row = {'Cleaned Name': player, 'Group': row['Group']}
            played_for = determine_player_team_for_row(mock_row, player_club_map, domain)
            if not get_team_league(played_for, team_keys, league_dict, domain):
                continue

        if player not in player_match_cache:
            if domain == "Men's" and player.lower() == 'james shannon':
                if 'holywood' in str(team_a).lower() or 'holywood' in str(team_b).lower():
                    reg_record = registered_players[(registered_players[reg_name_col].str.lower() == 'james shannon') & (registered_players['Individual Membership Primary Club'].str.contains('Holywood', case=False, na=False))].copy()
                    if not reg_record.empty: reg_record['Date Registered'] = pd.to_datetime('2026-03-05', dayfirst=True).normalize()
                elif 'saintfield' in str(team_a).lower() or 'saintfield' in str(team_b).lower():
                    reg_record = registered_players[(registered_players[reg_name_col].str.lower() == 'james shannon') & (registered_players['Individual Membership Primary Club'].str.contains('Saintfield', case=False, na=False))]
                else: reg_record = pd.DataFrame()
                match_type, matched_name = "Exact (Contextual Override)", "James Shannon"
                
            elif '(' in player and player.strip().endswith(')'):
                base_name = player.split('(')[0].strip()
                club_hint = player.split('(')[-1].replace(')', '').strip()
                
                # 1. Search for the base name
                potential_matches = registered_players[registered_players[reg_name_col].str.strip().str.lower() == base_name.lower()]
                if potential_matches.empty:
                    best_match, score = process.extractOne(base_name, official_names, scorer=fuzz.token_sort_ratio)
                    if score >= 90:
                        potential_matches = registered_players[registered_players[reg_name_col] == best_match]
                
                # 2. Filter by the club inside the brackets
                if not potential_matches.empty:
                    reg_record = potential_matches[potential_matches['Individual Membership Primary Club'].astype(str).str.contains(club_hint, case=False, na=False)]
                    if not reg_record.empty:
                        match_type, matched_name = f"Duplicate Match ({club_hint})", reg_record.iloc[0][reg_name_col]
                    else:
                        reg_record = pd.DataFrame()
                        match_type, matched_name = "Failed", "NO MATCH FOUND"
                else:
                    reg_record = pd.DataFrame()
                    match_type, matched_name = "Failed", "NO MATCH FOUND"
                    
            else:
                reg_record = registered_players[registered_players[reg_name_col].str.strip().str.lower() == player.lower()]
                match_type, matched_name = "Exact", player
                if reg_record.empty:
                    best_match, score = process.extractOne(player, official_names, scorer=fuzz.token_sort_ratio)
                    if score >= 90:
                        reg_record = registered_players[registered_players[reg_name_col] == best_match]
                        match_type, matched_name = f"Fuzzy ({score}%)", best_match
            player_match_cache[player] = (reg_record, match_type, matched_name)
        else:
            reg_record, match_type, matched_name = player_match_cache[player]

        is_registered = False
        reg_date, reg_club = pd.NaT, "Unknown Club"
        if not reg_record.empty:
            reg_date = reg_record.iloc[0]['Date Registered']
            raw_club = reg_record.iloc[0].get('Individual Membership Primary Club', pd.NA)
            if pd.notna(raw_club) and str(raw_club).strip() != '': reg_club = str(raw_club).strip()
            if pd.notna(reg_date) and reg_date <= match_date: is_registered = True
                
        if not is_registered:
            status_text = 'Unregistered for this match' if not reg_record.empty else 'Unregistered / Missing completely'
            f_match_logic = match_type if not reg_record.empty else 'Failed'
            f_matched_name = matched_name if not reg_record.empty else 'NO MATCH FOUND'
            
            if player not in first_unreg_match_date:
                first_unreg_match_date[player] = match_date
                if in_date_range:
                    violation_matches.add((team_a, team_b, match_league, match_date))
                    unregistered_audit.append({
                        'Stats Name (Cleaned)': player, 'Original Scorecard Name': scorecard_name,
                        'Matched Registered Name': f_matched_name, 'Registered Club': reg_club,
                        'Match Date': match_date, 'Date Registered': reg_date, 'Status': status_text,
                        'Team A': team_a, 'Team B': team_b, 'Match League': match_league, 'Match Logic': f_match_logic
                    })
            else:
                if in_date_range:
                    deemed_registered.append({
                        'Stats Name (Cleaned)': player, 'Original Scorecard Name': scorecard_name,
                        'Matched Registered Name': f_matched_name, 'Registered Club': reg_club,
                        'Match Date': match_date, 'Deemed Registered Date': first_unreg_match_date[player],
                        'Date Registered': reg_date, 'Team A': team_a, 'Team B': team_b,
                        'Match League': match_league, 'Match Logic': f_match_logic
                    })

    valid_matches = all_appearances[(all_appearances['Match Date'] >= start_date) & (all_appearances['Match Date'] <= end_date)].copy()
    if not starring_df.empty and 'Full Name' in starring_df.columns:
        starring_df['Cleaned Name'] = starring_df['Full Name'].apply(lambda x: cleanse_name(x, alias_map))
        for idx, row in valid_matches.iterrows():
            player, scorecard_name, team_a, team_b = row['Player'], row['Scorecard Name'], str(row['Team A']), str(row['Team B'])
            if str(player).strip().lower() in exclusions: continue
            
            p_stars = starring_df[starring_df['Cleaned Name'].str.strip().str.lower() == player.lower()]
            if not p_stars.empty:
                s_rank, s_club = str(p_stars.iloc[0]['XI_Level']), str(p_stars.iloc[0]['Club'])
                played_team = team_a if s_club.lower() in team_a.lower() else (team_b if s_club.lower() in team_b.lower() else None)
                if played_team:
                    try:
                        s_rm = re.search(r'(\d+)', s_rank)
                        p_rm = re.search(r'\s([1-9])$', played_team)
                        
                        s_rank_int = int(s_rm.group(1)) if s_rm else 1
                        p_rank_int = int(p_rm.group(1)) if p_rm else 1
                        
                        if p_rank_int > s_rank_int:
                            starring_violations.append({
                                'Player (Cleaned)': player, 'Original Scorecard Name': scorecard_name,
                                'Starred For': f"{s_club} {s_rank}", 'Actually Played For': played_team,
                                'Team A': team_a, 'Team B': team_b, 'Match Date': row['Match Date'], 'Match Group': row['Group']
                            })
                    except: pass

    unregistered_audit.sort(key=lambda x: violation_sort_key(x, 'Stats Name (Cleaned)'))
    deemed_registered.sort(key=lambda x: violation_sort_key(x, 'Stats Name (Cleaned)'))
    starring_violations.sort(key=lambda x: violation_sort_key(x, 'Player (Cleaned)'))

    excel_io = io.BytesIO()
    with pd.ExcelWriter(excel_io, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as writer:
        if unregistered_audit: export_and_format_excel(pd.DataFrame(unregistered_audit), writer, "Unregistered Matches")
        else: pd.DataFrame(columns=["Status"]).to_excel(writer, index=False, sheet_name="Unregistered Matches")
        if deemed_registered: export_and_format_excel(pd.DataFrame(deemed_registered), writer, "Deemed Registered")
        else: pd.DataFrame(columns=["Status"]).to_excel(writer, index=False, sheet_name="Deemed Registered")
        if starring_violations: export_and_format_excel(pd.DataFrame(starring_violations), writer, "Starring Violations")
        else: pd.DataFrame(columns=["Status"]).to_excel(writer, index=False, sheet_name="Starring Violations")

    doc = Document()
    doc.styles['Normal'].font.name, doc.styles['Normal'].font.size = 'Calibri', Pt(11)
    start_str, end_str = get_ordinal_date(start_date, False), get_ordinal_date(end_date, True)
    
    p_intro = doc.add_paragraph()
    run_intro = p_intro.add_run(f"Here is the audit for the matches played between {start_str} and {end_str}.")
    run_intro.bold = True
    run_intro.font.size = Pt(13)
    run_intro.font.color.rgb = RGBColor(0, 0, 128)

    doc.add_paragraph("I have checked the player names against the official NCU registrations, applied fuzzy logic mapping for variances, and cross-referenced the Date Registered field to ensure players were officially registered on or before the date of their match.")
    
    p_noviol = doc.add_paragraph()
    p_noviol.paragraph_format.space_before = Pt(14)
    p_noviol.paragraph_format.space_after = Pt(4)
    run_noviol = p_noviol.add_run("--- Matches WITH Violations ---")
    run_noviol.bold = True
    run_noviol.font.size = Pt(13)
    run_noviol.font.color.rgb = RGBColor(0, 0, 128)

    if violation_matches:
        render_grouped_matches(doc, violation_matches, domain)
    else:
        doc.add_paragraph("No matches with violations recorded.")

    p_noviol = doc.add_paragraph()
    p_noviol.paragraph_format.space_before = Pt(14)  
    p_noviol.paragraph_format.space_after = Pt(4)    
    run_noviol = p_noviol.add_run("--- Matches WITHOUT Violations ---")
    run_noviol.bold = True
    run_noviol.font.size = Pt(13)
    run_noviol.font.color.rgb = RGBColor(0, 0, 128)

    matches_without_violations = all_matches_in_range - violation_matches
    if matches_without_violations:
        render_grouped_matches(doc, matches_without_violations, domain)
    else:
        doc.add_paragraph("No matches without violations recorded.")

    doc.add_page_break()
    p_unreg = doc.add_paragraph()
    run_unreg = p_unreg.add_run("--- Unregistered Players / Date Violations ---")
    run_unreg.bold = True
    run_unreg.font.size = Pt(13)
    run_unreg.font.color.rgb = RGBColor(0, 0, 128)

    doc.add_paragraph('These players took the field prior to the official "Date Registered" logged in the master file database.')
    for r in unregistered_audit:
        p_name, s_name = r['Stats Name (Cleaned)'], r['Original Scorecard Name']
        d_name = p_name if p_name.lower() == str(s_name).lower() else f"{p_name} (Played as: {s_name})"
        t_str = f" [Registered Club: {doc_formal_team_name(r.get('Registered Club', 'Unknown Club'))}] (Match: {doc_formal_team_name(r['Team A'])} v {doc_formal_team_name(r['Team B'])})"
        
        p_p = doc.add_paragraph(style='List Bullet')
        p_p.add_run(f"{d_name}").bold = True
        p_p.add_run(f"{t_str}")
        
        r_date = r['Date Registered']
        if pd.notna(r_date): d_text = f"Appeared in the match on {get_ordinal_date(r['Match Date'])}. The official database indicates a registration date of {get_ordinal_date(r_date)} ({(r_date - r['Match Date']).days} days late)."
        else: d_text = f"Appeared in the match on {get_ordinal_date(r['Match Date'])} under the scorecard name \"{s_name}\" (and verified via alias map). This official profile is entirely unregistered on the master registry."
        
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Pt(36)
        p.add_run("o  Violation Detail: ").bold = True
        p.add_run(d_text)

    doc.add_page_break()
    p_deemed = doc.add_paragraph()
    run_deemed = p_deemed.add_run("--- Deemed Registered (Played Previously in 2026) ---")
    run_deemed.bold = True
    run_deemed.font.size = Pt(13)
    run_deemed.font.color.rgb = RGBColor(0, 0, 128)

    doc.add_paragraph("These players were late or unverified on their match date, but are deemed registered for this specific game because they already played an active match earlier in the recorded 2026 season logs.")
    for r in deemed_registered:
        p_name, s_name = r['Stats Name (Cleaned)'], r['Original Scorecard Name']
        d_name = p_name if p_name.lower() == str(s_name).lower() else f"{p_name} (Played as: {s_name})"
        t_str = f" [Registered Club: {doc_formal_team_name(r.get('Registered Club', 'Unknown Club'))}] (Match: {doc_formal_team_name(r['Team A'])} v {doc_formal_team_name(r['Team B'])})"
        
        p_p = doc.add_paragraph(style='List Bullet')
        p_p.add_run(f"{d_name}").bold = True
        p_p.add_run(f"{t_str}")
        
        r_date = r['Date Registered']
        r_det = f"Registered late on {get_ordinal_date(r_date)}" if pd.notna(r_date) else "Unregistered profile"
        
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Pt(36)
        p.add_run("o  Deemed Status: ").bold = True
        p.add_run(f"Played on {get_ordinal_date(r['Match Date'])} ({r_det}). Deemed registered because {pronoun} previously played on {get_ordinal_date(r['Deemed Registered Date'])}.")

    doc.add_page_break()
    p_star = doc.add_paragraph()
    run_star = p_star.add_run("--- Starring Violations ---")
    run_star.bold = True
    run_star.font.size = Pt(13)
    run_star.font.color.rgb = RGBColor(0, 0, 128)

    doc.add_paragraph("These players played for a team at a lower level than their starred rank.")
    for r in starring_violations:
        p_name, s_name = r['Player (Cleaned)'], r['Original Scorecard Name']
        d_name = p_name if p_name.lower() == str(s_name).lower() else f"{p_name} (Played as: {s_name})"
        t_str = f" (Match: {doc_formal_team_name(r['Team A'])} v {doc_formal_team_name(r['Team B'])})"
        
        p_p = doc.add_paragraph(style='List Bullet')
        p_p.add_run(f"{d_name}").bold = True
        p_p.add_run(f"{t_str}")
        
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Pt(36)
        p.add_run("o  Violation Detail: ").bold = True
        p.add_run(f"Appeared in the match on {get_ordinal_date(r['Match Date'])} for {doc_formal_team_name(r['Actually Played For'])}, but is starred for {doc_formal_team_name(r['Starred For'])}.")

    doc_io = io.BytesIO()
    doc.save(doc_io)
    
    return excel_io, doc_io

# ==========================================
# MIDWEEK REGISTRATION ENGINE
# ==========================================
def run_midweek_registration_audit(start_date, end_date, f_reg, f_alias, f_starring, f_weekend_league, f_midweek_league, f_bat, f_bowl):
    registered_players = pd.read_excel(f_reg)
    aliases = pd.read_excel(f_alias)
    weekend_structure = pd.read_excel(f_weekend_league)
    midweek_structure = pd.read_excel(f_midweek_league)
    batting_stats = pd.read_excel(f_bat)
    bowling_stats = pd.read_excel(f_bowl)

    reg_name_col = 'Full Name' if 'Full Name' in registered_players.columns else registered_players.columns[0]
    registered_players[reg_name_col] = registered_players[reg_name_col].astype(str).str.replace('‡', '', regex=False).str.strip()

    registered_players.loc[registered_players[reg_name_col].str.lower() == 'matthew humphreys', 'Individual Membership Primary Club'] = 'Instonians Cricket Club'
    registered_players[reg_name_col] = registered_players[reg_name_col].str.replace('Hoffmeyr', 'Hofmeyr', regex=False)
    registered_players['Date Registered'] = pd.to_datetime(registered_players['Date Registered'], dayfirst=True, errors='coerce').dt.normalize()

    starring_df = pd.DataFrame(columns=['Rank', 'Surname', 'Forename', 'XI_Level', 'Club', 'Full Name'])
    if f_starring and os.path.exists(f_starring):
        starring_excel = pd.read_excel(f_starring, sheet_name=None, header=None)
        starring_data_list = []
        for club_name, df in starring_excel.items():
            try:
                while len(df.columns) < 5: df[len(df.columns)] = None
                df = df.iloc[:, [0, 1, 4]]
                df.columns = ['Rank', 'Surname', 'Forename']
                df['XI_Level'] = df['Rank'].apply(lambda x: str(x).strip() if 'XI' in str(x) else None).ffill()
                df['Is_Numeric'] = pd.to_numeric(df['Rank'], errors='coerce').notna()
                df = df[df['Is_Numeric']].copy().dropna(subset=['Surname'])
                if not df.empty:
                    df['Club'] = str(club_name).strip()
                    df['Forename'] = df['Forename'].fillna('')
                    df['Full Name'] = (df['Forename'].astype(str).str.strip() + ' ' + df['Surname'].astype(str).str.strip()).str.replace('‡', '', regex=False).str.strip()
                    starring_data_list.append(df)
            except Exception: pass
        if starring_data_list: starring_df = pd.concat(starring_data_list, ignore_index=True)

    alias_map = build_alias_map(aliases, "Midweek")
    
    # Apply row-based contextual name cleansing
    batting_stats['Cleaned Name'] = batting_stats.apply(lambda r: cleanse_name_contextual(r['Name'], r, alias_map), axis=1)
    bowling_stats['Cleaned Name'] = bowling_stats.apply(lambda r: cleanse_name_contextual(r['Bowler'], r, alias_map), axis=1)
    
    batting_stats['Group'] = batting_stats['Group'].apply(lambda x: doc_format_cricket_names(x, "Midweek"))
    bowling_stats['Group'] = bowling_stats['Group'].apply(lambda x: doc_format_cricket_names(x, "Midweek"))

    def parse_match_group(group_str):
        try:
            group_str = str(group_str).strip()
            parts = group_str.rsplit(' - ', 1)
            date_str = parts[1].strip() if len(parts) == 2 else group_str
            rest = parts[0].strip() if len(parts) == 2 else group_str
            match_date = pd.to_datetime(date_str, dayfirst=True, errors='coerce').normalize()
            if ' v ' in rest:
                team_a, remainder = rest.split(' v ', 1)
                team_b = remainder.rsplit(', ', 1)[0] if ', ' in remainder else (remainder.rsplit(' - ', 1)[0] if ' - ' in remainder else remainder)
            else:
                team_a, team_b = rest, "Unknown"
            return team_a.strip(), team_b.strip(), match_date
        except: return None, None, None

    batters = batting_stats[['Group', 'Cleaned Name', 'Name']].rename(columns={'Cleaned Name': 'Player', 'Name': 'Scorecard Name'})
    bowlers = bowling_stats[['Group', 'Cleaned Name', 'Bowler']].rename(columns={'Cleaned Name': 'Player', 'Bowler': 'Scorecard Name'})
    all_appearances = pd.concat([batters, bowlers]).drop_duplicates(subset=['Group', 'Player'])
    all_appearances[['Team A', 'Team B', 'Match Date']] = all_appearances['Group'].apply(lambda x: pd.Series(parse_match_group(x)))
    all_appearances = all_appearances.sort_values(by=['Match Date'])

    mw_league_dict, mw_team_keys, _ = build_league_dict(midweek_structure)
    wknd_league_dict, wknd_team_keys, _ = build_league_dict(weekend_structure)

    def determine_midweek_league(t_a, t_b):
        league_a, league_b = get_team_league(t_a, mw_team_keys, mw_league_dict, "Midweek"), get_team_league(t_b, mw_team_keys, mw_league_dict, "Midweek")
        return league_a if league_a and league_b and league_a == league_b else "Midweek Cup/Fixture"

    official_names = registered_players[reg_name_col].dropna().unique()
    deemed_registered, unregistered_audit, starring_violations = [], [], []
    all_matches_in_range, violation_matches = set(), set()
    first_unreg_match_date, player_match_cache = {}, {}

    for idx, row in all_appearances.iterrows():
        player, scorecard_name, match_date = row['Player'], row['Scorecard Name'], row['Match Date']
        if pd.isna(match_date): continue
            
        team_a, team_b = row['Team A'], row['Team B']
        match_league = determine_midweek_league(team_a, team_b)
        in_date_range = (start_date <= match_date <= end_date)
        
        if in_date_range: all_matches_in_range.add((team_a, team_b, match_league, match_date))

        if player not in player_match_cache:
            if player.lower() == 'james shannon':
                if 'holywood' in str(team_a).lower() or 'holywood' in str(team_b).lower():
                    reg_record = registered_players[(registered_players[reg_name_col].str.lower() == 'james shannon') & (registered_players['Individual Membership Primary Club'].str.contains('Holywood', case=False, na=False))].copy()
                    if not reg_record.empty: reg_record['Date Registered'] = pd.to_datetime('2026-03-05', dayfirst=True).normalize()
                elif 'saintfield' in str(team_a).lower() or 'saintfield' in str(team_b).lower():
                    reg_record = registered_players[(registered_players[reg_name_col].str.lower() == 'james shannon') & (registered_players['Individual Membership Primary Club'].str.contains('Saintfield', case=False, na=False))]
                else: reg_record = pd.DataFrame()
                match_type, matched_name = "Exact (Contextual Override)", "James Shannon"
                
            elif '(' in player and player.strip().endswith(')'):
                base_name = player.split('(')[0].strip()
                club_hint = player.split('(')[-1].replace(')', '').strip()
                
                potential_matches = registered_players[registered_players[reg_name_col].str.strip().str.lower() == base_name.lower()]
                if potential_matches.empty:
                    best_match, score = process.extractOne(base_name, official_names, scorer=fuzz.token_sort_ratio)
                    if score >= 90:
                        potential_matches = registered_players[registered_players[reg_name_col] == best_match]
                
                if not potential_matches.empty:
                    reg_record = potential_matches[potential_matches['Individual Membership Primary Club'].astype(str).str.contains(club_hint, case=False, na=False)]
                    if not reg_record.empty:
                        match_type, matched_name = f"Duplicate Match ({club_hint})", reg_record.iloc[0][reg_name_col]
                    else:
                        reg_record = pd.DataFrame()
                        match_type, matched_name = "Failed", "NO MATCH FOUND"
                else:
                    reg_record = pd.DataFrame()
                    match_type, matched_name = "Failed", "NO MATCH FOUND"
                    
            else:
                reg_record = registered_players[registered_players[reg_name_col].str.strip().str.lower() == player.lower()]
                match_type, matched_name = "Exact", player
                if reg_record.empty:
                    best_match, score = process.extractOne(player, official_names, scorer=fuzz.token_sort_ratio)
                    if score >= 90:
                        reg_record = registered_players[registered_players[reg_name_col] == best_match]
                        match_type, matched_name = f"Fuzzy ({score}%)", best_match
            player_match_cache[player] = (reg_record, match_type, matched_name)
        else:
            reg_record, match_type, matched_name = player_match_cache[player]

        is_registered = False
        reg_date, reg_club = pd.NaT, "Unknown Club"
        if not reg_record.empty:
            reg_date = reg_record.iloc[0]['Date Registered']
            raw_club = reg_record.iloc[0].get('Individual Membership Primary Club', pd.NA)
            if pd.notna(raw_club) and str(raw_club).strip() != '': reg_club = str(raw_club).strip()
            if pd.notna(reg_date) and reg_date <= match_date: is_registered = True
                
        if not is_registered:
            status_text = 'Unregistered for this match' if not reg_record.empty else 'Unregistered / Missing completely'
            f_match_logic = match_type if not reg_record.empty else 'Failed'
            f_matched_name = matched_name if not reg_record.empty else 'NO MATCH FOUND'
            
            if player not in first_unreg_match_date:
                first_unreg_match_date[player] = match_date
                if in_date_range:
                    violation_matches.add((team_a, team_b, match_league, match_date))
                    unregistered_audit.append({
                        'Stats Name (Cleaned)': player, 'Original Scorecard Name': scorecard_name,
                        'Matched Registered Name': f_matched_name, 'Registered Club': reg_club,
                        'Match Date': match_date, 'Date Registered': reg_date, 'Status': status_text,
                        'Team A': team_a, 'Team B': team_b, 'Match League': match_league, 'Match Logic': f_match_logic
                    })
            else:
                if in_date_range:
                    deemed_registered.append({
                        'Stats Name (Cleaned)': player, 'Original Scorecard Name': scorecard_name,
                        'Matched Registered Name': f_matched_name, 'Registered Club': reg_club,
                        'Match Date': match_date, 'Deemed Registered Date': first_unreg_match_date[player],
                        'Date Registered': reg_date, 'Team A': team_a, 'Team B': team_b,
                        'Match League': match_league, 'Match Logic': f_match_logic
                    })

    valid_matches = all_appearances[(all_appearances['Match Date'] >= start_date) & (all_appearances['Match Date'] <= end_date)].copy()
    international_exclusions = ['mark adair', 'ben calitz']

    if not starring_df.empty and 'Full Name' in starring_df.columns:
        starring_df['Cleaned Name'] = starring_df['Full Name'].apply(lambda x: cleanse_name(x, alias_map))
        
        for idx, row in valid_matches.iterrows():
            player, scorecard_name = row['Player'], row['Scorecard Name']
            team_a, team_b = str(row['Team A']), str(row['Team B'])
            if str(player).strip().lower() in international_exclusions: continue
                
            player_stars = starring_df[starring_df['Cleaned Name'].str.strip().str.lower() == player.lower()]
            if not player_stars.empty:
                starred_level, starred_club = str(player_stars.iloc[0]['XI_Level']), str(player_stars.iloc[0]['Club'])
                full_weekend_team = f"{starred_club} {starred_level}"
                weekend_division = get_team_league(full_weekend_team, wknd_team_keys, wknd_league_dict, "Men's")
                
                if weekend_division:
                    div_lower = str(weekend_division).lower()
                    is_illegal = False
                    
                    if 'premier' in div_lower or 'senior' in div_lower: is_illegal = True
                    elif 'junior' in div_lower:
                        match_num = re.search(r'junior(?: league)? (\d+)', div_lower)
                        if match_num and int(match_num.group(1)) <= 3: is_illegal = True
                            
                    if is_illegal:
                        violation_matches.add((team_a, team_b, determine_midweek_league(team_a, team_b), row['Match Date']))
                        mw_team = team_a if starred_club.lower() in team_a.lower() else (team_b if starred_club.lower() in team_b.lower() else team_a)
                        starring_violations.append({
                            'Player (Cleaned)': player, 'Original Scorecard Name': scorecard_name,
                            'Starred Rank': full_weekend_team, 'Weekend Division': weekend_division,
                            'Midweek Team': mw_team, 'Team A': team_a, 'Team B': team_b,
                            'Match Date': row['Match Date'], 'Match Group': row['Group']
                        })

    unregistered_audit.sort(key=lambda x: violation_sort_key(x, 'Stats Name (Cleaned)'))
    deemed_registered.sort(key=lambda x: violation_sort_key(x, 'Stats Name (Cleaned)'))
    starring_violations.sort(key=lambda x: violation_sort_key(x, 'Player (Cleaned)'))

    excel_io = io.BytesIO()
    with pd.ExcelWriter(excel_io, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as writer:
        if unregistered_audit: export_and_format_excel(pd.DataFrame(unregistered_audit), writer, "Unregistered Matches")
        else: pd.DataFrame(columns=["Status"]).to_excel(writer, index=False, sheet_name="Unregistered Matches")
        if deemed_registered: export_and_format_excel(pd.DataFrame(deemed_registered), writer, "Deemed Registered")
        else: pd.DataFrame(columns=["Status"]).to_excel(writer, index=False, sheet_name="Deemed Registered")
        if starring_violations: export_and_format_excel(pd.DataFrame(starring_violations), writer, "Starring Violations")
        else: pd.DataFrame(columns=["Status"]).to_excel(writer, index=False, sheet_name="Starring Violations")

    doc = Document()
    doc.styles['Normal'].font.name, doc.styles['Normal'].font.size = 'Calibri', Pt(11)
    start_str, end_str = get_ordinal_date(start_date, False), get_ordinal_date(end_date, True)
    
    p_intro = doc.add_paragraph()
    run_intro = p_intro.add_run(f"Here is the Midweek League audit for the matches played between {start_str} and {end_str}.")
    run_intro.bold = True
    run_intro.font.size = Pt(13)
    run_intro.font.color.rgb = RGBColor(0, 0, 128)

    doc.add_paragraph("This summary runs automated data integrity audits across competitor profiles, verifying registration windows and cross-referencing weekend tiering to flag ineligible players starred above Junior League 3.")
    
    p_viol = doc.add_paragraph()
    run_viol = p_viol.add_run("--- Matches WITH Violations ---")
    run_viol.bold = True
    run_viol.font.size = Pt(13)
    run_viol.font.color.rgb = RGBColor(0, 0, 128)

    if violation_matches:
        render_grouped_matches(doc, violation_matches, "Midweek")
    else:
        doc.add_paragraph("No matches with violations recorded.")

    p_noviol = doc.add_paragraph()
    p_noviol.paragraph_format.space_before = Pt(14)  
    p_noviol.paragraph_format.space_after = Pt(4)    
    run_noviol = p_noviol.add_run("--- Matches WITHOUT Violations ---")
    run_noviol.bold = True
    run_noviol.font.size = Pt(13)
    run_noviol.font.color.rgb = RGBColor(0, 0, 128)

    matches_without_violations = all_matches_in_range - violation_matches
    if matches_without_violations:
        render_grouped_matches(doc, matches_without_violations, "Midweek")
    else:
        doc.add_paragraph("No matches without violations recorded.")

    doc.add_page_break()
    p_unreg = doc.add_paragraph()
    run_unreg = p_unreg.add_run("--- Unregistered Players / Date Violations ---")
    run_unreg.bold = True
    run_unreg.font.size = Pt(13)
    run_unreg.font.color.rgb = RGBColor(0, 0, 128)

    for r in unregistered_audit:
        p_name, s_name = r['Stats Name (Cleaned)'], r['Original Scorecard Name']
        d_name = p_name if p_name.lower() == str(s_name).lower() else f"{p_name} (Played as: {s_name})"
        t_str = f" [Registered Club: {doc_formal_team_name(r.get('Registered Club', 'Unknown Club'))}] (Match: {doc_formal_team_name(r['Team A'])} v {doc_formal_team_name(r['Team B'])})"
        
        p_p = doc.add_paragraph(style='List Bullet')
        p_p.add_run(f"{d_name}").bold = True
        p_p.add_run(f"{t_str}")
        
        r_date = r['Date Registered']
        if pd.notna(r_date): d_text = f"Appeared in the match on {get_ordinal_date(r['Match Date'])}. The official database indicates a registration date of {get_ordinal_date(r_date)} ({(r_date - r['Match Date']).days} days late)."
        else: d_text = f"Appeared in the match on {get_ordinal_date(r['Match Date'])} under the scorecard name \"{s_name}\" (and verified via alias map). This official profile is entirely unregistered on the master registry."
        
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Pt(36)
        p.add_run("o  Violation Detail: ").bold = True
        p.add_run(d_text)

    doc.add_page_break()
    p_deemed = doc.add_paragraph()
    run_deemed = p_deemed.add_run("--- Deemed Registered (Played Previously in 2026) ---")
    run_deemed.bold = True
    run_deemed.font.size = Pt(13)
    run_deemed.font.color.rgb = RGBColor(0, 0, 128)

    for r in deemed_registered:
        p_name, s_name = r['Stats Name (Cleaned)'], r['Original Scorecard Name']
        d_name = p_name if p_name.lower() == str(s_name).lower() else f"{p_name} (Played as: {s_name})"
        t_str = f" [Registered Club: {doc_formal_team_name(r.get('Registered Club', 'Unknown Club'))}] (Match: {doc_formal_team_name(r['Team A'])} v {doc_formal_team_name(r['Team B'])})"
        
        p_p = doc.add_paragraph(style='List Bullet')
        p_p.add_run(f"{d_name}").bold = True
        p_p.add_run(f"{t_str}")
        
        r_date = r['Date Registered']
        r_det = f"Registered late on {get_ordinal_date(r_date)}" if pd.notna(r_date) else "Unregistered profile"
        
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Pt(36)
        p.add_run("o  Deemed Status: ").bold = True
        p.add_run(f"Played on {get_ordinal_date(r['Match Date'])} ({r_det}). Deemed registered because he previously played an active fixture on {get_ordinal_date(r['Deemed Registered Date'])}.")

    doc.add_page_break()
    p_star = doc.add_paragraph()
    run_star = p_star.add_run("--- Starring Ceiling Violations ---")
    run_star.bold = True
    run_star.font.size = Pt(13)
    run_star.font.color.rgb = RGBColor(0, 0, 128)

    doc.add_paragraph("The following players are barred from participating in the Midweek League because they hold weekend starring rankings of Junior League 3 or above (Premier, Senior, or Junior 1–3).")

    for r in starring_violations:
        p_name, s_name = r['Player (Cleaned)'], r['Original Scorecard Name']
        d_name = p_name if p_name.lower() == str(s_name).lower() else f"{p_name} (Played as: {s_name})"
        t_str = f" (Match: {doc_formal_team_name(r['Team A'])} v {doc_formal_team_name(r['Team B'])})"
        
        p_p = doc.add_paragraph(style='List Bullet')
        p_p.add_run(f"{d_name}").bold = True
        p_p.add_run(f"{t_str}")
        
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Pt(36)
        p.add_run("o  Violation Detail: ").bold = True
        p.add_run(f"Represented {doc_formal_team_name(r['Midweek Team'])} on {get_ordinal_date(r['Match Date'])}. This player is completely ineligible for Midweek cricket as he is officially starred for weekend squad '{r['Starred Rank']}', which plays in weekend tier '{r['Weekend Division']}' (Junior League 3 or above tiers are ineligible).")

    doc_io = io.BytesIO()
    doc.save(doc_io)
    
    return excel_io, doc_io

# ==========================================
# STARRING & INACTIVITY REPORT FUNCTIONS 
# ==========================================
def report_clean_spaces(text):
    if pd.isna(text) or str(text).strip().lower() == 'nan': return ""
    return re.sub(r'\s+', ' ', str(text)).strip()

def report_clean_team_name(team):
    if pd.isna(team): return team
    cleaned = report_clean_spaces(team)
    if "Holywood" in cleaned and "1881" not in cleaned:
        cleaned = cleaned.replace("Holywood", "Holywood 1881")
    return cleaned

def report_clean_score_display(text):
    if pd.isna(text): return text
    return re.sub(r'(\d+)/(\d+)', r'\1-\2', str(text))

def report_parse_match_date(group_str):
    months = {'April': 4, 'May': 5, 'June': 6, 'July': 7, 'August': 8, 'September': 9}
    match = re.search(r'(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)\s+(\d{4})', str(group_str))
    if match:
        day, month_str, year = int(match.group(1)), match.group(2), int(match.group(3))
        if month_str in months: return datetime(year, months[month_str], day)
    return None

def report_extract_competition(group_str):
    """Extracts competition string natively from NV Play exported Group strings."""
    if pd.isna(group_str) or str(group_str).strip().lower() == 'nan': 
        return "Unknown Competition"
    parts = str(group_str).split(' - ')
    if len(parts) >= 2:
        return parts[0].strip()
    return "Unknown Competition"

def report_get_team_dates_flexible(starred_team, team_match_dates_dict):
    def norm(t):
        t = str(t).lower().replace('1881', '')
        t = t.replace('1st', '1').replace('firsts', '1').replace('first', '1')
        t = t.replace('2nd', '2').replace('seconds', '2').replace('second', '2')
        t = t.replace('3rd', '3').replace('thirds', '3').replace('third', '3')
        t = t.replace('4th', '4').replace('fourths', '4').replace('fourth', '4')
        t = t.replace('5th', '5').replace('fifths', '5').replace('fifth', '5')
        t = re.sub(r'\bxi\b', '', t)
        t = re.sub(r'\bteam\b', '', t)
        return re.sub(r'[^a-z0-9]', '', t)
    
    if starred_team in team_match_dates_dict: return team_match_dates_dict[starred_team]
        
    n_starred = norm(starred_team)
    for k, dates in team_match_dates_dict.items():
        if norm(k) == n_starred: return dates
            
    st_nums = re.findall(r'\d+', n_starred)
    if st_nums:
        for k, dates in team_match_dates_dict.items():
            k_nums = re.findall(r'\d+', norm(k))
            if k_nums and st_nums[0] == k_nums[0]: return dates
                
    all_dates = []
    for dates in team_match_dates_dict.values(): all_dates.extend(dates)
    return list(set(all_dates))

def report_build_club_matches_df(club_name, all_app, override_map, comp_map):
    if all_app.empty: return pd.DataFrame(), {}
    matches = all_app['Group'].dropna().unique()
    club_name_clean = report_clean_spaces(club_name).lower()
    search_term = override_map.get(club_name_clean, club_name_clean)
    
    club_matches = [m for m in matches if search_term in report_clean_spaces(m).lower()]
    
    def extract_and_clean_target_team(m):
        match_teams = re.split(r'\s+v\s+', str(m).split(',')[0])
        for t in match_teams:
            if search_term in report_clean_spaces(t).lower(): return report_clean_team_name(t)
        return "Unknown"
        
    team_dict = {}
    for m in club_matches:
        team = extract_and_clean_target_team(m)
        if team not in team_dict: team_dict[team] = []
        team_dict[team].append(m)
        
    formatted_rows = []
    team_match_dates = {}
    
    for team in sorted(team_dict.keys()):
        formatted_rows.append({"Club Match History": f"Team: {team}", "Competition": ""})
        sorted_team_matches = sorted(team_dict[team], key=lambda x: report_parse_match_date(x) or datetime.min)
        dates_list = []
        for m in sorted_team_matches:
            comp = comp_map.get(m, "Unknown Competition")
            formatted_rows.append({"Club Match History": f"  - {report_clean_score_display(report_clean_spaces(m))}", "Competition": comp})
            parsed_d = report_parse_match_date(m)
            if parsed_d: dates_list.append(parsed_d)
        
        team_match_dates[team.lower().strip()] = dates_list
        formatted_rows.append({"Club Match History": "", "Competition": ""})
        
    return pd.DataFrame(formatted_rows), team_match_dates

def report_build_player_stats_dfs(player_list, player_team_map, all_app, get_official_func, comp_map):
    raw_summary, log = [], []
    for p in player_list:
        official_p = get_official_func(p)
        official_p_lower = official_p.lower()
        p_apps = all_app[all_app['Official_Player_Lower'] == official_p_lower] if not all_app.empty else pd.DataFrame()
        matches_played = len(p_apps)
        team = report_clean_team_name(player_team_map.get(report_clean_spaces(p).lower(), "Unassigned"))
        
        if matches_played > 0:
            sorted_apps = p_apps.sort_values('Parsed_Date')
            last_date = sorted_apps['Parsed_Date'].max().strftime('%d %B %Y')
            last_match_grp = report_clean_spaces(sorted_apps.iloc[-1]['Group'])
            last_comp = comp_map.get(last_match_grp, "Unknown Competition")
            
            raw_summary.append({
                "Starred Team": team, "Input Name": report_clean_spaces(p), "Player (Official)": official_p, 
                "Matches Played": matches_played, "Last Played Date": last_date, 
                "Last Match Details": last_match_grp, "Competition": last_comp, "Highlight Reason": ""
            })
            log.append({"Player Log": f"Player: {official_p} ({report_clean_spaces(p)}) | Starred Team: {team}", "Competition": "", "Highlight Reason": ""})
            for _, row in sorted_apps.iterrows(): 
                grp = report_clean_spaces(row['Group'])
                comp = comp_map.get(grp, "Unknown Competition")
                log.append({"Player Log": f"  - {grp}", "Competition": comp, "Highlight Reason": ""})
        else:
            raw_summary.append({
                "Starred Team": team, "Input Name": report_clean_spaces(p), "Player (Official)": official_p, 
                "Matches Played": 0, "Last Played Date": "N/A", "Last Match Details": "N/A", "Competition": "N/A", "Highlight Reason": ""
            })
            log.append({"Player Log": f"Player: {official_p} ({report_clean_spaces(p)}) | Starred Team: {team}\n  - No match records found.\n", "Competition": "", "Highlight Reason": ""})
        log.append({"Player Log": "", "Competition": "", "Highlight Reason": ""})
        
    if not raw_summary: return pd.DataFrame(), pd.DataFrame(log)
        
    df_raw = pd.DataFrame(raw_summary).sort_values(by=["Starred Team", "Player (Official)"])
    formatted_summary = []
    current_team = None
    
    for _, row in df_raw.iterrows():
        if row["Starred Team"] != current_team:
            if current_team is not None:
                formatted_summary.append({
                    "Input Name": "", "Player (Official)": "", "Matches Played": "", 
                    "Last Played Date": "", "Last Match Details": "", "Competition": "", "Highlight Reason": ""
                })
            current_team = row["Starred Team"]
            formatted_summary.append({
                "Input Name": f"Team Starred For: {current_team}", "Player (Official)": "", "Matches Played": "", 
                "Last Played Date": "", "Last Match Details": "", "Competition": "", "Highlight Reason": ""
            })
            
        formatted_summary.append({
            "Input Name": row["Input Name"], "Player (Official)": row["Player (Official)"], "Matches Played": row["Matches Played"],
            "Last Played Date": row["Last Played Date"], "Last Match Details": row["Last Match Details"], "Competition": row["Competition"], "Highlight Reason": row["Highlight Reason"]
        })
        
    return pd.DataFrame(formatted_summary), pd.DataFrame(log)

def report_autofit_columns(ws):
    for col in ws.columns:
        max_length = 0
        column_letter = col[0].column_letter 
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except Exception: pass
        ws.column_dimensions[column_letter].width = max_length + 2

def generate_starring_inactivity_reports(domain, f_reg, f_alias, f_starring, f_bat, f_bowl, f_irish_bat=None, f_irish_bowl=None):
    df_reg = pd.read_excel(f_reg)
    df_alias = pd.read_excel(f_alias)
    df_bat = pd.read_excel(f_bat)
    df_bowl = pd.read_excel(f_bowl)
    
    # Conditionally append Irish batting/bowling stats if the files are provided
    df_bat['Is_Irish_Match'] = False
    df_bowl['Is_Irish_Match'] = False
    
    if f_irish_bat and os.path.exists(f_irish_bat):
        irish_bat = pd.read_excel(f_irish_bat)
        irish_bat['Is_Irish_Match'] = True
        df_bat = pd.concat([df_bat, irish_bat], ignore_index=True)
        
    if f_irish_bowl and os.path.exists(f_irish_bowl):
        irish_bowl = pd.read_excel(f_irish_bowl)
        irish_bowl['Is_Irish_Match'] = True
        df_bowl = pd.concat([df_bowl, irish_bowl], ignore_index=True)
    
    international_players = ["cara murray"] if domain == "Women's" else ["mark adair", "paul stirling"]
    override_map = {"holywood 1881": "holywood"}

    df_alias = df_alias.drop_duplicates(subset=['Input Name (Scorecard/Stats)'], keep='last')
    alias_map = dict(zip(
        df_alias['Input Name (Scorecard/Stats)'].apply(lambda x: report_clean_spaces(x).lower() if pd.notna(x) else x), 
        df_alias['Official Registered Name'].apply(lambda x: report_clean_spaces(x) if pd.notna(x) else x)
    ))
    if domain == "Men's": alias_map['tom mayes'] = 'Thomas Mayes' 
        
    registered_players_map = {}
    cols_lower = [str(c).lower() for c in df_reg.columns]
    if 'forename' in cols_lower and 'surname' in cols_lower:
        f_col = df_reg.columns[cols_lower.index('forename')]
        s_col = df_reg.columns[cols_lower.index('surname')]
        combined = df_reg[f_col].astype(str) + " " + df_reg[s_col].astype(str)
        registered_players_map = {report_clean_spaces(x).lower(): report_clean_spaces(x) for x in combined if 'nan' not in str(x).lower()}
    else:
        name_col = next((c for c in ['Name', 'Player', 'Player Name', 'Full Name', 'Registered Name'] if c in df_reg.columns), df_reg.columns[0])
        registered_players_map = {report_clean_spaces(str(x)).lower(): report_clean_spaces(str(x)) for x in df_reg[name_col].dropna()}
        
    def is_player_registered(name):
        cleaned = report_clean_spaces(name)
        mapped = alias_map.get(cleaned.lower(), cleaned)
        return mapped.lower() in registered_players_map

    def get_official_name_contextual(name, row):
        cleaned = report_clean_spaces(name)
        cleaned_lower = cleaned.lower()
        group_lower = str(row.get('Group', '')).lower()
        row_team = str(row.get('Team', '')).lower()
        
        if cleaned_lower == 'callum weir':
            if 'derriaghy' in group_lower or 'derriaghy' in row_team:
                return 'John Weir'
            elif 'cliftonville' in group_lower or 'cliftonville academy' in group_lower or 'cliftonville' in row_team:
                return 'Callum Weir'
                
        mapped = alias_map.get(cleaned_lower, cleaned)
        if mapped.lower() in registered_players_map: return registered_players_map[mapped.lower()]
        return mapped

    def get_official_name(name):
        cleaned = report_clean_spaces(name)
        mapped = alias_map.get(cleaned.lower(), cleaned)
        if mapped.lower() in registered_players_map: return registered_players_map[mapped.lower()]
        return mapped

    bat_cols = ['Name', 'Group', 'Is_Irish_Match']
    if 'Team' in df_bat.columns: bat_cols.append('Team')
    bowl_cols = ['Bowler', 'Group', 'Is_Irish_Match']
    if 'Team' in df_bowl.columns: bowl_cols.append('Team')

    all_app = pd.concat([
        df_bat[bat_cols].rename(columns={'Name': 'P'}) if not df_bat.empty else pd.DataFrame(columns=['P', 'Group', 'Is_Irish_Match']),
        df_bowl[bowl_cols].rename(columns={'Bowler': 'P'}) if not df_bowl.empty else pd.DataFrame(columns=['P', 'Group', 'Is_Irish_Match'])
    ], ignore_index=True)
    
    if not all_app.empty:
        all_app['Group'] = all_app['Group'].apply(report_clean_score_display) 
        all_app['Official_Player'] = all_app.apply(lambda r: get_official_name_contextual(r['P'], r), axis=1)
        all_app['Official_Player_Lower'] = all_app['Official_Player'].str.lower()
        all_app['Parsed_Date'] = all_app['Group'].apply(report_parse_match_date)
        all_app.drop_duplicates(subset=['Official_Player', 'Group'], inplace=True)

    f_league = DEFAULT_FILES[domain]["league"]
    league_dict, team_keys = {}, []
    if os.path.exists(f_league):
        league_structure = pd.read_excel(f_league)
        league_dict, team_keys, _ = build_league_dict(league_structure)
        
    f_cup = "NCU_Cup_Fixtures.xlsx"
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
                cleaned_match_str = doc_format_cricket_names(match_str_raw, domain)
                c_team_a, c_team_b, c_date = local_parse(cleaned_match_str)
                if c_team_a and c_team_b:
                    teams = sorted([str(c_team_a).lower(), str(c_team_b).lower()])
                    if pd.notna(c_date):
                        robust_key = f"{teams[0]}_{teams[1]}_{c_date.strftime('%Y-%m-%d')}"
                        cup_match_dict[robust_key] = cup_name
                    else:
                        robust_key_no_date = f"{teams[0]}_{teams[1]}"
                        cup_match_dict[robust_key_no_date] = cup_name
        except Exception: pass

    def apply_competition(grp_str, is_irish):
        try:
            grp_str_clean = doc_format_cricket_names(str(grp_str), domain)
            parts = grp_str_clean.rsplit(' - ', 1)
            date_str = parts[1].strip() if len(parts) == 2 else grp_str_clean
            rest = parts[0].strip() if len(parts) == 2 else grp_str_clean
            
            match_date = pd.to_datetime(date_str, dayfirst=True, errors='coerce')
            if pd.notna(match_date): match_date = match_date.normalize()
            
            if ' v ' in rest:
                team_a, remainder = rest.split(' v ', 1)
                team_b = remainder.rsplit(', ', 1)[0] if ', ' in remainder else (remainder.rsplit(' - ', 1)[0] if ' - ' in remainder else remainder)
            else:
                team_a, team_b = rest, "Unknown"
                
            t_a, t_b = team_a.strip(), team_b.strip()
            
            if t_a == "Unknown" or t_b == "Unknown":
                return "Irish Competition" if is_irish else "Unknown Competition"
                
            teams = sorted([str(t_a).lower(), str(t_b).lower()])
            if pd.notna(match_date):
                robust_key_date = f"{teams[0]}_{teams[1]}_{match_date.strftime('%Y-%m-%d')}"
                if robust_key_date in cup_match_dict: return cup_match_dict[robust_key_date]
            robust_key_no_date = f"{teams[0]}_{teams[1]}"
            if robust_key_no_date in cup_match_dict: return cup_match_dict[robust_key_no_date]
            
            if is_irish: return "Irish Competition"
            
            league_a = get_team_league(t_a, team_keys, league_dict, domain)
            league_b = get_team_league(t_b, team_keys, league_dict, domain)
            if league_a and league_b and league_a == league_b:
                league_str = str(league_a)
                target_words = ['premier', 'senior league 1', 'senior league 2', 'senior league 3'] if domain == "Men's" else ['premier', 'senior league 1', 'senior league 2', 'senior league 3', 'senior']
                if any(word in league_str.lower() for word in target_words): 
                    return league_str.replace('NCU', 'Mercury')
                return league_str
                
            return "Possible Cup Match / Friendly"
        except:
            return "Irish Competition" if is_irish else "Unknown Competition"

    comp_map = {}
    if not all_app.empty:
        unique_groups_df = all_app[['Group', 'Is_Irish_Match']].drop_duplicates(subset=['Group'])
        for _, row in unique_groups_df.iterrows():
            grp = row['Group']
            if pd.isna(grp): continue
            is_irish = row.get('Is_Irish_Match', False)
            comp_map[grp] = apply_competition(grp, is_irish)

    zip_buffer = io.BytesIO()
    unregistered_starred_players = []
    run_date = datetime.now()
    
    yellow_fill = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
    red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
    green_fill = PatternFill(start_color="FF92D050", end_color="FF92D050", fill_type="solid")
    bold_font = Font(bold=True)
    center_alignment = Alignment(horizontal='center', vertical='center')

    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        excel_file = pd.ExcelFile(f_starring)
        excluded_tabs = ["summary", "overview", "sheet1"]
        
        for sheet_name in excel_file.sheet_names:
            if sheet_name.lower().strip() in excluded_tabs: continue
                
            club_input = sheet_name.strip()
            safe_club = re.sub(r'[\\/*?:"<>|]', "", club_input).replace(" ", "_")
            df_club, team_match_dates = report_build_club_matches_df(club_input, all_app, override_map, comp_map)
            df_stars_raw = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
            
            players, current_team = [], "Unassigned"
            for i in range(len(df_stars_raw)):
                row = df_stars_raw.iloc[i]
                val0, val_s, val_f = str(row[0]).strip(), str(row[1]).strip(), str(row[4]).strip()
                if 'surname' in val_s.lower() or 'forename' in val_f.lower(): continue
                if val0 and val0.lower() != 'nan' and val0.lower() not in ['id', 'surname', 'forename(s)', 'team']:
                    if not val_s or val_s.lower() == 'nan':
                        current_team = report_clean_team_name(val0) 
                        continue
                if val_s and val_s.lower() != 'nan' and val_f and val_f.lower() != 'nan':
                    players.append({"name": report_clean_spaces(f"{val_f} {val_s}"), "team": current_team})
            
            for p in players:
                if not is_player_registered(p['name']):
                    unregistered_starred_players.append({
                        "Club": club_input, "Starred Team": p['team'], "Starred Name": p['name'],
                        "Current Mapped Alias (if any)": alias_map.get(report_clean_spaces(p['name']).lower(), "None")
                    })
            
            player_list = [p['name'] for p in players]
            player_team_map = {report_clean_spaces(p['name']).lower(): p['team'] for p in players}
            df_p_summary, df_p_details = report_build_player_stats_dfs(player_list, player_team_map, all_app, get_official_name, comp_map)
        
            club_io = io.BytesIO()
            with pd.ExcelWriter(club_io, engine='openpyxl') as writer:
                if not df_club.empty:
                    safe_club_tab_name = report_clean_spaces(club_input)[:20].replace(":", "").replace("/", "")
                    sheet_tab_name = f"{safe_club_tab_name} Matches"
                    df_club.to_excel(writer, sheet_name=sheet_tab_name, index=False)
                    ws_matches = writer.sheets[sheet_tab_name]
                    ws_matches.cell(row=1, column=1).font = bold_font
                    ws_matches.cell(row=1, column=2).font = bold_font
                    for row_idx in range(2, len(df_club) + 2):
                        if str(ws_matches.cell(row=row_idx, column=1).value).startswith("Team:"):
                            ws_matches.cell(row=row_idx, column=1).font = bold_font
                    report_autofit_columns(ws_matches)
                    
                if not df_p_summary.empty:
                    sheet_tab_name = "Player Match Count Summary"
                    df_p_summary.to_excel(writer, sheet_name=sheet_tab_name, index=False)
                    worksheet = writer.sheets[sheet_tab_name]
                    for col_idx in range(1, len(df_p_summary.columns) + 1): worksheet.cell(row=1, column=col_idx).font = bold_font
                    
                    current_starred_team = None 
                    reason_col_summary = len(df_p_summary.columns)
                    
                    for row_idx in range(2, len(df_p_summary) + 2):
                        for col_idx, col_name in enumerate(df_p_summary.columns, start=1):
                            if col_name in ["Matches Played", "Last Played Date"]: worksheet.cell(row=row_idx, column=col_idx).alignment = center_alignment
                        cell_value_1 = str(worksheet.cell(row=row_idx, column=1).value)
                        
                        if cell_value_1.startswith("Team Starred For:"):
                            current_starred_team = cell_value_1.replace("Team Starred For:", "").strip()
                            for col_idx in range(1, len(df_p_summary.columns) + 1): worksheet.cell(row=row_idx, column=col_idx).font = bold_font
                        elif cell_value_1.strip(): 
                            official_name_val = str(worksheet.cell(row=row_idx, column=2).value).strip().lower()
                            is_intl = any(intl in official_name_val or intl in cell_value_1.strip().lower() for intl in international_players)
                            
                            if is_intl:
                                worksheet.cell(row=row_idx, column=reason_col_summary).value = "International Exemption"
                                for col_idx in range(1, len(df_p_summary.columns) + 1): worksheet.cell(row=row_idx, column=col_idx).fill = green_fill
                            else:
                                matches_played = worksheet.cell(row=row_idx, column=3).value 
                                if str(matches_played) == "0":
                                    worksheet.cell(row=row_idx, column=reason_col_summary).value = "0 Matches Played"
                                    for col_idx in range(1, len(df_p_summary.columns) + 1): worksheet.cell(row=row_idx, column=col_idx).fill = red_fill
                                else:
                                    date_str = str(worksheet.cell(row=row_idx, column=4).value) 
                                    if date_str and date_str != "N/A" and date_str.lower() != "nan":
                                        try:
                                            last_played = datetime.strptime(date_str, '%d %B %Y')
                                            days_since = (run_date - last_played).days
                                            team_dates = report_get_team_dates_flexible(current_starred_team, team_match_dates)
                                            matches_since = sum(1 for d in team_dates if d > last_played)
                                            if days_since > 21 and matches_since >= 3:
                                                worksheet.cell(row=row_idx, column=reason_col_summary).value = f"Inactive > 21 days ({days_since} days) and missed {matches_since} matches"
                                                for col_idx in range(1, len(df_p_summary.columns) + 1): worksheet.cell(row=row_idx, column=col_idx).fill = yellow_fill
                                        except ValueError: pass 
                    report_autofit_columns(worksheet)

                if not df_p_details.empty:
                    sheet_tab_name = "Player Season Fixture Log"
                    df_p_details.to_excel(writer, sheet_name=sheet_tab_name, index=False)
                    ws_log = writer.sheets[sheet_tab_name]
                    player_status, current_team_for_dict = {}, "Unassigned"
                    for _, row in df_p_summary.iterrows():
                        if str(row.get('Input Name', '')).startswith('Team Starred For:'): 
                            current_team_for_dict = str(row.get('Input Name', '')).replace("Team Starred For:", "").strip()
                            continue
                        p_name = str(row.get('Player (Official)', '')).strip()
                        if p_name:
                            player_status[p_name] = {'matches': str(row.get('Matches Played', '0')), 'date': str(row.get('Last Played Date', 'N/A')), 'starred_team': current_team_for_dict}
                    
                    for col_idx in range(1, len(df_p_details.columns) + 1): ws_log.cell(row=1, column=col_idx).font = bold_font
                    
                    reason_col_log = len(df_p_details.columns)

                    for row_idx in range(2, len(df_p_details) + 2):
                        cell_value = str(ws_log.cell(row=row_idx, column=1).value)
                        if cell_value.startswith("Player:"):
                            for col_idx in range(1, len(df_p_details.columns) + 1): ws_log.cell(row=row_idx, column=col_idx).font = bold_font
                            try:
                                official_p = cell_value.split(" | Starred Team:")[0].replace("Player: ", "").strip().rsplit(" (", 1)[0].strip()
                                if any(intl in official_p.lower() for intl in international_players):
                                    ws_log.cell(row=row_idx, column=reason_col_log).value = "International Exemption"
                                    for col_idx in range(1, len(df_p_details.columns) + 1): ws_log.cell(row=row_idx, column=col_idx).fill = green_fill
                                else:
                                    status = player_status.get(official_p)
                                    if status:
                                        if status['matches'] == '0':
                                            ws_log.cell(row=row_idx, column=reason_col_log).value = "0 Matches Played"
                                            for col_idx in range(1, len(df_p_details.columns) + 1): ws_log.cell(row=row_idx, column=col_idx).fill = red_fill
                                        else:
                                            d_str = status['date']
                                            if d_str and d_str != "N/A" and d_str.lower() != "nan":
                                                last_played = datetime.strptime(d_str, '%d %B %Y')
                                                t_dates = report_get_team_dates_flexible(status.get('starred_team', ''), team_match_dates)
                                                days_since = (run_date - last_played).days
                                                matches_since = sum(1 for d in t_dates if d > last_played)
                                                if days_since > 21 and matches_since >= 3:
                                                    ws_log.cell(row=row_idx, column=reason_col_log).value = f"Inactive > 21 days ({days_since} days) and missed {matches_since} matches"
                                                    for col_idx in range(1, len(df_p_details.columns) + 1): ws_log.cell(row=row_idx, column=col_idx).fill = yellow_fill
                            except Exception: pass 
                    report_autofit_columns(ws_log)

            zip_file.writestr(f"NCU_Master_Audit_{safe_club}.xlsx", club_io.getvalue())
            
        if unregistered_starred_players:
            unreg_io = io.BytesIO()
            df_unreg = pd.DataFrame(unregistered_starred_players)
            with pd.ExcelWriter(unreg_io, engine='openpyxl') as writer:
                df_unreg.to_excel(writer, sheet_name="Unregistered Players", index=False)
                ws_unreg = writer.sheets["Unregistered Players"]
                for col_idx in range(1, len(df_unreg.columns) + 1): ws_unreg.cell(row=1, column=col_idx).font = bold_font
                report_autofit_columns(ws_unreg)
            zip_file.writestr("Unregistered_Starred_Players.xlsx", unreg_io.getvalue())

    return zip_buffer

# ==========================================
# CLUB FINES GENERATOR FUNCTIONS
# ==========================================
def extract_base_club_name(team_name):
    if pd.isna(team_name): return "Unknown Club"
    t = str(team_name).strip()
    t = re.sub(r'(?i)\b\d(?:st|nd|rd|th)?\s*XI\b', '', t)
    t = re.sub(r'(?i)\b(?:1st|2nd|3rd|4th|5th|6th|7th)\b', '', t)
    t = re.sub(r'(?i)\bWomen\'?s?\b', '', t)
    t = re.sub(r'(?i)\bMW\d?\b', '', t)
    t = re.sub(r'(?i)\bCricket Club\b|\bCC\b', '', t)
    t = re.sub(r'\s+\d$', '', t.strip())
    t = re.sub(r'\s+', ' ', t).strip()
    return t if t else "Unknown Club"

def format_fine_date(dt):
    if pd.isna(dt) or dt is None: return "N/A"
    day = dt.day
    if 11 <= (day % 100) <= 13: suffix = 'th'
    else: suffix = ['th', 'st', 'nd', 'rd', 'th'][min(day % 10, 4)]
    month = dt.strftime('%B')
    return f"{day}{suffix} {month}"

def parse_flexible_date(date_str):
    try:
        if isinstance(date_str, datetime): return date_str
        if pd.isna(date_str): return None
        if hasattr(date_str, 'to_pydatetime'): return date_str.to_pydatetime()
        match = re.search(r'(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)', str(date_str))
        if match:
            day, month_str = int(match.group(1)), match.group(2)
            year = datetime.now().year if datetime.now().year >= 2026 else 2026
            dt = pd.to_datetime(f"{day} {month_str} {year}", errors='coerce')
            if pd.notna(dt): return dt.to_pydatetime()
        dt = pd.to_datetime(str(date_str), errors='coerce', dayfirst=True)
        if pd.notna(dt): return dt.to_pydatetime()
    except: pass
    return None

def extract_competition_from_group(group_str):
    if pd.isna(group_str) or not group_str: return ""
    parts = str(group_str).split(' - ')
    if len(parts) >= 3: return parts[1].strip()
    elif len(parts) == 2:
        if re.search(r'\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+|\d{4}', parts[1]): return ""
        return parts[1].strip()
    return ""

def generate_club_fines_report(audit_file, forfeit_file, start_date, end_date):
    fines_data = []
    
    # Normalize the boundary dates to ensure clean comparisons
    s_bound = pd.to_datetime(start_date).normalize()
    e_bound = pd.to_datetime(end_date).normalize()

    # 1. Parse Forfeited Matches Data
    if forfeit_file:
        try:
            df_forfeit = pd.read_excel(forfeit_file)
            for _, row in df_forfeit.iterrows():
                date_raw = row.get('Date', '')
                team_forfeit = str(row.get('Team Forfeiting', '')).strip()
                team_against = str(row.get('Team against', '')).strip()
                comp = str(row.get('Competition', '')).strip()
                fine = row.get('Fine', 0)
                
                if team_forfeit.lower() == 'nan' or not team_forfeit: continue
                try: fine = int(fine)
                except: fine = 0
                
                club = extract_base_club_name(team_forfeit)
                date_obj = parse_flexible_date(date_raw)
                
                # DATE FILTER LOGIC: Skip if outside the selected range
                if date_obj:
                    match_dt = pd.to_datetime(date_obj).normalize()
                    if match_dt < s_bound or match_dt > e_bound:
                        continue
                else:
                    continue
                    
                date_str = format_fine_date(date_obj) if date_obj else str(date_raw)
                team_part_str = f"{team_forfeit} (v {team_against})"
                
                fines_data.append({
                    'Club': club, 'Date_obj': date_obj, 'Date_str': date_str,
                    'Reason': 'Unable to field a team', 'Player': None,
                    'Team_Part_Str': team_part_str,
                    'Competition': comp, 'Fine': fine, 'Type': 'Team'
                })
        except Exception: pass

    # 2. Parse the output of the internal audit engine
    if audit_file:
        try:
            excel_file = pd.ExcelFile(audit_file)
            
            # --- Pre-process to cross-reference common teams for unregistered players ---
            df_unreg = pd.read_excel(audit_file, sheet_name="Unregistered Matches") if "Unregistered Matches" in excel_file.sheet_names else pd.DataFrame()
            df_deemed = pd.read_excel(audit_file, sheet_name="Deemed Registered") if "Deemed Registered" in excel_file.sheet_names else pd.DataFrame()
            
            player_true_team = {}
            if not df_unreg.empty:
                # Combine their 1st occurrence with any subsequent matches they played
                all_unreg_matches = pd.concat([df_unreg, df_deemed], ignore_index=True) if not df_deemed.empty else df_unreg
                
                if 'Stats Name (Cleaned)' in all_unreg_matches.columns:
                    for player, group in all_unreg_matches.groupby('Stats Name (Cleaned)'):
                        reg_club = str(group.iloc[0].get('Registered Club', 'Unknown Club')).strip()
                        reg_club_base = extract_base_club_name(reg_club).lower()
                        
                        # 1. If we eventually know their registered club (e.g., they registered late)
                        if reg_club_base != 'unknown club':
                            player_true_team[player] = ('known_reg', reg_club)
                        
                        # 2. If the engine appended the club to their name via KNOWN_DUPLICATES
                        elif '(' in player and player.strip().endswith(')'):
                            club_in_name = player.split('(')[-1].replace(')', '').strip().lower()
                            player_true_team[player] = ('inferred', club_in_name)
                            
                        # 3. Completely unregistered: look for common teams across matches
                        else:
                            teams_in_matches = []
                            for _, r in group.iterrows():
                                t_a = str(r.get('Team A', '')).strip()
                                t_b = str(r.get('Team B', '')).strip()
                                teams_in_matches.append({extract_base_club_name(t_a).lower(), extract_base_club_name(t_b).lower()})
                            
                            if len(teams_in_matches) == 1:
                                t_a = str(group.iloc[0].get('Team A', '')).strip()
                                t_b = str(group.iloc[0].get('Team B', '')).strip()
                                player_true_team[player] = ('ambiguous', (t_a, t_b))
                            else:
                                common_teams = set.intersection(*teams_in_matches)
                                if len(common_teams) == 1:
                                    player_true_team[player] = ('inferred', list(common_teams)[0])
                                else:
                                    t_a = str(group.iloc[0].get('Team A', '')).strip()
                                    t_b = str(group.iloc[0].get('Team B', '')).strip()
                                    player_true_team[player] = ('ambiguous', (t_a, t_b))

            # --- Extract Unregistered Player Fines ---
            if not df_unreg.empty and len(df_unreg.columns) > 1:
                for _, row in df_unreg.iterrows():
                    match_date = row.get('Match Date')
                    date_obj = pd.to_datetime(match_date) if pd.notna(match_date) else None
                    date_str = format_fine_date(date_obj) if date_obj else str(match_date)
                    
                    player_key = str(row.get('Stats Name (Cleaned)', '')).strip()
                    player_disp = str(row.get('Original Scorecard Name', player_key)).strip()
                    
                    team_a = str(row.get('Team A', '')).strip()
                    team_b = str(row.get('Team B', '')).strip()
                    comp = str(row.get('Match League', '')).strip()
                    
                    # Fetch the cross-referenced team status
                    status, info = player_true_team.get(player_key, ('known_reg', str(row.get('Registered Club', 'Unknown Club')).strip()))
                    
                    if status == 'known_reg':
                        reg_club_base = extract_base_club_name(info).lower()
                        if reg_club_base in extract_base_club_name(team_b).lower() and reg_club_base != 'unknown club':
                            team_played, opponent = team_b, team_a
                        else:
                            team_played, opponent = team_a, team_b
                        club = extract_base_club_name(team_played)
                        team_part_str = f"{team_played} (v {opponent})"
                    
                    elif status == 'inferred':
                        if info in extract_base_club_name(team_b).lower():
                            team_played, opponent = team_b, team_a
                        else:
                            team_played, opponent = team_a, team_b
                        club = extract_base_club_name(team_played)
                        team_part_str = f"{team_played} (v {opponent})"
                        
                    elif status == 'ambiguous':
                        t_a, t_b = info
                        club = f"{extract_base_club_name(t_a)} / {extract_base_club_name(t_b)}"
                        team_part_str = f"{t_a} v {t_b}" 
                        
                    fines_data.append({
                        'Club': club, 'Date_obj': date_obj, 'Date_str': date_str,
                        'Reason': 'playing an unregistered player', 'Player': player_disp,
                        'Team_Part_Str': team_part_str,
                        'Competition': comp, 'Fine': 10, 'Type': 'Player'
                    })
                        
            # --- Extract Starring Violation Fines ---
            if "Starring Violations" in excel_file.sheet_names:
                df_star = pd.read_excel(audit_file, sheet_name="Starring Violations")
                if not df_star.empty and len(df_star.columns) > 1:
                    for _, row in df_star.iterrows():
                        match_date = row.get('Match Date')
                        date_obj = pd.to_datetime(match_date) if pd.notna(match_date) else None
                        date_str = format_fine_date(date_obj) if date_obj else str(match_date)
                        
                        player = str(row.get('Original Scorecard Name', row.get('Player (Cleaned)', ''))).strip()
                        team_played = str(row.get('Actually Played For', row.get('Midweek Team', ''))).strip()
                        team_a = str(row.get('Team A', '')).strip()
                        team_b = str(row.get('Team B', '')).strip()
                        
                        opponent = team_b if team_played.lower() == team_a.lower() else team_a
                        comp = extract_competition_from_group(str(row.get('Match Group', '')))
                        club = extract_base_club_name(team_played)
                        
                        team_part_str = f"{team_played} (v {opponent})"
                        
                        fines_data.append({
                            'Club': club, 'Date_obj': date_obj, 'Date_str': date_str,
                            'Reason': 'playing a starred player', 'Player': player,
                            'Team_Part_Str': team_part_str,
                            'Competition': comp, 'Fine': 10, 'Type': 'Player'
                        })
        except Exception: pass

    # Sort alphabetically by Club Name, then chronologically by Match Date
    def sort_key(x):
        d = x['Date_obj'] if pd.notna(x['Date_obj']) and x['Date_obj'] is not None else datetime.min
        return (x['Club'].lower(), d)
    
    fines_data.sort(key=sort_key)
    
    from collections import defaultdict
    fines_by_club = defaultdict(list)
    for f in fines_data:
        fines_by_club[f['Club']].append(f)
        
    doc = Document()
    style_normal = doc.styles['Normal']
    style_normal.font.name, style_normal.font.size = 'Calibri', Pt(11)
    
    p_title = doc.add_paragraph()
    r_title = p_title.add_run("Club Fines Report")
    r_title.bold = True
    r_title.font.size = Pt(16)
    
    for club in sorted(fines_by_club.keys(), key=lambda c: c.lower()):
        doc.add_paragraph() 
        
        p_club = doc.add_paragraph()
        r_club = p_club.add_run(club)
        r_club.bold = True
        r_club.font.size = Pt(12)
        
        for f in fines_by_club[club]:
            p_fine = doc.add_paragraph()
            
            date_part = f['Date_str']
            reason_part = f['Reason']
            player_part = f" - {f['Player']}" if f['Type'] == 'Player' else ""
            team_part = f"{f['Team_Part_Str']}"
            comp_part = f" – {f['Competition']}" if f['Competition'] and str(f['Competition']).lower() != 'nan' else ""
            fine_part = f"Fine: £{f['Fine']}"
            
            r_date = p_fine.add_run(f"{date_part}")
            r_date.bold = True
            
            text_str = f" – {reason_part}{player_part} - {team_part}{comp_part} - "
            p_fine.add_run(text_str)
            
            r_fine = p_fine.add_run(fine_part)
            r_fine.bold = True
            
            p_fine.paragraph_format.space_after = Pt(6)
            
    doc_io = io.BytesIO()
    doc.save(doc_io)
    return doc_io
    
# ==========================================
# UNREGISTERED ONLY FINES GENERATOR 17-07-2026
# ==========================================
def generate_unregistered_fines_only(audit_file):
    from collections import defaultdict
    fines_data = []
    
    if audit_file:
        try:
            excel_file = pd.ExcelFile(audit_file)
            
            # --- Pre-process sheets ---
            df_unreg = pd.read_excel(audit_file, sheet_name="Unregistered Matches") if "Unregistered Matches" in excel_file.sheet_names else pd.DataFrame()
            df_deemed = pd.read_excel(audit_file, sheet_name="Deemed Registered") if "Deemed Registered" in excel_file.sheet_names else pd.DataFrame()
            
            # Map out each player's exact matched team layout (handling ambiguity/duplicates)
            player_true_team = {}
            player_deemed_matches = defaultdict(list)
            
            # 1. Compile a lookup dictionary of subsequent matches played by each player
            if not df_deemed.empty and 'Stats Name (Cleaned)' in df_deemed.columns:
                for _, r in df_deemed.iterrows():
                    p_key = str(r.get('Stats Name (Cleaned)', '')).strip()
                    m_date = r.get('Match Date')
                    d_obj = pd.to_datetime(m_date) if pd.notna(m_date) else None
                    d_str = format_fine_date(d_obj) if d_obj else str(m_date)
                    t_a = str(r.get('Team A', '')).strip()
                    t_b = str(r.get('Team B', '')).strip()
                    comp = str(r.get('Match League', '')).strip()
                    
                    # Format subsequent match presentation string
                    match_desc = f"{d_str} – Deemed Registered Match: {t_a} v {t_b} ({comp})"
                    player_deemed_matches[p_key].append(match_desc)

            if not df_unreg.empty:
                all_unreg_matches = pd.concat([df_unreg, df_deemed], ignore_index=True) if not df_deemed.empty else df_unreg
                
                if 'Stats Name (Cleaned)' in all_unreg_matches.columns:
                    for player, group in all_unreg_matches.groupby('Stats Name (Cleaned)'):
                        reg_club = str(group.iloc[0].get('Registered Club', 'Unknown Club')).strip()
                        reg_club_base = extract_base_club_name(reg_club).lower()
                        
                        if reg_club_base != 'unknown club':
                            player_true_team[player] = ('known_reg', reg_club)
                        elif '(' in player and player.strip().endswith(')'):
                            club_in_name = player.split('(')[-1].replace(')', '').strip().lower()
                            player_true_team[player] = ('inferred', club_in_name)
                        else:
                            teams_in_matches = []
                            for _, r in group.iterrows():
                                t_a = str(r.get('Team A', '')).strip()
                                t_b = str(r.get('Team B', '')).strip()
                                teams_in_matches.append({extract_base_club_name(t_a).lower(), extract_base_club_name(t_b).lower()})
                            
                            if len(teams_in_matches) == 1:
                                t_a = str(group.iloc[0].get('Team A', '')).strip()
                                t_b = str(group.iloc[0].get('Team B', '')).strip()
                                player_true_team[player] = ('ambiguous', (t_a, t_b))
                            else:
                                common_teams = set.intersection(*teams_in_matches)
                                if len(common_teams) == 1:
                                    player_true_team[player] = ('inferred', list(common_teams)[0])
                                else:
                                    t_a = str(group.iloc[0].get('Team A', '')).strip()
                                    t_b = str(group.iloc[0].get('Team B', '')).strip()
                                    player_true_team[player] = ('ambiguous', (t_a, t_b))

            # --- Extract Unregistered Player Fines ---
            if not df_unreg.empty and len(df_unreg.columns) > 1:
                for _, row in df_unreg.iterrows():
                    match_date = row.get('Match Date')
                    date_obj = pd.to_datetime(match_date) if pd.notna(match_date) else None
                    date_str = format_fine_date(date_obj) if date_obj else str(match_date)
                    
                    player_key = str(row.get('Stats Name (Cleaned)', '')).strip()
                    player_disp = str(row.get('Original Scorecard Name', player_key)).strip()
                    
                    team_a = str(row.get('Team A', '')).strip()
                    team_b = str(row.get('Team B', '')).strip()
                    comp = str(row.get('Match League', '')).strip()
                    
                    status, info = player_true_team.get(player_key, ('known_reg', str(row.get('Registered Club', 'Unknown Club')).strip()))
                    
                    if status == 'known_reg':
                        reg_club_base = extract_base_club_name(info).lower()
                        if reg_club_base in extract_base_club_name(team_b).lower() and reg_club_base != 'unknown club':
                            team_played, opponent = team_b, team_a
                        else:
                            team_played, opponent = team_a, team_b
                        club = extract_base_club_name(team_played)
                        team_part_str = f"{team_played} (v {opponent})"
                    
                    elif status == 'inferred':
                        if info in extract_base_club_name(team_b).lower():
                            team_played, opponent = team_b, team_a
                        else:
                            team_played, opponent = team_a, team_b
                        club = extract_base_club_name(team_played)
                        team_part_str = f"{team_played} (v {opponent})"
                        
                    elif status == 'ambiguous':
                        t_a, t_b = info
                        club = f"{extract_base_club_name(t_a)} / {extract_base_club_name(t_b)}"
                        team_part_str = f"{t_a} v {t_b}" 
                    
                    # Fetch compiled list of subsequent matches played
                    subsequent_matches = player_deemed_matches.get(player_key, [])
                        
                    fines_data.append({
                        'Club': club, 'Date_obj': date_obj, 'Date_str': date_str,
                        'Reason': 'playing an unregistered player', 'Player': player_disp,
                        'Team_Part_Str': team_part_str,
                        'Competition': comp, 'Fine': 10, 'Type': 'Player',
                        'Deemed_Matches': subsequent_matches
                    })
        except Exception as e: 
            print(f"Error parsing audit file: {e}")

    # Sort alphabetically by Club Name, then chronologically by Match Date
    def sort_key(x):
        d = x['Date_obj'] if pd.notna(x['Date_obj']) and x['Date_obj'] is not None else datetime.min
        return (x['Club'].lower(), d)
    
    fines_data.sort(key=sort_key)
    
    fines_by_club = defaultdict(list)
    for f in fines_data:
        fines_by_club[f['Club']].append(f)
        
    doc = Document()
    style_normal = doc.styles['Normal']
    style_normal.font.name, style_normal.font.size = 'Calibri', Pt(11)
    
    p_title = doc.add_paragraph()
    r_title = p_title.add_run("Unregistered Player Fines Report")
    r_title.bold = True
    r_title.font.size = Pt(16)
    
    for club in sorted(fines_by_club.keys(), key=lambda c: c.lower()):
        doc.add_paragraph() 
        
        p_club = doc.add_paragraph()
        r_club = p_club.add_run(club)
        r_club.bold = True
        r_club.font.size = Pt(12)
        
        for f in fines_by_club[club]:
            p_fine = doc.add_paragraph()
            
            date_part = f['Date_str']
            reason_part = f['Reason']
            player_part = f" - {f['Player']}"
            team_part = f"{f['Team_Part_Str']}"
            comp_part = f" – {f['Competition']}" if f['Competition'] and str(f['Competition']).lower() != 'nan' else ""
            fine_part = f"Fine: £{f['Fine']}"
            
            r_date = p_fine.add_run(f"{date_part}")
            r_date.bold = True
            
            text_str = f" – {reason_part}{player_part} - {team_part}{comp_part} - "
            p_fine.add_run(text_str)
            
            r_fine = p_fine.add_run(fine_part)
            r_fine.bold = True
            
            # If the player went on to play more matches while unregistered, list them right below
            if f['Deemed_Matches']:
                p_fine.paragraph_format.space_after = Pt(2)
                
                p_info = doc.add_paragraph()
                p_info.paragraph_format.left_indent = Pt(18)
                p_info.paragraph_format.space_before = Pt(0)
                p_info.paragraph_format.space_after = Pt(3)
                r_info = p_info.add_run(f"→ Subsequent matches played while deemed registered ({len(f['Deemed_Matches'])}):")
                r_info.font.italic = True
                r_info.font.size = Pt(10)
                
                for sub_m in f['Deemed_Matches']:
                    p_sub = doc.add_paragraph()
                    p_sub.paragraph_format.left_indent = Pt(36)
                    p_sub.paragraph_format.space_before = Pt(0)
                    p_sub.paragraph_format.space_after = Pt(2)
                    r_sub = p_sub.add_run(f"• {sub_m}")
                    r_sub.font.size = Pt(10)
                    r_sub.font.color.rgb = RGBColor(100, 100, 100)
            else:
                p_fine.paragraph_format.space_after = Pt(6)
            
    doc_io = io.BytesIO()
    doc.save(doc_io)
    return doc_io