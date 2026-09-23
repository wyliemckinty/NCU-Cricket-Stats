"""
starring_rules.py
=================
Core validation and rules engine for NCU Rules A10–A14:
1. Rule A10: Automatic Starring Quota Checker by club team count.
2. Rule A11 & A12: Absence & De-Starring Tracker with scorecard reader integration,
   Irish International Duty exemption, and July 31st roster modification deadline lock.
3. Rule A13: Transfer Monitor, 2-transfer seasonal ceiling enforcement, and
   dynamic £25.00 transfer fee infraction calculation for finance_app.py.

Helper Apps:
    app.py, finance_app.py, engine.py, tests/test_starring_rules.py.
"""

from datetime import datetime
import os
import re
from typing import Any, Dict, List, Optional, Set, Tuple
import openpyxl
import pandas as pd
from engine import read_excel_calamine, strip_club_suffix, format_excel_worksheet_standard

# Canonical tier hierarchy ordered from highest (1st XI) to lowest (6th XI)
TIER_HIERARCHY = ["1st XI", "2nd XI", "3rd XI", "4th XI", "5th XI", "6th XI"]
TIER_RANK_MAP = {tier: idx for idx, tier in enumerate(TIER_HIERARCHY, start=1)}


def get_tier_level_rank(tier_name: str) -> int:
    """
    Returns numeric rank for a team tier (1 for 1st XI, 2 for 2nd XI, etc.).
    Lower numeric value denotes a higher team tier.

    Inputs:
        tier_name: Name of the team tier (e.g. '1st XI', '2nd XI', '3rd XI').

    Outputs:
        int: Rank 1 through 6, or 99 if unrecognized.

    Helper Apps:
        starring_rules.py, app.py.
    """
    t_clean = str(tier_name).strip()
    for tier, rank in TIER_RANK_MAP.items():
        if tier.lower() in t_clean.lower():
            return rank
    return 99


def sort_starring_roster_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Sorts a starring roster dataframe according to NCU administrative standards:
    1. Team tier in ascending order (1st XI < 2nd XI < 3rd XI < 4th XI < 5th XI)
    2. Surname (or Last Name) alphabetically ascending
    3. First Name (or Forename) alphabetically ascending

    Inputs:
        df: Input DataFrame containing starring records.

    Outputs:
        pd.DataFrame: Sorted DataFrame with reset index.

    Helper Apps:
        app.py, engine.py, starring_rules.py, tests/test_starring_rules.py.
    """
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()

    res = df.copy()

    # 1. Resolve Tier / Team column
    tier_col = next((c for c in ["XI_Level", "Starred Tier", "Team", "Tier"] if c in res.columns), None)
    if tier_col:
        res["_tier_rank"] = res[tier_col].apply(lambda t: get_tier_level_rank(str(t)))
    else:
        res["_tier_rank"] = 99

    # 2. Resolve Surname / Last Name column
    if "Surname" in res.columns:
        res["_surname"] = res["Surname"].fillna("").astype(str).str.strip().str.lower()
    elif "Last Name" in res.columns:
        res["_surname"] = res["Last Name"].fillna("").astype(str).str.strip().str.lower()
    elif "Full Name" in res.columns:
        res["_surname"] = res["Full Name"].apply(lambda n: strip_club_suffix(str(n)).strip().split()[-1].lower() if str(n).strip() else "")
    elif "Player" in res.columns:
        res["_surname"] = res["Player"].apply(lambda n: strip_club_suffix(str(n)).strip().split()[-1].lower() if str(n).strip() else "")
    elif "Name" in res.columns:
        res["_surname"] = res["Name"].apply(lambda n: strip_club_suffix(str(n)).strip().split()[-1].lower() if str(n).strip() else "")
    else:
        res["_surname"] = ""

    # 3. Resolve First Name / Forename column
    if "Forename" in res.columns:
        res["_forename"] = res["Forename"].fillna("").astype(str).str.strip().str.lower()
    elif "First Name" in res.columns:
        res["_forename"] = res["First Name"].fillna("").astype(str).str.strip().str.lower()
    elif "Full Name" in res.columns:
        res["_forename"] = res["Full Name"].apply(lambda n: str(n).strip().split()[0].lower() if str(n).strip() else "")
    elif "Player" in res.columns:
        res["_forename"] = res["Player"].apply(lambda n: str(n).strip().split()[0].lower() if str(n).strip() else "")
    elif "Name" in res.columns:
        res["_forename"] = res["Name"].apply(lambda n: str(n).strip().split()[0].lower() if str(n).strip() else "")
    else:
        res["_forename"] = ""

    res = res.sort_values(by=["_tier_rank", "_surname", "_forename"], ascending=[True, True, True]).reset_index(drop=True)
    res = res.drop(columns=["_tier_rank", "_surname", "_forename"])

    # Re-sequence the Rank column within each tier as a 1-based sequential slot counter (1 to 8, 1 to 10, etc.)
    if "Rank" in res.columns:
        if tier_col and tier_col in res.columns:
            res["Rank"] = res.groupby(tier_col).cumcount() + 1
        else:
            res["Rank"] = range(1, len(res) + 1)

    return res


def sort_transfer_records_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Sorts a transfer records dataframe according to NCU administrative standards:
    1. Transfer Date chronologically ascending (oldest to newest)
    2. Surname alphabetically ascending
    3. First Name alphabetically ascending

    Inputs:
        df: Input DataFrame containing transfer records.

    Outputs:
        pd.DataFrame: Sorted DataFrame with reset index.

    Helper Apps:
        app.py, engine.py, starring_rules.py, tests/test_starring_rules.py.
    """
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()

    res = df.copy()

    # 1. Parse / resolve Transfer Date
    if "Transfer Date Parsed" in res.columns:
        res["_date_sort"] = pd.to_datetime(res["Transfer Date Parsed"], errors="coerce")
    elif "Transfer Date" in res.columns:
        res["_date_sort"] = pd.to_datetime(res["Transfer Date"], errors="coerce", dayfirst=True)
    elif "Date" in res.columns:
        res["_date_sort"] = pd.to_datetime(res["Date"], errors="coerce", dayfirst=True)
    else:
        res["_date_sort"] = pd.Timestamp.max

    res["_date_sort"] = res["_date_sort"].fillna(pd.Timestamp.max)

    # 2. Resolve Surname
    if "Surname" in res.columns:
        res["_surname"] = res["Surname"].fillna("").astype(str).str.strip().str.lower()
    elif "Last Name" in res.columns:
        res["_surname"] = res["Last Name"].fillna("").astype(str).str.strip().str.lower()
    elif "Player" in res.columns:
        res["_surname"] = res["Player"].apply(lambda n: str(n).strip().split()[-1].lower() if str(n).strip() else "")
    elif "Full Name" in res.columns:
        res["_surname"] = res["Full Name"].apply(lambda n: str(n).strip().split()[-1].lower() if str(n).strip() else "")
    else:
        res["_surname"] = ""

    # 3. Resolve First Name
    if "First Name" in res.columns:
        res["_forename"] = res["First Name"].fillna("").astype(str).str.strip().str.lower()
    elif "Forename" in res.columns:
        res["_forename"] = res["Forename"].fillna("").astype(str).str.strip().str.lower()
    elif "Player" in res.columns:
        res["_forename"] = res["Player"].apply(lambda n: str(n).strip().split()[0].lower() if str(n).strip() else "")
    elif "Full Name" in res.columns:
        res["_forename"] = res["Full Name"].apply(lambda n: str(n).strip().split()[0].lower() if str(n).strip() else "")
    else:
        res["_forename"] = ""

    res = res.sort_values(by=["_date_sort", "_surname", "_forename"], ascending=[True, True, True]).reset_index(drop=True)
    res = res.drop(columns=["_date_sort", "_surname", "_forename"])
    return res


# ==============================================================================
# RULE A10: AUTOMATIC QUOTA CHECKER
# ==============================================================================

def extract_base_club_name(name: str) -> str:
    """Helper to extract base club name."""
    return str(name).strip()


def get_starring_quotas(num_teams: int, domain: str = "Men's") -> Dict[str, int]:
    """
    Returns the exact starring quotas per team tier according to NCU Rule A10 (Men's) or Rule WA10 (Women's).

    Rule A10 Schedule (Men's):
    - 2 Teams: Exactly 8 players for 1st XI.
    - 3 Teams: 3 remaining 1st XI + 7 players for 2nd XI (1st XI: 8, 2nd XI: 10).
    - 4 Teams: Next 3 from 2nd XI + 5 players for 3rd XI (1st XI: 8, 2nd XI: 10, 3rd XI: 8).
    - 5 Teams: Next 5 from 3rd XI + 2 players for 4th XI (1st XI: 8, 2nd XI: 10, 3rd XI: 8, 4th XI: 7).
    - 6 Teams: Next 7 from 4th XI for 5th XI (1st XI: 8, 2nd XI: 10, 3rd XI: 8, 4th XI: 7, 5th XI: 7).

    Rule WA10 Schedule (Women's):
    - 2 Teams: First 7 players normally selected for 1st XI (1st XI: 7).
    - 3 Teams: 4 remaining 1st XI + first 5 normally selected for 2nd XI (1st XI: 7, 2nd XI: 9).
    - 4 Teams: Next 6 from 2nd XI + first 5 normally selected for 3rd XI (1st XI: 7, 2nd XI: 9, 3rd XI: 11).
    Note: Lowest team in any club is never starred.

    Inputs:
        num_teams: Total number of senior teams fielded by the club (e.g. 2 to 6).
        domain: Competition domain ("Men's" or "Women's").

    Outputs:
        Dict[str, int]: Dictionary mapping tier names to required starred player counts.

    Helper Apps:
        app.py, engine.py, tests/test_starring_rules.py.
    """
    if num_teams <= 1:
        return {}

    if "women" in domain.lower():
        if num_teams == 2:
            return {"1st XI": 7}
        elif num_teams == 3:
            return {"1st XI": 7, "2nd XI": 9}
        elif num_teams == 4:
            return {"1st XI": 7, "2nd XI": 9, "3rd XI": 11}
        else:
            return {"1st XI": 7, "2nd XI": 9, "3rd XI": 11, "4th XI": 7}

    if num_teams == 2:
        return {"1st XI": 8}
    elif num_teams == 3:
        return {"1st XI": 8, "2nd XI": 10}
    elif num_teams == 4:
        return {"1st XI": 8, "2nd XI": 10, "3rd XI": 8}
    elif num_teams == 5:
        return {"1st XI": 8, "2nd XI": 10, "3rd XI": 8, "4th XI": 7}
    else:  # 6 or more teams
        return {"1st XI": 8, "2nd XI": 10, "3rd XI": 8, "4th XI": 7, "5th XI": 7}


def get_club_senior_team_count(
    club_name: str,
    domain: str = "Men's",
    club_starring_df: Optional[pd.DataFrame] = None,
    all_club_counts: Optional[Dict[str, Dict[str, int]]] = None
) -> int:
    """
    Calculates the effective senior team count for a given club to drive Rule A10 / WA10 quotas.

    Priority:
    1. Highest team Roman numeral from league structure / all_club_counts (e.g. 'CSNI 5' -> 5).
    2. Highest starring tier in club's submitted starring list (e.g. '5th XI' requires 6 senior teams).
    3. Minimum fallback of 2 senior teams (or 1 for women's).

    Inputs:
        club_name: Normalized club name.
        domain: Competition domain ("Men's" or "Women's").
        club_starring_df: Optional DataFrame of club's starred players.
        all_club_counts: Optional dictionary returned by get_all_club_team_counts().

    Outputs:
        int: Total number of senior teams.

    Helper Apps:
        app.py, tests/test_starring_rules.py.
    """
    clean_club = extract_base_club_name(club_name).strip()
    league_count = 0

    if all_club_counts and clean_club in all_club_counts:
        dom_key = "women" if "women" in domain.lower() else "men"
        league_count = all_club_counts[clean_club].get(dom_key, 0)

    # Cross-reference submitted starring tiers: if starring list contains a tier,
    # the club must field at least (tier_index + 1) teams.
    starring_count = 0
    if club_starring_df is not None and not club_starring_df.empty:
        tier_col = "XI_Level" if "XI_Level" in club_starring_df.columns else ("Team" if "Team" in club_starring_df.columns else club_starring_df.columns[0])
        tiers = club_starring_df[tier_col].dropna().astype(str).str.strip().str.lower().unique()
        if any("5th xi" in t for t in tiers):
            starring_count = 6
        elif any("4th xi" in t for t in tiers):
            starring_count = 5
        elif any("3rd xi" in t for t in tiers):
            starring_count = 4
        elif any("2nd xi" in t for t in tiers):
            starring_count = 3
        elif any("1st xi" in t for t in tiers):
            starring_count = 2

    final_count = max(league_count, starring_count)
    if "women" in domain.lower():
        return max(final_count, 1)
    return max(final_count, 2)


def validate_club_starring_quotas(
    club_starring_df: pd.DataFrame,
    num_teams: int,
    domain: str = "Men's"
) -> Dict[str, Dict[str, Any]]:
    """
    Enforces and validates exact starring slots per tier against Rule A10 (Men's) or Rule WA10 (Women's) quotas.

    Inputs:
        club_starring_df: DataFrame of starred players containing 'XI_Level' and player names.
        num_teams: Total number of senior teams fielded by the club.
        domain: Competition domain ("Men's" or "Women's").

    Outputs:
        Dict[str, Dict[str, Any]]: Per-tier validation dictionary containing:
            - 'expected': int (quota required)
            - 'actual': int (actual starred count)
            - 'is_complete': bool (actual == expected)
            - 'status_label': str (e.g. "1st XI Starring: 8/8 Complete")
            - 'players': List[str] of player names in that tier
            - 'delta': int (actual - expected)

    Helper Apps:
        app.py, tests/test_starring_rules.py.
    """
    quotas = get_starring_quotas(num_teams, domain=domain)
    results: Dict[str, Dict[str, Any]] = {}

    if club_starring_df is None or club_starring_df.empty:
        for tier, expected in quotas.items():
            results[tier] = {
                "expected": expected,
                "actual": 0,
                "is_complete": False,
                "status_label": f"{tier} Starring: 0/{expected} Incomplete",
                "players": [],
                "delta": -expected
            }
        return results

    df = club_starring_df.copy()
    tier_col = "XI_Level" if "XI_Level" in df.columns else ("Team" if "Team" in df.columns else df.columns[0])
    name_col = "Full Name" if "Full Name" in df.columns else ("Name" if "Name" in df.columns else ("Surname" if "Surname" in df.columns else df.columns[1]))

    for tier, expected in quotas.items():
        mask = df[tier_col].astype(str).str.strip().str.lower() == tier.lower()
        tier_players_df = df[mask]
        actual = len(tier_players_df)
        players = tier_players_df[name_col].dropna().astype(str).str.strip().tolist()

        is_complete = (actual == expected)
        if actual == expected:
            status_label = f"{tier} Starring: {actual}/{expected} Complete"
        elif actual < expected:
            status_label = f"{tier} Starring: {actual}/{expected} Incomplete"
        else:
            status_label = f"{tier} Starring: {actual}/{expected} Over Quota"

        results[tier] = {
            "expected": expected,
            "actual": actual,
            "is_complete": is_complete,
            "status_label": status_label,
            "players": players,
            "delta": actual - expected
        }

    return results


def check_starred_roster_registration(
    club_star_df: pd.DataFrame,
    df_reg: Optional[pd.DataFrame] = None,
    alias_map: Optional[Dict[str, str]] = None
) -> pd.DataFrame:
    """
    Checks each starred player in a club's roster against official NCU registration records.
    Appends a 'Registered' column containing '✅' for registered players and '❌' for unregistered players.
    Supports dynamic synthesis of 'Full Name' from 'First Name' + 'Last Name' if absent, case normalization,
    middle name stripping, first/last token cross-matching ([Forename Surname] and [Surname Forename]),
    and alias resolution against NCU_Validated_Aliases_Master.xlsx.

    Inputs:
        club_star_df: DataFrame of starred players containing player names.
        df_reg: Optional DataFrame of registration records (from Sport80/NCU registry).
        alias_map: Optional dictionary of known player name aliases.

    Outputs:
        pd.DataFrame: Copy of club_star_df with 'Registered' column added.

    Helper Apps:
        app.py, starring_rules.py, tests/test_starring_rules.py.
    """
    if club_star_df is None or club_star_df.empty:
        df_empty = pd.DataFrame() if club_star_df is None else club_star_df.copy()
        if "Registered" not in df_empty.columns:
            df_empty["Registered"] = []
        return df_empty

    df = club_star_df.copy()

    if df_reg is None or df_reg.empty:
        df["Registered"] = "❌"
        return df

    alias_dict: Dict[str, str] = {
        str(k).strip().lower(): str(v).strip().lower()
        for k, v in (alias_map or {}).items()
        if str(k).strip() and str(v).strip()
    }

    clean_regex = lambda s: re.sub(r"[^a-zA-Z0-9\s]", "", str(s).lower()).strip()

    def get_tokens(name_str: str) -> List[str]:
        cleaned = clean_regex(name_str)
        return [t for t in cleaned.split() if t]

    reg_df = df_reg.copy()

    # Dynamic synthesis of Full Name from First Name and Last Name if missing
    if "Full Name" not in reg_df.columns:
        f_cand = next((c for c in reg_df.columns if "first" in str(c).lower() or "forename" in str(c).lower()), None)
        l_cand = next((c for c in reg_df.columns if "last" in str(c).lower() or "surname" in str(c).lower()), None)
        if f_cand and l_cand:
            reg_df["Full Name"] = (
                reg_df[f_cand].fillna("").astype(str).str.strip() + " " +
                reg_df[l_cand].fillna("").astype(str).str.strip()
            ).str.strip()

    reg_name_col = next(
        (c for c in ["Full Name", "Name", "Player", "Registered Name", "Player Name"] if c in reg_df.columns),
        reg_df.columns[0]
    )

    raw_reg_names: Set[str] = set()
    clean_reg_names: Set[str] = set()
    reg_token_sets: Set[frozenset] = set()
    reg_sig_token_sets: Set[frozenset] = set()
    reg_first_last_pairs: Set[Tuple[str, str]] = set()

    # Build registration lookups across all name representations
    for _, r in reg_df.iterrows():
        names_in_row = []
        val = r.get(reg_name_col, None)
        if pd.notna(val) and str(val).strip():
            names_in_row.append(str(val).strip())

        f_col = next((c for c in reg_df.columns if "first" in str(c).lower() or "forename" in str(c).lower()), None)
        l_col = next((c for c in reg_df.columns if "last" in str(c).lower() or "surname" in str(c).lower()), None)
        if f_col and l_col and pd.notna(r.get(f_col)) and pd.notna(r.get(l_col)):
            f_val = str(r[f_col]).strip()
            l_val = str(r[l_col]).strip()
            if f_val and l_val:
                names_in_row.append(f"{f_val} {l_val}")
                names_in_row.append(f"{l_val} {f_val}")

        for nm in names_in_row:
            nm_lower = nm.lower().strip()
            raw_reg_names.add(nm_lower)
            nm_clean = clean_regex(nm)
            if nm_clean:
                clean_reg_names.add(nm_clean)
            tokens = get_tokens(nm)
            if tokens:
                reg_token_sets.add(frozenset(tokens))
                sig_tokens = [t for t in tokens if len(t) > 1]
                if sig_tokens:
                    reg_sig_token_sets.add(frozenset(sig_tokens))
                if len(tokens) >= 2:
                    reg_first_last_pairs.add((tokens[0], tokens[-1]))
                    reg_first_last_pairs.add((tokens[-1], tokens[0]))
                if len(sig_tokens) >= 2:
                    reg_first_last_pairs.add((sig_tokens[0], sig_tokens[-1]))
                    reg_first_last_pairs.add((sig_tokens[-1], sig_tokens[0]))

    # Determine name representations in club_star_df
    f_cols = [c for c in df.columns if "forename" in str(c).lower()]
    s_cols = [c for c in df.columns if "surname" in str(c).lower()]
    star_name_col = next((c for c in ["Full Name", "Name", "Player"] if c in df.columns), None)

    if not star_name_col:
        if f_cols and s_cols:
            df["Full Name"] = (
                df[f_cols[0]].fillna("").astype(str).str.strip() + " " +
                df[s_cols[0]].fillna("").astype(str).str.strip()
            ).str.strip()
            star_name_col = "Full Name"
        else:
            star_name_col = df.columns[0]

    registered_flags: List[str] = []
    for _, row in df.iterrows():
        p_name = str(row.get(star_name_col, "")).strip()
        f_val = str(row[f_cols[0]]).strip() if f_cols and pd.notna(row.get(f_cols[0])) else ""
        s_val = str(row[s_cols[0]]).strip() if s_cols and pd.notna(row.get(s_cols[0])) else ""

        candidates = [p_name]
        if f_val and s_val:
            candidates.append(f"{f_val} {s_val}")
            candidates.append(f"{s_val} {f_val}")

        expanded_candidates = list(candidates)
        for cand in candidates:
            c_low = cand.lower().strip()
            c_cln = clean_regex(cand)
            if c_low in alias_dict:
                expanded_candidates.append(alias_dict[c_low])
            if c_cln in alias_dict:
                expanded_candidates.append(alias_dict[c_cln])

        is_reg = False
        for cand in expanded_candidates:
            if not cand:
                continue
            cand_low = cand.lower().strip()
            cand_cln = clean_regex(cand)

            if cand_low in raw_reg_names or cand_cln in clean_reg_names:
                is_reg = True
                break

            cand_tokens = get_tokens(cand)
            if not cand_tokens:
                continue

            if frozenset(cand_tokens) in reg_token_sets:
                is_reg = True
                break

            cand_sig = [t for t in cand_tokens if len(t) > 1]
            if cand_sig and frozenset(cand_sig) in reg_sig_token_sets:
                is_reg = True
                break

            if len(cand_tokens) >= 2:
                if (cand_tokens[0], cand_tokens[-1]) in reg_first_last_pairs or (cand_tokens[-1], cand_tokens[0]) in reg_first_last_pairs:
                    is_reg = True
                    break
            if len(cand_sig) >= 2:
                if (cand_sig[0], cand_sig[-1]) in reg_first_last_pairs or (cand_sig[-1], cand_sig[0]) in reg_first_last_pairs:
                    is_reg = True
                    break

        registered_flags.append("✅" if is_reg else "❌")

    df["Registered"] = registered_flags
    return df



# ==============================================================================
# RULE A11 & A12: ABSENCE TRACKER & DEADLINE LOCK
# ==============================================================================

def is_roster_modification_locked(eval_date: Optional[datetime] = None) -> bool:
    """
    Enforces a hard deadline lock blocking standard roster modifications after 31st July per NCU Rule A12.

    Inputs:
        eval_date: Optional datetime to evaluate (defaults to current datetime).

    Outputs:
        bool: True if modifications are locked (after July 31st 23:59:59), False otherwise.

    Helper Apps:
        app.py, starring_rules.py, tests/test_starring_rules.py.
    """
    if eval_date is None:
        eval_date = datetime.now()
    deadline = datetime(eval_date.year, 7, 31, 23, 59, 59)
    return eval_date > deadline


def can_edit_roster(
    eval_date: Optional[datetime] = None,
    admin_override: bool = False
) -> Tuple[bool, str]:
    """
    Determines if roster modifications are permitted under NCU Rule A12,
    supporting Administrative Override for emergency and injury adjustments after 31st July.

    Inputs:
        eval_date: Optional datetime to evaluate (defaults to current datetime).
        admin_override: Boolean flag indicating if administrator enabled emergency/injury override.

    Outputs:
        Tuple[bool, str]:
            - bool: True if edits are allowed, False if locked.
            - str: Status explanation message.

    Helper Apps:
        app.py, starring_rules.py, tests/test_starring_rules.py.
    """
    locked = is_roster_modification_locked(eval_date)
    if not locked:
        return True, "Roster modifications are open (before 31st July deadline)."
    if admin_override:
        return True, "Roster edits permitted via Administrative Override (Emergency / Injury Adjustment)."
    return False, "Standard roster modifications locked after 31st July per NCU Rule A12."


def get_club_registered_players_list(
    df_reg: pd.DataFrame,
    club_name: str,
    clean_club_fn: Optional[Any] = None,
    club_star_df: Optional[pd.DataFrame] = None,
    df_alias: Optional[pd.DataFrame] = None,
    id_map_df: Optional[pd.DataFrame] = None
) -> List[str]:
    """
    Retrieves a deduplicated, sorted list of all registered player names belonging to a specified club.
    Evaluates primary club registrations, transfer destination overrides, and any players
    present in the club starring roster.
    Enforces official NV Play names linked to Sport80 registered names to eliminate duplicates
    (e.g., maps 'Christopher Dempsey' to 'Chris Dempsey').

    Inputs:
        df_reg: Official registration DataFrame (Sport80).
        club_name: Target club name.
        clean_club_fn: Optional club normalization/matching callable.
        club_star_df: Optional club starring DataFrame to ensure complete player coverage.
        df_alias: Optional alias DataFrame linking NV Play names to Sport80 names.
        id_map_df: Optional master ID mapping DataFrame.

    Outputs:
        List[str]: Alphabetically sorted list of registered player official NV Play names (Surname, First Name).

    Helper Apps:
        app.py, tests/test_starring_rules.py.
    """
    # Build Sport80 -> NV Play official mapping from alias master and ID mapping
    sport80_to_nvplay: Dict[str, str] = {}
    if df_alias is None or df_alias.empty:
        for f_cand in ["2. NCU_Validated_Aliases_Master.xlsx", "NCU_Validated_Aliases_Master.xlsx", "12. NCU_Validated_Women's Aliases_Master.xlsx"]:
            if os.path.exists(f_cand):
                try:
                    df_alias = read_excel_calamine(f_cand)
                    break
                except Exception:
                    pass

    if df_alias is not None and not df_alias.empty:
        col_s80 = next((c for c in ["Official Registered Name", "Registered Name"] if c in df_alias.columns), None)
        col_input = "Input Name (Scorecard/Stats)" if "Input Name (Scorecard/Stats)" in df_alias.columns else None
        col_nv_explicit = "NV Play Name" if "NV Play Name" in df_alias.columns else None
        if col_s80:
            for _, r in df_alias.iterrows():
                s80_val = str(r.get(col_s80, "")).strip()
                nv_val = ""
                if col_nv_explicit and pd.notna(r.get(col_nv_explicit)) and str(r.get(col_nv_explicit)).strip() and str(r.get(col_nv_explicit)).strip().lower() not in ["nan", "none"]:
                    nv_val = str(r.get(col_nv_explicit)).strip()
                elif col_input and pd.notna(r.get(col_input)) and str(r.get(col_input)).strip() and str(r.get(col_input)).strip().lower() not in ["nan", "none"]:
                    nv_val = str(r.get(col_input)).strip()

                if s80_val and nv_val and s80_val.lower() not in ["nan", "none"] and nv_val.lower() not in ["nan", "none"]:
                    sport80_to_nvplay[s80_val.lower()] = nv_val

    if id_map_df is None or id_map_df.empty:
        if os.path.exists("NCU_Mens_Master_ID_Mapping.xlsx"):
            try:
                id_map_df = read_excel_calamine("NCU_Mens_Master_ID_Mapping.xlsx")
            except Exception:
                pass

    if id_map_df is not None and not id_map_df.empty:
        if "Sport80_Name" in id_map_df.columns and "NV_Play_Name" in id_map_df.columns:
            for _, r in id_map_df.iterrows():
                s80_val = str(r.get("Sport80_Name", "")).strip()
                nv_val = str(r.get("NV_Play_Name", "")).strip()
                if s80_val and nv_val and s80_val.lower() not in ["nan", "none"] and nv_val.lower() not in ["nan", "none"]:
                    sport80_to_nvplay.setdefault(s80_val.lower(), nv_val)

    def resolve_nv_name(name_str: str) -> str:
        n_clean = name_str.strip().lower()
        if n_clean in sport80_to_nvplay:
            return sport80_to_nvplay[n_clean]
        n_no_punct = n_clean.replace("'", "").replace("-", " ")
        if n_no_punct in sport80_to_nvplay:
            return sport80_to_nvplay[n_no_punct]
        n_spaced = " ".join(n_clean.split())
        if n_spaced in sport80_to_nvplay:
            return sport80_to_nvplay[n_spaced]
        return name_str.strip()

    if df_reg is None or df_reg.empty or not club_name:
        if club_star_df is not None and not club_star_df.empty and "Full Name" in club_star_df.columns:
            starred_names = [resolve_nv_name(str(n).strip()) for n in club_star_df["Full Name"].dropna() if str(n).strip()]
            return sort_players_by_surname(set(starred_names))
        return []

    df = df_reg.copy()

    # 1. Enforce ID Deduplication: Group or filter registered club players by unique Sport80 Registration ID
    reg_id_col = next((c for c in df.columns if any(k in str(c).lower() for k in [
        "registration_id", "registration id", "reg_id", "membership ci",
        "individual membership ci no", "ci no", "ci_no", "member id", "sport80_id", "sport80 id"
    ])), None)
    if reg_id_col:
        valid_id_mask = (
            df[reg_id_col].notna() &
            (df[reg_id_col].astype(str).str.strip() != "") &
            (~df[reg_id_col].astype(str).str.strip().str.lower().isin(["nan", "none"]))
        )
        df_with_id = df[valid_id_mask].drop_duplicates(subset=[reg_id_col], keep="last")
        df_no_id = df[~valid_id_mask]
        df = pd.concat([df_with_id, df_no_id], ignore_index=True)

    # Identify player name column
    name_col = next((c for c in ["Full Name", "Full_Name", "Name", "Player Name", "Player"] if c in df.columns), None)
    if not name_col:
        f_cols = [c for c in df.columns if any(k in str(c).lower() for k in ["first", "forename"])]
        s_cols = [c for c in df.columns if any(k in str(c).lower() for k in ["last", "surname"])]
        if f_cols and s_cols:
            df["Full Name"] = (
                df[f_cols[0]].fillna("").astype(str).str.strip() + " " +
                df[s_cols[0]].fillna("").astype(str).str.strip()
            ).str.strip()
            name_col = "Full Name"
        else:
            name_col = df.columns[0]

    primary_col = next((c for c in df.columns if any(k in str(c).lower() for k in ["primary club", "club"])), None)
    t1_col = next((c for c in df.columns if "transfer club 1" in str(c).lower() or (("transfer club" in str(c).lower() or "transfer" in str(c).lower()) and "2" not in str(c) and "date" not in str(c).lower())), None)
    t2_col = next((c for c in df.columns if "transfer club 2" in str(c).lower() or ("transfer 2" in str(c).lower() and "date" not in str(c).lower())), None)

    import engine as eng
    def is_club_match(c_val: Any) -> bool:
        if not c_val or pd.isna(c_val):
            return False
        if clean_club_fn:
            try:
                res = clean_club_fn(club_name, c_val)
                if isinstance(res, bool):
                    return res
            except Exception:
                pass
            c_norm = str(clean_club_fn(str(c_val)))
            target_norm = str(clean_club_fn(str(club_name)))
            return c_norm.strip().lower() == target_norm.strip().lower()
        try:
            return eng.club_matches_team_base(club_name, str(c_val))
        except Exception:
            return str(club_name).strip().lower() in str(c_val).strip().lower()

    matched_names = set()
    for _, row in df.iterrows():
        p_name = str(row.get(name_col, "")).strip()
        if not p_name or p_name.lower() in ["nan", "none", "unknown", "unknown player"]:
            continue

        prim_val = row.get(primary_col, "") if primary_col else ""
        t1_val = row.get(t1_col, "") if t1_col else ""
        t2_val = row.get(t2_col, "") if t2_col else ""

        # Current effective club: latest transfer if exists, else primary
        effective_club = prim_val
        if pd.notna(t2_val) and str(t2_val).strip() and str(t2_val).strip().lower() not in ["nan", "none"]:
            effective_club = t2_val
        elif pd.notna(t1_val) and str(t1_val).strip() and str(t1_val).strip().lower() not in ["nan", "none"]:
            effective_club = t1_val

        if is_club_match(effective_club):
            # Normalize to official NV Play name linked to Sport80 name
            matched_names.add(resolve_nv_name(p_name))

    # Also include any currently starred players for this club to prevent omissions
    if club_star_df is not None and not club_star_df.empty and "Full Name" in club_star_df.columns:
        starred_names = [str(n).strip() for n in club_star_df["Full Name"].dropna() if str(n).strip()]
        for sn in starred_names:
            if sn.lower() not in ["nan", "none", ""]:
                matched_names.add(resolve_nv_name(sn))

    # Final deduplication pass ensuring all entries are resolved NV Play names with club/team suffixes stripped
    deduped_names = {strip_club_suffix(resolve_nv_name(n), club_name) for n in matched_names if n and n.lower() not in ["nan", "none", ""]}

    return sort_players_by_surname(deduped_names)


def player_surname_sort_key(name: str) -> Tuple[str, str]:
    """
    Sort key that orders player names alphabetically by surname first, then forename(s).
    e.g. 'Mark Adair' -> ('adair', 'mark')
         'Ross Adair' -> ('adair', 'ross')
         'Paul Stirling' -> ('stirling', 'paul')

    Inputs:
        name: Player name string.

    Outputs:
        Tuple[str, str]: (surname_lowercase, forename_lowercase).

    Helper Apps:
        app.py, secretary_app.py, starring_rules.py, tests/test_starring_predictor.py.
    """
    parts = str(name).strip().split()
    if not parts:
        return ("", "")
    surname = parts[-1].lower()
    forename = " ".join(parts[:-1]).lower() if len(parts) > 1 else ""
    return (surname, forename)


def sort_players_by_surname(names: Iterable[str]) -> List[str]:
    """
    Sorts an iterable of player names alphabetically by surname first, then forename(s).

    Inputs:
        names: Iterable of player name strings.

    Outputs:
        List[str]: List of sorted player names.

    Helper Apps:
        app.py, secretary_app.py, starring_rules.py, tests/test_starring_predictor.py.
    """
    return sorted(list(names), key=player_surname_sort_key)


def log_starring_override_change(
    club: str,
    tier: str,
    player_out: str,
    player_in: str,
    admin_comment: str,
    prev_tier: Optional[str] = None,
    file_path: str = "NCU_Club_Starring_History.xlsx",
    timestamp: Optional[datetime] = None
) -> Dict[str, Any]:
    """
    Logs an administrative roster override modification into NCU_Club_Starring_History.xlsx.
    Enforces mandatory explanation comments prefixed with '[ADMIN OVERRIDE]' and applies
    strict GEMINI.md styling (navy headers, A2 freeze panes, auto column widths).
    If the replacement player is being promoted from a lower tier, records the previous tier.

    Inputs:
        club: Club name.
        tier: Team tier (e.g. '1st XI', '2nd XI').
        player_out: Name of outgoing or de-starred player.
        player_in: Name of incoming replacement player.
        admin_comment: Mandatory explanation comment.
        prev_tier: Optional previous starring tier if promoted from lower XI (e.g. '2nd XI').
        file_path: Destination Excel workbook path (default 'NCU_Club_Starring_History.xlsx').
        timestamp: Optional datetime of adjustment (defaults to datetime.now()).

    Outputs:
        Dict[str, Any]: Saved audit record dictionary.

    Helper Apps:
        app.py, tests/test_starring_rules.py.
    """
    if not admin_comment or not str(admin_comment).strip():
        raise ValueError("Mandatory comment / explanation required for post-July 31st roster modification.")

    clean_comment = str(admin_comment).strip()
    if prev_tier and f"[PROMOTED FROM {prev_tier.upper()}]" not in clean_comment.upper():
        clean_comment = f"[PROMOTED FROM {prev_tier}] {clean_comment}"
    if not clean_comment.startswith("[ADMIN OVERRIDE]"):
        clean_comment = f"[ADMIN OVERRIDE] {clean_comment}"

    if timestamp is None:
        timestamp = datetime.now()

    new_record = {
        "Timestamp": timestamp.strftime("%Y-%m-%d %H:%M:%S"),
        "Club": str(club).strip(),
        "Tier": str(tier).strip(),
        "Player Out": str(player_out).strip(),
        "Player In": str(player_in).strip(),
        "Admin Comment": clean_comment
    }

    import openpyxl
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    headers = ["Timestamp", "Club", "Tier", "Player Out", "Player In", "Admin Comment"]
    existing_records: List[Dict[str, Any]] = []

    wb = None
    if os.path.exists(file_path):
        try:
            wb = openpyxl.load_workbook(file_path)
            if "Starring Audit Trail" in wb.sheetnames:
                ws_exist = wb["Starring Audit Trail"]
            else:
                ws_exist = wb.active
            exist_headers = [ws_exist.cell(row=1, column=c).value for c in range(1, ws_exist.max_column + 1)]
            for r in range(2, ws_exist.max_row + 1):
                row_vals = [ws_exist.cell(row=r, column=c).value for c in range(1, ws_exist.max_column + 1)]
                if any(v is not None for v in row_vals):
                    r_dict = {str(exist_headers[i]): row_vals[i] for i in range(min(len(exist_headers), len(row_vals)))}
                    existing_records.append(r_dict)
        except Exception:
            existing_records = []

    existing_records.append(new_record)

    dir_name = os.path.dirname(file_path)
    if dir_name:
        os.makedirs(dir_name, exist_ok=True)

    if wb is None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Starring Audit Trail"
    elif "Starring Audit Trail" in wb.sheetnames:
        ws = wb["Starring Audit Trail"]
        ws.delete_rows(1, ws.max_row + 1)
    else:
        ws = wb.create_sheet("Starring Audit Trail")

    ws.append(headers)
    for r in existing_records:
        ws.append([r.get(h, "") for h in headers])

    # Format per GEMINI.md:
    # 1. First row navy fill (#1F4E78), bold white font
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font

    # 2. Freeze panes at row 2
    ws.freeze_panes = "A2"

    # 3. Column widths auto-fitted max(len + 2, 10)
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 2, 10)

    wb.save(file_path)
    return new_record


def save_international_exemptions(
    club_name: str,
    exempt_players: List[str],
    domain: str = "Men's",
    file_path: str = "NCU_Club_Starring_History.xlsx",
    timestamp: Optional[datetime] = None
) -> List[Dict[str, Any]]:
    """
    Saves and persists player names granted an Irish International Duty exemption for a specific club.
    Records data into a dedicated 'International Exemptions' worksheet in NCU_Club_Starring_History.xlsx.
    Maintains GEMINI.md styling (navy header fill, bold white text, A2 freeze panes, auto column widths).

    Inputs:
        club_name: Name of the cricket club.
        exempt_players: List of player names granted international exemption.
        domain: Cricket domain ('Men\\'s' or 'Women\\'s').
        file_path: Excel workbook path (default 'NCU_Club_Starring_History.xlsx').
        timestamp: Optional timestamp of update (defaults to datetime.now()).

    Outputs:
        List[Dict[str, Any]]: Saved exemption records for the club.

    Helper Apps:
        app.py, starring_rules.py, tests/test_starring_rules.py.
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    if timestamp is None:
        timestamp = datetime.now()
    time_str = timestamp.strftime("%Y-%m-%d %H:%M:%S")

    clean_club = str(club_name).strip()
    clean_domain = str(domain).strip()

    headers = ["Domain", "Club", "Player Name", "Status", "Updated At"]
    existing_records: List[Dict[str, Any]] = []

    dir_name = os.path.dirname(file_path)
    if dir_name:
        os.makedirs(dir_name, exist_ok=True)

    wb = None
    if os.path.exists(file_path):
        try:
            wb = openpyxl.load_workbook(file_path)
            if "International Exemptions" in wb.sheetnames:
                ws_intl = wb["International Exemptions"]
                exist_headers = [ws_intl.cell(row=1, column=c).value for c in range(1, ws_intl.max_column + 1)]
                for r in range(2, ws_intl.max_row + 1):
                    row_vals = [ws_intl.cell(row=r, column=c).value for c in range(1, ws_intl.max_column + 1)]
                    if any(v is not None for v in row_vals):
                        r_dict = {str(exist_headers[i]): row_vals[i] for i in range(min(len(exist_headers), len(row_vals)))}
                        r_dom = str(r_dict.get("Domain", "")).strip().lower()
                        r_clb = str(r_dict.get("Club", "")).strip().lower()
                        if not (r_dom == clean_domain.lower() and r_clb == clean_club.lower()):
                            existing_records.append(r_dict)
        except Exception:
            existing_records = []

    if wb is None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "International Exemptions"
    elif "International Exemptions" in wb.sheetnames:
        ws = wb["International Exemptions"]
        ws.delete_rows(1, ws.max_row + 1)
    else:
        ws = wb.create_sheet("International Exemptions")

    new_records: List[Dict[str, Any]] = []
    clean_players = [str(p).strip() for p in exempt_players if str(p).strip()]
    if clean_players:
        for p in clean_players:
            rec = {
                "Domain": clean_domain,
                "Club": clean_club,
                "Player Name": p,
                "Status": "Active Exemption",
                "Updated At": time_str
            }
            new_records.append(rec)
            existing_records.append(rec)
    else:
        rec = {
            "Domain": clean_domain,
            "Club": clean_club,
            "Player Name": "[None]",
            "Status": "No Exemptions",
            "Updated At": time_str
        }
        new_records.append(rec)
        existing_records.append(rec)

    ws.append(headers)
    for r in existing_records:
        ws.append([r.get(h, "") for h in headers])

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font

    ws.freeze_panes = "A2"

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 2, 10)

    wb.save(file_path)
    return new_records


def load_international_exemptions(
    club_name: str,
    domain: str = "Men's",
    file_path: str = "NCU_Club_Starring_History.xlsx"
) -> Optional[List[str]]:
    """
    Loads saved international duty exemption player names for a specific club from NCU_Club_Starring_History.xlsx.
    Returns None if the club has never been configured in the worksheet.
    Returns an empty list [] if the club was explicitly saved with no exemptions ([None]).

    Inputs:
        club_name: Name of the cricket club.
        domain: Cricket domain ('Men\\'s' or 'Women\\'s').
        file_path: Excel workbook path (default 'NCU_Club_Starring_History.xlsx').

    Outputs:
        Optional[List[str]]: List of player names, or None if never configured.

    Helper Apps:
        app.py, starring_rules.py, tests/test_starring_rules.py.
    """
    if not os.path.exists(file_path):
        return None

    import openpyxl
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        if "International Exemptions" not in wb.sheetnames:
            return None
        ws = wb["International Exemptions"]
        headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        headers_lower = [str(h).strip().lower() if h else "" for h in headers]

        if "domain" not in headers_lower or "club" not in headers_lower or "player name" not in headers_lower:
            return None

        idx_dom = headers_lower.index("domain") + 1
        idx_clb = headers_lower.index("club") + 1
        idx_p = headers_lower.index("player name") + 1

        clean_club = str(club_name).strip().lower()
        clean_domain = str(domain).strip().lower()

        found_club_rows = False
        exempt_list = []

        for r in range(2, ws.max_row + 1):
            dom_val = str(ws.cell(row=r, column=idx_dom).value or "").strip().lower()
            clb_val = str(ws.cell(row=r, column=idx_clb).value or "").strip().lower()
            if dom_val == clean_domain and clb_val == clean_club:
                found_club_rows = True
                p_val = str(ws.cell(row=r, column=idx_p).value or "").strip()
                if p_val and p_val != "[None]":
                    exempt_list.append(p_val)

        if found_club_rows:
            return exempt_list
        return None
    except Exception:
        return None


def load_board_dispensations(
    club_name: Optional[str] = None,
    season: Optional[int] = 2027,
    file_path: str = "NCU_Club_Starring_History.xlsx"
) -> pd.DataFrame:
    """
    Loads Board-Approved Executive Availability Dispensations from NCU_Club_Starring_History.xlsx.
    If the 'Board_Dispensations' worksheet is missing, automatically creates and seeds it
    with the default approved exception:
    Simon King | Larne | 2nd XI | "Cannot play Saturdays - Sunday availability only" | 2027.
    Enforces GEMINI.md styling (navy fill, bold white text, A2 freeze panes, dynamic column widths).

    Inputs:
        club_name: Optional club name to filter results (e.g. 'Larne').
        season: Optional season to filter results (e.g. 2027).
        file_path: Excel workbook path (default 'NCU_Club_Starring_History.xlsx').

    Outputs:
        pd.DataFrame: Table of active executive availability dispensations.

    Helper Apps:
        app.py, starring_rules.py, engine.py, tests/test_starring_rules.py.
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    headers = ["Player Name", "Club Name", "Allowed Lower Tier", "Dispensation Reason", "Effective Season"]

    dir_name = os.path.dirname(file_path)
    if dir_name:
        os.makedirs(dir_name, exist_ok=True)

    wb = None
    needs_seeding = False

    if os.path.exists(file_path):
        try:
            wb = openpyxl.load_workbook(file_path)
            if "Board_Dispensations" not in wb.sheetnames:
                needs_seeding = True
        except Exception:
            wb = None
            needs_seeding = True
    else:
        needs_seeding = True

    if wb is None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Board_Dispensations"
    elif "Board_Dispensations" in wb.sheetnames:
        ws = wb["Board_Dispensations"]
    else:
        ws = wb.create_sheet("Board_Dispensations")

    if needs_seeding or (ws.max_row <= 1 and ws.cell(row=1, column=1).value != "Player Name"):
        ws.delete_rows(1, max(ws.max_row + 1, 10))
        ws.append(headers)
        # Default seed entry
        ws.append([
            "Simon King",
            "Larne",
            "2nd XI",
            "Cannot play Saturdays - Sunday availability only",
            2027
        ])

        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font

        ws.freeze_panes = "A2"

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = max(max_len + 2, 10)

        wb.save(file_path)

    records: List[Dict[str, Any]] = []
    exist_headers = [str(ws.cell(row=1, column=c).value or "").strip() for c in range(1, ws.max_column + 1)]
    for r in range(2, ws.max_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if any(v is not None for v in row_vals):
            rec = {exist_headers[i]: row_vals[i] for i in range(min(len(exist_headers), len(row_vals)))}
            p_name = str(rec.get("Player Name", "")).strip()
            if p_name and p_name != "None":
                try:
                    eff_season = int(rec.get("Effective Season", 2027))
                except (ValueError, TypeError):
                    eff_season = 2027
                records.append({
                    "Player Name": p_name,
                    "Club Name": str(rec.get("Club Name", "")).strip(),
                    "Allowed Lower Tier": str(rec.get("Allowed Lower Tier", "2nd XI")).strip(),
                    "Dispensation Reason": str(rec.get("Dispensation Reason", "")).strip(),
                    "Effective Season": eff_season
                })

    try:
        wb.close()
    except Exception:
        pass

    df = pd.DataFrame(records)
    if df.empty:
        df = pd.DataFrame(columns=headers)

    if club_name:
        c_clean = club_name.strip().lower()
        df = df[df["Club Name"].astype(str).str.strip().str.lower() == c_clean]

    if season is not None:
        s_val = int(season)
        df = df[df["Effective Season"] == s_val]

    return df.reset_index(drop=True)


def save_board_dispensation(
    player_name: str,
    club_name: str,
    allowed_lower_tier: str,
    dispensation_reason: str,
    effective_season: int = 2027,
    file_path: str = "NCU_Club_Starring_History.xlsx"
) -> Dict[str, Any]:
    """
    Appends or updates an individual Board-Approved Executive Availability Dispensation
    in the 'Board_Dispensations' worksheet of NCU_Club_Starring_History.xlsx.
    Maintains GEMINI.md styling (navy fill, bold white text, A2 freeze panes, dynamic column widths).

    Inputs:
        player_name: Player receiving the dispensation.
        club_name: Player's registered club.
        allowed_lower_tier: Approved lower playing level (e.g. '2nd XI').
        dispensation_reason: Formal rationale (e.g. 'Cannot play Saturdays - Sunday availability only').
        effective_season: Year of applicability (default 2027).
        file_path: Target Excel file.

    Outputs:
        Dict[str, Any]: Saved record details.

    Helper Apps:
        app.py, starring_rules.py, engine.py, tests/test_starring_rules.py.
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    p_clean = str(player_name).strip()
    c_clean = str(club_name).strip()
    tier_clean = str(allowed_lower_tier).strip()
    reason_clean = str(dispensation_reason).strip()
    season_val = int(effective_season)

    # Ensure sheet exists by loading
    _ = load_board_dispensations(file_path=file_path)

    wb = openpyxl.load_workbook(file_path)
    ws = wb["Board_Dispensations"]

    headers = ["Player Name", "Club Name", "Allowed Lower Tier", "Dispensation Reason", "Effective Season"]

    records: List[Dict[str, Any]] = []
    exist_headers = [str(ws.cell(row=1, column=c).value or "").strip() for c in range(1, ws.max_column + 1)]
    for r in range(2, ws.max_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if any(v is not None for v in row_vals):
            rec = {exist_headers[i]: row_vals[i] for i in range(min(len(exist_headers), len(row_vals)))}
            p_val = str(rec.get("Player Name", "")).strip()
            if p_val and p_val != "None":
                records.append(rec)

    matched = False
    new_rec = {
        "Player Name": p_clean,
        "Club Name": c_clean,
        "Allowed Lower Tier": tier_clean,
        "Dispensation Reason": reason_clean,
        "Effective Season": season_val
    }

    for r in records:
        if (
            str(r.get("Player Name", "")).strip().lower() == p_clean.lower()
            and str(r.get("Club Name", "")).strip().lower() == c_clean.lower()
            and int(r.get("Effective Season", 2027)) == season_val
        ):
            r["Allowed Lower Tier"] = tier_clean
            r["Dispensation Reason"] = reason_clean
            matched = True
            break

    if not matched:
        records.append(new_rec)

    ws.delete_rows(1, max(ws.max_row + 1, 10))
    ws.append(headers)
    for r in records:
        ws.append([
            r.get("Player Name", ""),
            r.get("Club Name", ""),
            r.get("Allowed Lower Tier", ""),
            r.get("Dispensation Reason", ""),
            int(r.get("Effective Season", 2027))
        ])

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font

    ws.freeze_panes = "A2"

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 2, 10)

    wb.save(file_path)
    return new_rec


save_board_dispensations = save_board_dispensation


def delete_board_dispensation(
    player_name: str,
    club_name: Optional[str] = None,
    effective_season: Optional[int] = 2027,
    file_path: str = "NCU_Club_Starring_History.xlsx"
) -> bool:
    """
    Removes an individual Board-Approved Executive Availability Dispensation record
    from the 'Board_Dispensations' worksheet in NCU_Club_Starring_History.xlsx.
    Maintains GEMINI.md styling (navy fill, bold white text, A2 freeze panes, dynamic column widths).

    Inputs:
        player_name: Player to remove from dispensations.
        club_name: Optional club name to match.
        effective_season: Optional season year (default 2027).
        file_path: Target Excel file.

    Outputs:
        bool: True if an entry was matched and deleted, False otherwise.

    Helper Apps:
        app.py, starring_rules.py, engine.py, tests/test_starring_rules.py.
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    p_clean = str(player_name).strip().lower()
    c_clean = str(club_name).strip().lower() if club_name else None

    if not os.path.exists(file_path):
        return False

    wb = openpyxl.load_workbook(file_path)
    if "Board_Dispensations" not in wb.sheetnames:
        wb.close()
        return False

    ws = wb["Board_Dispensations"]
    headers = ["Player Name", "Club Name", "Allowed Lower Tier", "Dispensation Reason", "Effective Season"]

    records: List[Dict[str, Any]] = []
    exist_headers = [str(ws.cell(row=1, column=c).value or "").strip() for c in range(1, ws.max_column + 1)]
    removed = False

    for r in range(2, ws.max_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if any(v is not None for v in row_vals):
            rec = {exist_headers[i]: row_vals[i] for i in range(min(len(exist_headers), len(row_vals)))}
            p_val = str(rec.get("Player Name", "")).strip().lower()
            c_val = str(rec.get("Club Name", "")).strip().lower()
            try:
                s_val = int(rec.get("Effective Season", 2027))
            except Exception:
                s_val = 2027

            match_player = (p_val == p_clean)
            match_club = (c_clean is None or c_val == c_clean)
            match_season = (effective_season is None or s_val == int(effective_season))

            if match_player and match_club and match_season:
                removed = True
                continue
            if p_val and p_val != "none":
                records.append(rec)

    if not removed:
        wb.close()
        return False

    ws.delete_rows(1, max(ws.max_row + 1, 10))
    ws.append(headers)
    for r in records:
        ws.append([
            r.get("Player Name", ""),
            r.get("Club Name", ""),
            r.get("Allowed Lower Tier", ""),
            r.get("Dispensation Reason", ""),
            int(r.get("Effective Season", 2027))
        ])

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font

    ws.freeze_panes = "A2"

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 2, 10)

    wb.save(file_path)
    wb.close()
    return True


delete_board_dispensations = delete_board_dispensation



def get_club_starred_roster_sorted(club_star_df: pd.DataFrame) -> List[str]:
    """
    Extracts and returns only the currently starred roster player names for a club,
    sorted alphabetically by Surname, then Forename.

    Inputs:
        club_star_df: DataFrame of starred players for a club.

    Outputs:
        List[str]: List of player names sorted alphabetically (Surname, Forename).

    Helper Apps:
        app.py, tests/test_starring_rules.py.
    """
    if club_star_df is None or club_star_df.empty:
        return []

    sorted_df = club_star_df.copy()
    sort_cols = [c for c in ["Surname", "Forename"] if c in sorted_df.columns]
    if sort_cols:
        sorted_df = sorted_df.sort_values(by=sort_cols, ascending=True)

    result: List[str] = []
    for _, row in sorted_df.iterrows():
        p_name = strip_club_suffix(str(row.get("Full Name", "")).strip())
        if not p_name or p_name.lower() == "nan":
            fn = str(row.get("Forename", "")).strip()
            sn = str(row.get("Surname", "")).strip()
            p_name = strip_club_suffix(f"{fn} {sn}".strip())
        if p_name and p_name not in result:
            result.append(p_name)
    return result


def update_master_starring_roster(
    file_path: str,
    club_name: str,
    tier: str,
    player_out: str,
    player_in: str,
    prev_tier: Optional[str] = None
) -> bool:
    """
    Updates a club's starring roster directly in the master Excel workbook.
    Replaces the outgoing player's surname and forename with the incoming replacement player
    under the specified XI tier. If the replacement player was promoted from a lower tier,
    clears their previous tier slot to prevent duplicate starring.

    Inputs:
        file_path: Path to the master starring Excel workbook (e.g. '3. NCU Complete -Men's- Starring List from 1st June.xlsx').
        club_name: Name of the club (matching the worksheet name).
        tier: Target team tier (e.g. '1st XI', '2nd XI').
        player_out: Outgoing player full name.
        player_in: Incoming replacement player full name.
        prev_tier: Optional previous tier if player_in was promoted from a lower tier.

    Outputs:
        bool: True if successfully updated and saved, False otherwise.

    Helper Apps:
        app.py, starring_rules.py, tests/test_starring_rules.py.
    """
    if not file_path or not os.path.exists(file_path):
        return False

    import openpyxl

    wb = openpyxl.load_workbook(file_path)
    sheet_name = next((s for s in wb.sheetnames if s.strip().lower() == club_name.strip().lower()), None)
    if not sheet_name:
        return False

    ws = wb[sheet_name]

    out_parts = player_out.strip().split()
    out_last = out_parts[-1].lower() if out_parts else ""
    out_first = " ".join(out_parts[:-1]).lower() if len(out_parts) > 1 else ""

    in_parts = player_in.strip().split()
    in_last = in_parts[-1] if in_parts else ""
    in_first = " ".join(in_parts[:-1]) if len(in_parts) > 1 else ""

    is_vacant_target = any(k in player_out.lower() for k in ["vacant", "empty"])

    current_tier = None
    target_row = None
    last_tier_row = None

    for r in range(1, ws.max_row + 1):
        c1_val = ws.cell(row=r, column=1).value
        col1 = str(c1_val or "").strip()
        if "XI" in col1:
            current_tier = col1
            continue

        if current_tier and tier.lower() in current_tier.lower():
            s_val = str(ws.cell(row=r, column=2).value or "").strip()
            f_val = str(ws.cell(row=r, column=5).value or "").strip()

            is_player_slot = False
            try:
                if c1_val is not None and str(c1_val).strip().isdigit():
                    is_player_slot = True
                    last_tier_row = r
            except Exception:
                pass

            if is_vacant_target:
                if is_player_slot and (not s_val or s_val.lower() in ["none", "nan", "[vacant]", "vacant", ""]):
                    target_row = r
                    break
            else:
                if s_val and s_val.lower() not in ["surname"]:
                    full_cell = f"{f_val} {s_val}".strip().lower()
                    if (full_cell == player_out.strip().lower() or
                        (s_val.lower() == out_last and (not out_first or out_first in f_val.lower() or f_val.lower() in out_first))):
                        target_row = r
                        break

    if is_vacant_target and not target_row and last_tier_row is not None:
        next_c1 = str(ws.cell(row=last_tier_row + 1, column=1).value or "").strip()
        next_c2 = str(ws.cell(row=last_tier_row + 1, column=2).value or "").strip()
        if "XI" in next_c1 or next_c2.lower() == "surname":
            ws.insert_rows(last_tier_row + 1)
        target_row = last_tier_row + 1
        prev_num = ws.cell(row=last_tier_row, column=1).value
        try:
            ws.cell(row=target_row, column=1).value = int(prev_num) + 1 if prev_num else 1
        except Exception:
            pass

    if not target_row:
        return False

    # Ensure target slot has a rank counter if empty
    if ws.cell(row=target_row, column=1).value is None and last_tier_row is not None:
        prev_num = ws.cell(row=last_tier_row, column=1).value
        try:
            ws.cell(row=target_row, column=1).value = int(prev_num) + 1 if prev_num else 1
        except Exception:
            pass

    # Update outgoing player with incoming player
    ws.cell(row=target_row, column=2).value = in_last
    ws.cell(row=target_row, column=5).value = in_first

    # If promoted from a lower tier, clear their old slot in prev_tier so they are not duplicated
    if prev_tier:
        prev_current_tier = None
        for r in range(1, ws.max_row + 1):
            col1 = str(ws.cell(row=r, column=1).value or "").strip()
            if "XI" in col1:
                prev_current_tier = col1
                continue

            if prev_current_tier and prev_tier.lower() in prev_current_tier.lower():
                s_val = str(ws.cell(row=r, column=2).value or "").strip()
                f_val = str(ws.cell(row=r, column=5).value or "").strip()
                if s_val and s_val.lower() not in ["surname"]:
                    full_cell = f"{f_val} {s_val}".strip().lower()
                    if (full_cell == player_in.strip().lower() or
                        (s_val.lower() == in_last.lower() and (not in_first or in_first.lower() in f_val.lower() or f_val.lower() in in_first.lower()))):
                        ws.cell(row=r, column=2).value = None
                        ws.cell(row=r, column=5).value = None
                        break

    # Format per GEMINI.md:
    ws.freeze_panes = "A2"
    from openpyxl.utils import get_column_letter
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 2, 10)

    wb.save(file_path)
    return True


def check_player_absence(
    starred_team: str,
    player_name: str,
    appearances: List[Dict[str, Any]],
    club_team_matches: List[Dict[str, Any]],
    is_international: bool = False,
    evaluation_date: Optional[datetime] = None
) -> Dict[str, Any]:
    """
    Evaluates player eligibility and absence against NCU Rules A11 & A12.
    Flags an administrative alert ("⚠️ De-Starring Action Required") if a starred player
    does not appear for their team or a higher one for 3 consecutive matches or 3 weeks
    (whichever is greater). Overridden if 'Irish International Duty' flag is toggled.

    Inputs:
        starred_team: The team tier the player is starred for (e.g. '1st XI', '2nd XI').
        player_name: Full name of the player.
        appearances: List of match records where player appeared, each dict having:
            - 'Match_Date': datetime
            - 'Team': str (team player played for)
        club_team_matches: List of all matches played by the club's starred team or higher, each dict having:
            - 'Match_Date': datetime
            - 'Team': str
        is_international: Boolean toggle for Irish International Duty exemption.
        evaluation_date: Datetime of check (defaults to datetime.now()).

    Outputs:
        Dict[str, Any]: Absence status dictionary containing:
            - 'player_name': str
            - 'starred_team': str
            - 'is_international': bool
            - 'destarring_required': bool
            - 'alert_message': str ('⚠️ De-Starring Action Required' or 'Eligible' or 'International Duty Exemption 🇮🇪')
            - 'matches_missed': int
            - 'days_inactive': int
            - 'last_played_date': Optional[datetime]
            - 'eligible_appearances_count': int

    Helper Apps:
        app.py, engine.py, tests/test_starring_rules.py.
    """
    if evaluation_date is None:
        evaluation_date = datetime.now()

    if is_international:
        return {
            "player_name": player_name,
            "starred_team": starred_team,
            "is_international": True,
            "destarring_required": False,
            "alert_message": "International Duty Exemption 🇮🇪",
            "matches_missed": 0,
            "days_inactive": 0,
            "last_played_date": None,
            "eligible_appearances_count": len(appearances)
        }

    starred_rank = get_tier_level_rank(starred_team)

    # Filter appearances to those in the player's starred team or any HIGHER team
    eligible_apps = []
    for app in appearances:
        app_team = app.get("Team", "")
        app_rank = get_tier_level_rank(app_team)
        # Lower rank number = higher or equal tier (e.g. rank 1 <= rank 2)
        if app_rank <= starred_rank:
            eligible_apps.append(app)

    # Sort eligible appearances chronologically
    eligible_apps.sort(key=lambda x: x.get("Match_Date", datetime.min))
    last_played_date = eligible_apps[-1].get("Match_Date") if eligible_apps else None

    # Filter club matches to matches played by the player's starred team or higher
    relevant_club_matches = [
        m for m in club_team_matches
        if get_tier_level_rank(m.get("Team", "")) <= starred_rank
    ]
    relevant_club_matches.sort(key=lambda x: x.get("Match_Date", datetime.min))

    if last_played_date is not None:
        days_inactive = (evaluation_date.date() - last_played_date.date()).days if hasattr(last_played_date, 'date') else (evaluation_date - last_played_date).days
        days_inactive = max(0, days_inactive)
        # Missed matches played strictly after last appearance date
        matches_missed = sum(
            1 for m in relevant_club_matches
            if m.get("Match_Date") and m.get("Match_Date") > last_played_date
        )
    else:
        # Zero eligible appearances all season
        if relevant_club_matches:
            first_match_date = relevant_club_matches[0].get("Match_Date")
            days_inactive = (evaluation_date.date() - first_match_date.date()).days if hasattr(first_match_date, 'date') else (evaluation_date - first_match_date).days
            days_inactive = max(0, days_inactive)
        else:
            days_inactive = 0
        matches_missed = len(relevant_club_matches)

    # Rule A11/A12 Condition: "3 consecutive matches or 3 weeks (whichever is greater)"
    # 3 weeks = 21 days. Both constraints must have elapsed for the threshold to be breached.
    destarring_required = (days_inactive >= 21 and matches_missed >= 3)

    alert_message = "⚠️ De-Starring Action Required" if destarring_required else "Eligible"

    return {
        "player_name": player_name,
        "starred_team": starred_team,
        "is_international": False,
        "destarring_required": destarring_required,
        "alert_message": alert_message,
        "matches_missed": matches_missed,
        "days_inactive": days_inactive,
        "last_played_date": last_played_date,
        "eligible_appearances_count": len(eligible_apps)
    }


# ==============================================================================
# RULE A13: TRANSFER MONITOR & FINANCE LINK
# ==============================================================================

def evaluate_player_transfers(
    df_reg: pd.DataFrame,
    season_year: int = 2026,
    clean_club_fn: Optional[Any] = None
) -> Dict[str, Any]:
    """
    Evaluates player transfers against NCU Rule A13:
    1. Explicitly blocks any player from transferring more than twice per season.
    2. Calculates a £25.00 Transfer Fee infraction for any valid transfer occurring
       on or after 1st April of the season year.
    3. Aggregates fee obligations by destination club for piping into finance_app.py.

    Inputs:
        df_reg: Registration records DataFrame with transfer columns.
        season_year: Cricket season year (default 2026).
        clean_club_fn: Optional club cleaning function (normalizes club names).

    Outputs:
        Dict[str, Any]:
            - 'valid_transfers': pd.DataFrame of all compliant transfers with fee calculations
            - 'blocked_transfers': pd.DataFrame of transfers exceeding the 2-transfer ceiling
            - 'club_summary': Dict[str, Dict[str, Any]] mapping club name to transfer counts and fees
            - 'total_fees_assessed': float total fee sum across all clubs

    Helper Apps:
        finance_app.py, app.py, tests/test_starring_rules.py.
    """
    def default_clean_club(c: Any) -> str:
        if not c or pd.isna(c):
            return "Unknown"
        c_str = str(c).strip()
        if c_str == "" or c_str.lower() in ["nan", "none"]:
            return "Unknown"
        # Strip parenthesized annotations and common club suffixes
        c_clean = re.sub(r'\(.*?\)', '', c_str)
        c_clean = re.sub(r'(?i)\bcricket\s+club\b|\bcc\b', '', c_clean).strip()
        return c_clean if c_clean else "Unknown"

    if clean_club_fn is None:
        clean_club_fn = default_clean_club

    valid_records = []
    blocked_records = []
    april_first = datetime(season_year, 4, 1)

    if df_reg is None or df_reg.empty:
        return {
            "valid_transfers": pd.DataFrame(),
            "blocked_transfers": pd.DataFrame(),
            "club_summary": {},
            "total_fees_assessed": 0.0
        }

    df = df_reg.copy()
    f_cols = [c for c in df.columns if any(k in str(c).lower() for k in ["first", "forename"])]
    s_cols = [c for c in df.columns if any(k in str(c).lower() for k in ["last", "surname"])]

    name_col = next((c for c in ["Full Name", "Full_Name", "Name", "Player Name", "Player"] if c in df.columns), None)
    if not name_col:
        if f_cols and s_cols:
            df["Full Name"] = (
                df[f_cols[0]].fillna("").astype(str).str.strip() + " " +
                df[s_cols[0]].fillna("").astype(str).str.strip()
            ).str.strip()
            name_col = "Full Name"
        else:
            name_col = df.columns[0]

    primary_club_col = next(
        (c for c in df.columns if any(k in str(c).lower() for k in ["primary club", "club"])),
        None
    )

    t1_club_col = next((c for c in df.columns if "transfer club 1" in str(c).lower() or (("transfer club" in str(c).lower() or "transfer" in str(c).lower()) and "2" not in str(c) and "date" not in str(c).lower())), None)
    t1_date_col = next((c for c in df.columns if "transfer date 1" in str(c).lower() or ("transfer date" in str(c).lower() and "2" not in str(c))), None)

    t2_club_col = next((c for c in df.columns if "transfer club 2" in str(c).lower() or ("transfer 2" in str(c).lower() and "date" not in str(c).lower())), None)
    t2_date_col = next((c for c in df.columns if "transfer date 2" in str(c).lower()), None)

    t3_club_col = next((c for c in df.columns if "transfer club 3" in str(c).lower() or "transfer 3" in str(c).lower()), None)
    t3_date_col = next((c for c in df.columns if "transfer date 3" in str(c).lower()), None)

    def parse_transfer_date(raw_date: Any) -> Optional[datetime]:
        if pd.isna(raw_date) or not raw_date:
            return None
        if isinstance(raw_date, (pd.Timestamp, datetime)):
            return raw_date.to_pydatetime() if isinstance(raw_date, pd.Timestamp) else raw_date
        s = str(raw_date).strip()
        if not s or s.lower() in ["nan", "nat", "none", ""]:
            return None
        if re.match(r"^\d{4}-\d{2}-\d{2}", s):
            ts = pd.to_datetime(s, errors="coerce", dayfirst=False)
        else:
            ts = pd.to_datetime(s, errors="coerce", dayfirst=True)
        return ts.to_pydatetime() if pd.notna(ts) else None

    for _, row in df.iterrows():
        p_name = str(row.get(name_col, "Unknown Player")).strip()
        orig_club = clean_club_fn(row.get(primary_club_col, "Unknown")) if primary_club_col else "Unknown"

        p_surname = ""
        p_firstname = ""
        if s_cols and pd.notna(row.get(s_cols[0])):
            p_surname = str(row.get(s_cols[0])).strip()
        if f_cols and pd.notna(row.get(f_cols[0])):
            p_firstname = str(row.get(f_cols[0])).strip()
        if not p_surname and p_name:
            tokens = p_name.split()
            p_surname = tokens[-1] if tokens else ""
            p_firstname = tokens[0] if len(tokens) > 1 else ""

        player_transfers = []

        # Transfer 1
        if t1_club_col and pd.notna(row.get(t1_club_col)) and str(row.get(t1_club_col)).strip():
            raw_t1_club = str(row.get(t1_club_col)).strip()
            raw_t1_date = row.get(t1_date_col) if t1_date_col else None
            t1_date = parse_transfer_date(raw_t1_date)
            player_transfers.append({
                "transfer_num": 1,
                "from_club": orig_club,
                "to_club": clean_club_fn(raw_t1_club),
                "date": t1_date
            })

        # Transfer 2
        if t2_club_col and pd.notna(row.get(t2_club_col)) and str(row.get(t2_club_col)).strip():
            raw_t2_club = str(row.get(t2_club_col)).strip()
            raw_t2_date = row.get(t2_date_col) if t2_date_col else None
            t2_date = parse_transfer_date(raw_t2_date)
            from_club_2 = player_transfers[-1]["to_club"] if player_transfers else orig_club
            player_transfers.append({
                "transfer_num": 2,
                "from_club": from_club_2,
                "to_club": clean_club_fn(raw_t2_club),
                "date": t2_date
            })

        # Transfer 3+ (Attempted / blocked per Rule A13)
        if t3_club_col and pd.notna(row.get(t3_club_col)) and str(row.get(t3_club_col)).strip():
            raw_t3_club = str(row.get(t3_club_col)).strip()
            raw_t3_date = row.get(t3_date_col) if t3_date_col else None
            t3_date = parse_transfer_date(raw_t3_date)
            from_club_3 = player_transfers[-1]["to_club"] if player_transfers else orig_club
            player_transfers.append({
                "transfer_num": 3,
                "from_club": from_club_3,
                "to_club": clean_club_fn(raw_t3_club),
                "date": t3_date
            })

        for t_info in player_transfers:
            t_num = t_info["transfer_num"]
            t_to = t_info["to_club"]
            t_from = t_info["from_club"]
            t_date = t_info["date"]

            if t_num > 2:
                blocked_records.append({
                    "Player": p_name,
                    "Surname": p_surname,
                    "First Name": p_firstname,
                    "From Club": t_from,
                    "Attempted To Club": t_to,
                    "Transfer Number": t_num,
                    "Transfer Date": t_date.strftime("%d/%m/%Y") if pd.notna(t_date) else "Unknown",
                    "Transfer Date Parsed": t_date,
                    "Status": "BLOCKED",
                    "Violation": "Rule A13: Maximum 2 transfers per season exceeded"
                })
            else:
                is_on_or_after_april = pd.notna(t_date) and t_date >= april_first
                fee = 25.00 if is_on_or_after_april else 0.00
                fee_infraction = is_on_or_after_april

                valid_records.append({
                    "Player": p_name,
                    "Surname": p_surname,
                    "First Name": p_firstname,
                    "From Club": t_from,
                    "To Club": t_to,
                    "Transfer Number": t_num,
                    "Transfer Date": t_date.strftime("%d/%m/%Y") if pd.notna(t_date) else "Unknown",
                    "Transfer Date Parsed": t_date,
                    "Fee Due (£)": fee,
                    "Fee Infraction": fee_infraction,
                    "Rule A13 Compliant": True
                })

    df_valid = sort_transfer_records_dataframe(pd.DataFrame(valid_records))
    df_blocked = sort_transfer_records_dataframe(pd.DataFrame(blocked_records))

    club_summary: Dict[str, Dict[str, Any]] = {}
    total_fees = 0.0

    if not df_valid.empty:
        for club, group in df_valid.groupby("To Club"):
            c_clean = clean_club_fn(club)
            infractions_group = group[group["Fee Infraction"]]
            infraction_count = len(infractions_group)
            fees = infraction_count * 25.00
            total_fees += fees
            club_summary[c_clean] = {
                "total_transfers": len(group),
                "fee_transfers_count": infraction_count,
                "total_transfer_fees": fees,
                "players": group["Player"].tolist()
            }

    return {
        "valid_transfers": df_valid,
        "blocked_transfers": df_blocked,
        "club_summary": club_summary,
        "total_fees_assessed": total_fees
    }


# ==============================================================================
# RULE A10 / WA10: 2027 WATERFALL STARRING ROSTER ALLOCATOR
# ==============================================================================

def allocate_waterfall_starring_roster(
    player_data_df: pd.DataFrame,
    team_count: Optional[int] = None,
    domain: str = "Men's",
    manual_locks: Optional[List[str]] = None,
    departures: Optional[List[str]] = None,
    dispensations: Optional[Union[Dict[str, str], pd.DataFrame, Set[str], List[Dict[str, Any]]]] = None,
    num_teams: Optional[int] = None
) -> pd.DataFrame:
    """
    Ranks club players by projected performance rating, respects manual locks, departures,
    and Board-Approved Executive Availability Dispensations, and applies strict Rule A10 / Rule WA10
    waterfall tier quotas to assign 2027 predicted starrings.

    Inputs:
        player_data_df: DataFrame with player performance, 2026 starring tier, and 'Projected_Rating'.
        team_count: Number of senior teams fielded by the club (e.g. 2 to 6).
        domain: Competition domain ("Men's" or "Women's").
        manual_locks: Optional list of player names locked to 1st XI (e.g. captain, overseas pro).
        departures: Optional list of player names who are leaving/transferred (excluded from quota).
        dispensations: Optional dictionary, DataFrame, or list of Board-Approved Availability Dispensations
                       mapping player names to their authorized lower tier (e.g. {'Simon King': '2nd XI'}).
        num_teams: Optional alias for team_count.

    Outputs:
        pd.DataFrame: Complete roster with 'Projected_Rank', '2027_Predicted_Tier', and 'Status' flags.

    Helper Apps:
        app.py, starring_rules.py, tests/test_starring_predictor.py.
    """
    if player_data_df is None or player_data_df.empty:
        return pd.DataFrame()

    effective_team_count = team_count if team_count is not None else (num_teams if num_teams is not None else 2)
    df = player_data_df.copy()
    if "Player" in df.columns:
        df["Player"] = df["Player"].apply(lambda p: strip_club_suffix(str(p)))
    quotas = get_starring_quotas(effective_team_count, domain)

    locks_set = set(str(p).strip().lower() for p in (manual_locks or []))
    departures_set = set(str(p).strip().lower() for p in (departures or []))

    disp_map: Dict[str, str] = {}
    if dispensations is not None:
        if isinstance(dispensations, pd.DataFrame):
            p_col = next((c for c in ["Player Name", "Player", "Full Name", "Name"] if c in dispensations.columns), None)
            t_col = next((c for c in ["Allowed Lower Tier", "Allowed Tier", "Tier", "Target Tier"] if c in dispensations.columns), None)
            if p_col and t_col:
                for _, d_row in dispensations.iterrows():
                    pn = str(d_row.get(p_col, "")).strip().lower()
                    pt = str(d_row.get(t_col, "2nd XI")).strip()
                    if pn:
                        disp_map[pn] = pt
        elif isinstance(dispensations, dict):
            for k, v in dispensations.items():
                disp_map[str(k).strip().lower()] = str(v).strip()
        elif isinstance(dispensations, (list, set)):
            for item in dispensations:
                if isinstance(item, dict):
                    pn = str(item.get("Player Name", item.get("Player", ""))).strip().lower()
                    pt = str(item.get("Allowed Lower Tier", item.get("Tier", "2nd XI"))).strip()
                    if pn:
                        disp_map[pn] = pt
                else:
                    disp_map[str(item).strip().lower()] = "2nd XI"

    def tier_num(t: Any) -> int:
        s = str(t).lower().strip()
        if "1st" in s: return 1
        if "2nd" in s: return 2
        if "3rd" in s: return 3
        if "4th" in s: return 4
        if "5th" in s: return 5
        if "6th" in s: return 6
        return 99

    departing_rows: List[Dict[str, Any]] = []
    locked_rows: List[Dict[str, Any]] = []
    dispensation_rows: List[Dict[str, Any]] = []
    pool_rows: List[Dict[str, Any]] = []

    for _, row in df.iterrows():
        p_name = str(row.get("Player", "")).strip()
        p_low = p_name.lower()
        r_dict = dict(row)

        disp_field = str(row.get("Availability_Dispensation", "")).strip()
        if disp_field and disp_field.lower() not in ["none", "nan", "", "false"]:
            disp_map.setdefault(p_low, disp_field)

        if p_low in departures_set:
            r_dict["2027_Predicted_Tier"] = "Departing"
            r_dict["Status"] = "⚪ Departing"
            departing_rows.append(r_dict)
            continue

        if p_low in disp_map:
            target_tier = disp_map[p_low]
            r_dict["2027_Predicted_Tier"] = target_tier
            r_dict["Status"] = f"📜 Dispensation ({target_tier})"
            r_dict["Availability_Dispensation"] = target_tier
            dispensation_rows.append(r_dict)
            continue

        is_auto_lock = bool(row.get("Is_International", False) or row.get("Is_Overseas_Pro", False))
        if p_low in locks_set or is_auto_lock:
            r_dict["2027_Predicted_Tier"] = "1st XI"
            r_dict["Status"] = "⭐ Locked"
            locked_rows.append(r_dict)
            continue

        pool_rows.append(r_dict)

    pool_rows.sort(
        key=lambda x: (
            float(x.get("Projected_Rating", 0)),
            int(x.get("Runs", 0)),
            int(x.get("Wickets", 0)),
            -int(x.get("2026_Starred_Rank", 99))
        ),
        reverse=True
    )

    tier_order = ["1st XI", "2nd XI", "3rd XI", "4th XI", "5th XI", "6th XI"]
    active_tiers = [t for t in tier_order if t in quotas]

    capacity: Dict[str, int] = {t: quotas.get(t, 0) for t in active_tiers}
    capacity["1st XI"] = max(0, capacity.get("1st XI", 0) - len(locked_rows))
    for dr in dispensation_rows:
        dt = dr.get("2027_Predicted_Tier", "")
        if dt in capacity:
            capacity[dt] = max(0, capacity[dt] - 1)

    def is_eligible_for_tier(player_dict: Dict[str, Any], target_tier: str) -> bool:
        """
        Enforces primary team gravity constraint: If a player's modal appearance tier in 2026
        was down in the 2nd XI or lower, they cannot be automatically pushed into a 1st XI starring
        slot by points alone, unless they have a minimum of 2 appearances in the 1st XI.
        """
        if target_tier == "1st XI":
            prim = str(player_dict.get("Primary_Team", "")).strip().lower()
            is_lower_modal = any(t in prim for t in ["2nd", "3rd", "4th", "5th", "6th", "midweek"])
            t_counts = player_dict.get("Tier_Counts") or {}
            first_xi_apps = t_counts.get("1st XI", 0) if isinstance(t_counts, dict) else 0
            if is_lower_modal and first_xi_apps < 2:
                return False
        return True

    assigned_pool_rows: List[Dict[str, Any]] = []
    allocated_indices: Set[int] = set()

    for tier in active_tiers:
        slots_needed = capacity.get(tier, 0)
        filled = 0

        # Pass 1: Allocate eligible players respecting primary team gravity constraint
        for idx, r in enumerate(pool_rows):
            if idx in allocated_indices:
                continue
            if filled >= slots_needed:
                break
            if is_eligible_for_tier(r, tier):
                r_assigned = dict(r)
                r_assigned["2027_Predicted_Tier"] = tier
                assigned_pool_rows.append(r_assigned)
                allocated_indices.add(idx)
                filled += 1

        # Pass 2: Fallback to fulfill remaining quota slots if eligible candidates were exhausted
        if filled < slots_needed:
            for idx, r in enumerate(pool_rows):
                if idx in allocated_indices:
                    continue
                if filled >= slots_needed:
                    break
                r_assigned = dict(r)
                r_assigned["2027_Predicted_Tier"] = tier
                assigned_pool_rows.append(r_assigned)
                allocated_indices.add(idx)
                filled += 1

    # Remaining unassigned players are allocated to lowest unstarred tier
    lowest_unstarred_label = "Unstarred"
    for idx, r in enumerate(pool_rows):
        if idx not in allocated_indices:
            r_assigned = dict(r)
            r_assigned["2027_Predicted_Tier"] = lowest_unstarred_label
            assigned_pool_rows.append(r_assigned)
            allocated_indices.add(idx)

    active_allocated = locked_rows + dispensation_rows + assigned_pool_rows

    for r in active_allocated:
        if r.get("Status") == "⭐ Locked" or "dispensation" in str(r.get("Status", "")).lower():
            continue
        curr_t = r.get("2026_Starred_Tier", "Unstarred")
        pred_t = r.get("2027_Predicted_Tier", "Unstarred")

        c_rank = tier_num(curr_t)
        p_rank = tier_num(pred_t)

        if p_rank < c_rank:
            r["Status"] = "🟢 Promoted"
        elif p_rank == c_rank:
            r["Status"] = "🔵 Retained"
        else:
            r["Status"] = "🟠 Relegated"

    active_allocated.sort(
        key=lambda x: (
            tier_num(x.get("2027_Predicted_Tier")),
            -float(x.get("Projected_Rating", 0)),
            -int(x.get("Runs", 0))
        )
    )

    tier_counter: Dict[str, int] = {}
    for idx, r in enumerate(active_allocated, start=1):
        r["Projected_Rank"] = idx
        t = r["2027_Predicted_Tier"]
        tier_counter[t] = tier_counter.get(t, 0) + 1
        r["Tier_Rank"] = tier_counter[t]

    for idx, r in enumerate(departing_rows, start=len(active_allocated) + 1):
        r["Projected_Rank"] = idx
        r["Tier_Rank"] = idx

    all_roster = active_allocated + departing_rows
    for r in all_roster:
        c_wk = int(r.get("Catches_As_Keeper", 0))
        stumps = int(r.get("Stumpings", 0))
        total_dism = int(r.get("Keeper_Dismissals", c_wk + stumps))
        is_wk = bool(
            r.get("Is_Wicket_Keeper", False) or
            stumps > 0 or
            c_wk >= 2 or
            total_dism >= 3
        )
        r["Is_Wicket_Keeper"] = is_wk
        r["Keeper_Dismissals"] = total_dism
        runs = int(r.get("Runs", 0))
        wkts = int(r.get("Wickets", 0))

        if is_wk:
            r["Role"] = "🧤 Wicket-Keeper"
        elif runs >= 150 and wkts >= 8:
            r["Role"] = "🏏 All-Rounder"
        elif wkts >= 8:
            r["Role"] = "🎯 Bowler"
        else:
            r["Role"] = "🏏 Batter"

    res_df = pd.DataFrame(all_roster)
    return res_df


def audit_first_xi_wicket_keeper_balance(
    roster_df: pd.DataFrame
) -> Dict[str, Any]:
    """
    Audits the projected starring roster to verify whether the 1st XI contains
    at least one recognized specialist wicket-keeper.

    Inputs:
        roster_df: DataFrame output of allocate_waterfall_starring_roster containing
                   'Player', '2027_Predicted_Tier', 'Role', 'Is_Wicket_Keeper', 'Stumpings'.

    Outputs:
        Dict[str, Any]:
            - 'is_compliant': bool (True if at least 1 keeper in 1st XI)
            - 'first_xi_keepers': List[str] (names of keepers in 1st XI)
            - 'top_candidate': Optional[str] (best keeper in club if 1st XI has none)
            - 'top_candidate_tier': Optional[str]
            - 'top_candidate_dismissals': int
            - 'status_message': str

    Helper Apps:
        app.py, starring_rules.py, tests/test_starring_predictor.py.
    """
    if roster_df is None or roster_df.empty:
        return {
            "is_compliant": False,
            "first_xi_keepers": [],
            "top_candidate": None,
            "top_candidate_tier": None,
            "top_candidate_dismissals": 0,
            "status_message": "No roster data available."
        }

    first_xi = roster_df[roster_df["2027_Predicted_Tier"].astype(str).str.strip().str.lower() == "1st xi"]

    first_xi_keepers: List[str] = []
    for _, row in first_xi.iterrows():
        p_name = str(row.get("Player", "")).strip()
        c_wk = int(row.get("Catches_As_Keeper", 0))
        stumps = int(row.get("Stumpings", 0))
        total_dism = int(row.get("Keeper_Dismissals", c_wk + stumps))
        is_wk = bool(
            row.get("Is_Wicket_Keeper", False) or
            "keeper" in str(row.get("Role", "")).lower() or
            stumps > 0 or
            c_wk >= 2 or
            total_dism >= 3
        )
        if is_wk and p_name:
            first_xi_keepers.append(p_name)

    if first_xi_keepers:
        top_k = first_xi_keepers[0]
        top_k_row = first_xi[first_xi["Player"] == top_k].iloc[0] if not first_xi[first_xi["Player"] == top_k].empty else first_xi.iloc[0]
        top_dism = int(top_k_row.get("Keeper_Dismissals", int(top_k_row.get("Catches_As_Keeper", 0)) + int(top_k_row.get("Stumpings", 0))))
        return {
            "is_compliant": True,
            "first_xi_keepers": first_xi_keepers,
            "top_candidate": top_k,
            "top_candidate_tier": "1st XI",
            "top_candidate_dismissals": top_dism,
            "status_message": f"✅ Compliant: {', '.join(first_xi_keepers)} projected in 1st XI."
        }

    # If no keeper in 1st XI, find the best keeper candidate across the rest of the club
    all_keepers: List[Dict[str, Any]] = []
    for _, row in roster_df.iterrows():
        c_wk = int(row.get("Catches_As_Keeper", 0))
        stumps = int(row.get("Stumpings", 0))
        total_dism = int(row.get("Keeper_Dismissals", c_wk + stumps))
        is_wk = bool(
            row.get("Is_Wicket_Keeper", False) or
            "keeper" in str(row.get("Role", "")).lower() or
            stumps > 0 or
            c_wk >= 2 or
            total_dism >= 3
        )
        if is_wk:
            k_dict = dict(row)
            k_dict["Keeper_Dismissals"] = total_dism
            all_keepers.append(k_dict)

    if all_keepers:
        all_keepers.sort(
            key=lambda x: (
                int(x.get("Keeper_Dismissals", int(x.get("Stumpings", 0)) + int(x.get("Catches_As_Keeper", 0)))),
                float(x.get("Projected_Rating", 0))
            ),
            reverse=True
        )
        top_cand = all_keepers[0]
        top_name = str(top_cand.get("Player", "")).strip()
        top_tier = str(top_cand.get("2027_Predicted_Tier", "2nd XI"))
        top_dism = int(top_cand.get("Keeper_Dismissals", int(top_cand.get("Catches_As_Keeper", 0)) + int(top_cand.get("Stumpings", 0))))
        return {
            "is_compliant": False,
            "first_xi_keepers": [],
            "top_candidate": top_name,
            "top_candidate_tier": top_tier,
            "top_candidate_dismissals": top_dism,
            "status_message": f"⚠️ Advisory: No specialist wicket-keeper projected in 1st XI. Top candidate: {top_name} ({top_tier})."
        }

    return {
        "is_compliant": False,
        "first_xi_keepers": [],
        "top_candidate": None,
        "top_candidate_tier": None,
        "top_candidate_dismissals": 0,
        "status_message": "⚠️ Advisory: No recognized wicket-keeper found across club rosters."
    }


def build_club_projected_starring_workbook(
    roster_df: pd.DataFrame,
    club_name: str
) -> openpyxl.Workbook:
    """
    Constructs a standard OpenPyXL workbook for a club's 2027 projected starring roster
    containing two worksheets formatted according to GEMINI.md guidelines:
    - Tab 1 ('2027 Projected Roster'): Complete waterfall allocation roster.
    - Tab 2 ('At Risk & Relegated'): Filtered roster containing players flagged for relegation (🟠).

    Inputs:
        roster_df: Complete DataFrame output from allocate_waterfall_starring_roster.
        club_name: Name of the NCU cricket club (e.g. 'Waringstown', 'Instonians').

    Outputs:
        openpyxl.Workbook: Formatted workbook with navy headers, white bold text, auto-fitted
                           columns, and freeze panes at row A2.

    Helper Apps:
        app.py, engine.py, starring_rules.py, tests/test_starring_predictor.py.
    """
    wb = openpyxl.Workbook()

    export_cols = [
        "Projected_Rank", "Player", "Role", "2027_Predicted_Tier", "Status",
        "2026_Starred_Tier", "Primary_Team", "Projected_Rating",
        "Runs", "Wickets", "Catches", "Stumpings", "Matches"
    ]
    valid_cols = [c for c in export_cols if c in roster_df.columns] if roster_df is not None and not roster_df.empty else export_cols

    # Tab 1: 2027 Projected Roster
    ws1 = wb.active
    ws1.title = "2027 Projected Roster"
    ws1.append(valid_cols)
    if roster_df is not None and not roster_df.empty:
        for _, row in roster_df[valid_cols].iterrows():
            ws1.append([row[c] for c in valid_cols])
    format_excel_worksheet_standard(ws1)

    # Tab 2: At Risk & Relegated
    ws2 = wb.create_sheet(title="At Risk & Relegated")
    ws2.append(valid_cols)
    if roster_df is not None and not roster_df.empty and "Status" in roster_df.columns:
        rel_df = roster_df[roster_df["Status"].astype(str).str.contains("🟠", na=False)]
        for _, row in rel_df[valid_cols].iterrows():
            ws2.append([row[c] for c in valid_cols])
    format_excel_worksheet_standard(ws2)

    return wb


# ==============================================================================
# 2027 IN-SEASON ROSTER REALISM & PERFORMANCE MONITOR
# ==============================================================================

def detect_ghost_stars(
    roster_df: pd.DataFrame,
    international_exemptions: Optional[Set[str]] = None,
    dispensations: Optional[Union[Set[str], Dict[str, str], pd.DataFrame, List[Any]]] = None,
    early_season_weeks: int = 3
) -> pd.DataFrame:
    """
    Evaluates early-season appearance and inactivity records for club starred rosters.
    Identifies top-tier assets (1st XI and 2nd XI) with zero appearances or >= 3 weeks
    (21 days) of inactivity who lack international duty exemptions or Board-Approved
    Availability Dispensations, flagging potential roster padding or ghost stars.

    Inputs:
        roster_df: DataFrame containing player roster data with columns such as
                   'Player', 'Club', 'Current Starred Tier' (or '2027_Predicted_Tier' / '2026_Starred_Tier'),
                   'Matches' (or 'Total_Matches'), 'Days Inactive', 'Is_International'.
        international_exemptions: Optional set of player names possessing active international duty exemptions.
        dispensations: Optional set, dict, or DataFrame of board-approved availability dispensations.
        early_season_weeks: Inactivity threshold in weeks (default 3 weeks / 21 days).

    Outputs:
        pd.DataFrame: Audit table of flagged ghost star players with columns:
                      ['Rank', 'Player', 'Club', 'Current Starred Tier', 'Matches', 'Days Inactive', 'Alert', 'Details'].

    Helper Apps:
        app.py, starring_rules.py, engine.py, tests/test_roster_realism.py.
    """
    if roster_df is None or roster_df.empty:
        return pd.DataFrame(columns=[
            "Rank", "Player", "Club", "Current Starred Tier", "Matches", "Days Inactive", "Alert", "Details"
        ])

    intl_set = set(str(p).strip().lower() for p in (international_exemptions or set()))
    threshold_days = early_season_weeks * 7

    disp_set: Set[str] = set()
    if dispensations is not None:
        if isinstance(dispensations, pd.DataFrame):
            p_col = next((c for c in ["Player Name", "Player", "Full Name", "Name"] if c in dispensations.columns), None)
            if p_col:
                disp_set = set(dispensations[p_col].dropna().astype(str).str.strip().str.lower())
        elif isinstance(dispensations, dict):
            disp_set = set(str(k).strip().lower() for k in dispensations.keys())
        elif isinstance(dispensations, (list, set)):
            for item in dispensations:
                if isinstance(item, dict):
                    pn = str(item.get("Player Name", item.get("Player", ""))).strip().lower()
                    if pn: disp_set.add(pn)
                else:
                    disp_set.add(str(item).strip().lower())

    flagged_rows: List[Dict[str, Any]] = []

    for idx, row in roster_df.iterrows():
        p_name = str(row.get("Player", row.get("Full Name", row.get("Name", "")))).strip()
        if not p_name:
            continue

        club_name = str(row.get("Club", row.get("Club Name", ""))).strip()

        # Resolve tier
        tier = ""
        for c in ["Current Starred Tier", "2027_Predicted_Tier", "2026_Starred_Tier", "Starred Tier", "XI_Level", "Team", "Tier"]:
            if c in row and pd.notna(row[c]):
                tier = str(row[c]).strip()
                break

        t_lower = tier.lower()
        # Only evaluate 1st XI and 2nd XI starred assets
        is_top_tier = ("1st" in t_lower) or ("2nd" in t_lower)
        if not is_top_tier:
            continue

        # Check International Duty Exemption
        is_intl = bool(row.get("Is_International", False) or row.get("is_international", False))
        if is_intl or (p_name.lower() in intl_set):
            continue

        # Check Board-Approved Availability Dispensation
        disp_val = str(row.get("Availability_Dispensation", "")).strip()
        has_disp = (
            (bool(disp_val) and disp_val.lower() not in ["none", "nan", "", "false"])
            or ("dispensation" in str(row.get("Status", "")).lower())
            or (p_name.lower() in disp_set)
        )
        if has_disp:
            continue

        # Matches count
        matches = 0
        for m_col in ["Matches", "Total_Matches", "Appearances", "Eligible Appearances"]:
            if m_col in row and pd.notna(row[m_col]):
                try:
                    matches = int(row[m_col])
                    break
                except (ValueError, TypeError):
                    pass

        # Days inactive
        days_inactive = 0
        for d_col in ["Days Inactive", "days_inactive"]:
            if d_col in row and pd.notna(row[d_col]):
                try:
                    days_inactive = int(row[d_col])
                    break
                except (ValueError, TypeError):
                    pass

        is_ghost = False
        reasons = []

        if matches == 0:
            is_ghost = True
            reasons.append("0 matches recorded in early season")
        elif days_inactive >= threshold_days:
            is_ghost = True
            reasons.append(f"Inactive for {days_inactive} days (>= {threshold_days} days / {early_season_weeks} wks)")

        if is_ghost:
            flagged_rows.append({
                "Rank": len(flagged_rows) + 1,
                "Player": p_name,
                "Club": club_name,
                "Current Starred Tier": tier,
                "Matches": matches,
                "Days Inactive": days_inactive,
                "Alert": "⚠️ Potential Ghost Star (Roster Padding)",
                "Details": "; ".join(reasons)
            })

    if not flagged_rows:
        return pd.DataFrame(columns=[
            "Rank", "Player", "Club", "Current Starred Tier", "Matches", "Days Inactive", "Alert", "Details"
        ])
    return pd.DataFrame(flagged_rows)


def compute_promotion_cascade_path(
    current_tier: str,
    team_count: int,
    domain: str = "Men's",
    is_elite: bool = False
) -> Tuple[str, str]:
    """
    Computes the recommended upward promotion tier and cascade displacement path
    for lower-tier players over-performing their registered level, aligned with Rule A10/WA10.

    Inputs:
        current_tier: Player's current team level (e.g. '3rd XI', '4th XI', '5th XI', 'Unstarred').
        team_count: Total senior teams fielded by the club (2 to 6).
        domain: Competition domain ("Men's" or "Women's").
        is_elite: Whether the player's metrics warrant fast-tracking directly to 1st XI.

    Outputs:
        Tuple[str, str]: (recommended_tier, cascade_description).

    Helper Apps:
        app.py, starring_rules.py, tests/test_roster_realism.py.
    """
    t_clean = str(current_tier).strip()
    t_low = t_clean.lower()

    if is_elite:
        return "1st XI", "Fast-Track Promotion to 1st XI (Displaces lowest 1st XI slot to 2nd XI)"

    if "3rd" in t_low:
        return "2nd XI", "Advance to 2nd XI (Displaces lowest 2nd XI slot to 3rd XI)"
    elif "4th" in t_low:
        return "3rd XI", "Advance to 3rd XI (Displaces lowest 3rd XI slot to 4th XI)"
    elif "5th" in t_low:
        return "4th XI", "Advance to 4th XI (Displaces lowest 4th XI slot to 5th XI)"
    elif "6th" in t_low:
        return "5th XI", "Advance to 5th XI (Displaces lowest 5th XI slot to 6th XI)"
    else:
        # Unstarred or Midweek
        if team_count <= 2:
            return "2nd XI", "Advance to 2nd XI (Displaces lowest 2nd XI slot to Unstarred)"
        elif team_count == 3:
            return "3rd XI", "Advance to 3rd XI (Displaces lowest 3rd XI slot to Unstarred)"
        else:
            return "4th XI", "Advance to 4th XI (Displaces lowest 4th XI slot to Unstarred)"


def evaluate_lower_tier_overperformance(
    player_perf_df: pd.DataFrame,
    team_count: int = 3,
    domain: str = "Men's",
    dispensations: Optional[Union[Set[str], Dict[str, str], pd.DataFrame, List[Any]]] = None
) -> pd.DataFrame:
    """
    Monitors player scorecards specifically for unstarred or lower-tier players (3rd XI and lower)
    to detect statistical ceiling breaches within their first 3 appearances:
    - Batting: Runs > 150 OR Strike Rate > 130% OR Average > 55.0 (Minimum 3 innings).
    - Bowling: Wickets > 10 OR Economy < 3.50 OR Average < 12.0 (Minimum 15 overs).
    Labels breached records and computes an automated upward promotion cascade path.
    Board-Approved Availability Dispensations are exempt from over-performance alerts.

    Inputs:
        player_perf_df: DataFrame with player performance aggregates or match breakdowns.
        team_count: Club senior teams count for Rule A10/WA10 quota modeling.
        domain: Competition domain ("Men's" or "Women's").
        dispensations: Optional set, dict, or DataFrame of board-approved availability dispensations.

    Outputs:
        pd.DataFrame: Table of flagged lower-tier over-performers with recommended promotion paths.

    Helper Apps:
        app.py, starring_rules.py, engine.py, tests/test_roster_realism.py.
    """
    if player_perf_df is None or player_perf_df.empty:
        return pd.DataFrame(columns=[
            "Rank", "Player", "Club", "Current Tier", "Recommended Tier",
            "Alert", "Breach Reasons", "Promotion Path",
            "Runs", "Bat Avg", "SR", "Wickets", "Bowl Avg", "Econ"
        ])

    disp_set: Set[str] = set()
    if dispensations is not None:
        if isinstance(dispensations, pd.DataFrame):
            p_col = next((c for c in ["Player Name", "Player", "Full Name", "Name"] if c in dispensations.columns), None)
            if p_col:
                disp_set = set(dispensations[p_col].dropna().astype(str).str.strip().str.lower())
        elif isinstance(dispensations, dict):
            disp_set = set(str(k).strip().lower() for k in dispensations.keys())
        elif isinstance(dispensations, (list, set)):
            for item in dispensations:
                if isinstance(item, dict):
                    pn = str(item.get("Player Name", item.get("Player", ""))).strip().lower()
                    if pn: disp_set.add(pn)
                else:
                    disp_set.add(str(item).strip().lower())

    def parse_overs_to_decimal(ov_val: Any) -> float:
        try:
            val = float(ov_val)
            ov_int = int(val)
            balls = round((val - ov_int) * 10)
            if balls >= 6:
                ov_int += balls // 6
                balls = balls % 6
            return float(ov_int) + (balls / 6.0)
        except (ValueError, TypeError):
            return 0.0

    flagged: List[Dict[str, Any]] = []

    for _, row in player_perf_df.iterrows():
        p_name = str(row.get("Player", row.get("Full Name", row.get("Name", "")))).strip()
        if not p_name:
            continue

        club_name = str(row.get("Club", row.get("Club Name", ""))).strip()

        # Check Board-Approved Availability Dispensation
        disp_val = str(row.get("Availability_Dispensation", "")).strip()
        has_disp = (
            (bool(disp_val) and disp_val.lower() not in ["none", "nan", "", "false"])
            or ("dispensation" in str(row.get("Status", "")).lower())
            or (p_name.lower() in disp_set)
        )
        if has_disp:
            continue

        # Determine current tier
        curr_tier = ""
        for c in ["Primary_Team", "Current Tier", "Current Starred Tier", "2026_Starred_Tier", "Starred Tier", "XI_Level"]:
            if c in row and pd.notna(row[c]):
                curr_tier = str(row[c]).strip()
                break
        if not curr_tier:
            curr_tier = "Unstarred"

        t_low = curr_tier.lower()
        # Only evaluate lower-tier or unstarred players (exclude 1st XI and 2nd XI)
        if ("1st" in t_low) or ("2nd" in t_low):
            continue

        # Extract stats from Match_Performances (first 3 lower-tier appearances) or row metrics
        matches = row.get("Match_Performances", [])
        lower_matches = [
            m for m in matches
            if any(lt in str(m.get("Tier", "")).lower() for lt in ["3rd", "4th", "5th", "6th", "midweek"])
        ] if isinstance(matches, list) and matches else []

        if lower_matches:
            # First 3 lower-tier appearances
            eval_matches = lower_matches[:3]
            runs = sum(int(m.get("Runs", 0)) for m in eval_matches)
            inns = sum(1 for m in eval_matches if m.get("Runs") is not None and str(m.get("Runs", "")).strip().upper() != "DNB")
            not_outs = sum(1 for m in eval_matches if m.get("Not_Out", False))
            balls_faced = sum(int(m.get("Balls", 0)) for m in eval_matches)
            sr_val = ((runs / balls_faced) * 100.0) if balls_faced > 0 else float(row.get("Strike_Rate", row.get("SR", 0.0)))

            wkts = sum(int(m.get("Wickets", 0)) for m in eval_matches)
            runs_conc = sum(int(m.get("Runs_Conceded", 0)) for m in eval_matches)
            overs_dec = sum(parse_overs_to_decimal(m.get("Overs", 0.0)) for m in eval_matches)
        else:
            runs = int(row.get("Runs", 0))
            inns = int(row.get("Innings_Bat", row.get("Innings", row.get("Inns", 3 if runs > 0 else 0))))
            not_outs = int(row.get("Not_Outs", row.get("NO", 0)))
            sr_val = float(row.get("Strike_Rate", row.get("SR", 0.0)))
            balls_faced = int(row.get("Balls", 0))
            if sr_val == 0.0 and balls_faced > 0:
                sr_val = (runs / balls_faced) * 100.0

            wkts = int(row.get("Wickets", 0))
            runs_conc = int(row.get("Runs_Conceded", 0))
            overs_raw = row.get("Overs", 0.0)
            overs_dec = parse_overs_to_decimal(overs_raw)

        # Batting metrics (requires min 3 innings)
        bat_breaches = []
        dismissals = max(0, inns - not_outs)
        bat_avg = (runs / dismissals) if dismissals > 0 else float(runs)

        if inns >= 3:
            if runs > 150:
                bat_breaches.append(f"Runs: {runs} (> 150)")
            if sr_val > 130.0:
                bat_breaches.append(f"SR: {sr_val:.1f}% (> 130%)")
            if bat_avg > 55.0:
                bat_breaches.append(f"Avg: {bat_avg:.1f} (> 55.0)")

        # Bowling metrics (requires min 15 overs)
        bowl_breaches = []
        econ = (runs_conc / overs_dec) if overs_dec > 0 else 999.0
        bowl_avg = (runs_conc / wkts) if wkts > 0 else 999.0

        if overs_dec >= 15.0:
            if wkts > 10:
                bowl_breaches.append(f"Wickets: {wkts} (> 10)")
            if econ < 3.50:
                bowl_breaches.append(f"Economy: {econ:.2f} (< 3.50)")
            if bowl_avg < 12.0 and wkts > 0:
                bowl_breaches.append(f"Avg: {bowl_avg:.1f} (< 12.0)")

        total_breaches = bat_breaches + bowl_breaches
        if total_breaches:
            is_elite = (runs >= 250) or (wkts >= 15)
            rec_tier, cascade_path = compute_promotion_cascade_path(curr_tier, team_count, domain, is_elite=is_elite)

            flagged.append({
                "Rank": len(flagged) + 1,
                "Player": p_name,
                "Club": club_name,
                "Current Tier": curr_tier,
                "Recommended Tier": rec_tier,
                "Alert": "🔥 Lower-Tier Over-Performer (Review Required)",
                "Breach Reasons": "; ".join(total_breaches),
                "Promotion Path": cascade_path,
                "Runs": runs,
                "Bat Avg": round(bat_avg, 1) if inns > 0 else 0.0,
                "SR": round(sr_val, 1) if sr_val > 0 else 0.0,
                "Wickets": wkts,
                "Bowl Avg": round(bowl_avg, 1) if wkts > 0 else 0.0,
                "Econ": round(econ, 2) if overs_dec > 0 else 0.0
            })

    if not flagged:
        return pd.DataFrame(columns=[
            "Rank", "Player", "Club", "Current Tier", "Recommended Tier",
            "Alert", "Breach Reasons", "Promotion Path",
            "Runs", "Bat Avg", "SR", "Wickets", "Bowl Avg", "Econ"
        ])

    df_res = pd.DataFrame(flagged)
    if not df_res.empty:
        df_res = df_res.sort_values(by=["Runs", "Wickets"], ascending=[False, False]).reset_index(drop=True)
        df_res["Rank"] = range(1, len(df_res) + 1)
    return df_res


# ==============================================================================
# SEASONAL ARCHIVING & MULTI-YEAR INGESTION ENGINE (RULE INTEGRATION)
# ==============================================================================

def ensure_archive_directory(project_root: Optional[str] = None) -> str:
    """
    Validates that an archive/ folder exists within the project root, creating it if needed.

    Inputs:
        project_root: Optional root directory path (defaults to current working directory).

    Outputs:
        str: Absolute path to the validated archive/ directory.

    Helper Apps:
        app.py, engine.py, starring_rules.py, tests/test_seasonal_rollover.py.
    """
    import engine as eng
    return eng.ensure_archive_directory(project_root=project_root)


def extract_season_from_path(file_path: str, default_season: int = 2026) -> int:
    """
    Extracts a 4-digit season year from a file path or filename string.

    Inputs:
        file_path: File system path or filename.
        default_season: Fallback year if no valid 4-digit year is found (default 2026).

    Outputs:
        int: Extracted season year integer.

    Helper Apps:
        app.py, engine.py, starring_rules.py, tests/test_seasonal_rollover.py.
    """
    import engine as eng
    return eng.extract_season_from_path(file_path=file_path, default_season=default_season)


def archive_completed_season(target_year: int = 2026, project_root: Optional[str] = None) -> List[str]:
    """
    Safely archives active season raw data sheets for Saturday/Open, Women's, and Midweek cricket
    into the archive/ directory stamped with the year suffix (e.g., Open_Season_2026.xlsx).

    Inputs:
        target_year: Season year to stamp onto archived files (default 2026).
        project_root: Optional project root folder (defaults to current working directory).

    Outputs:
        List[str]: Paths of the newly created archive files.

    Helper Apps:
        app.py, engine.py, starring_rules.py, tests/test_seasonal_rollover.py.
    """
    import engine as eng
    return eng.archive_completed_season(target_year=target_year, project_root=project_root)


def generate_clean_season_templates(
    target_year: int = 2027,
    project_root: Optional[str] = None,
    files_to_clean: Optional[List[str]] = None
) -> List[str]:
    """
    Overwrites the active season scorecard and structure files with blank rows,
    preserving only the formal structural columns, validation strings, and default
    #1F4E78 dark navy header styles (GEMINI.md protocol).

    Inputs:
        target_year: Upcoming season to initialize (default 2027).
        project_root: Optional project root folder.
        files_to_clean: Optional list of file paths to wipe/initialize.

    Outputs:
        List[str]: Paths of the newly generated clean season templates.

    Helper Apps:
        app.py, engine.py, starring_rules.py, tests/test_seasonal_rollover.py.
    """
    import engine as eng
    return eng.generate_clean_season_templates(
        target_year=target_year,
        project_root=project_root,
        files_to_clean=files_to_clean
    )


def load_multi_season_scorecard_frames(
    domain: str = "Men's",
    stat_type: str = "bat",
    primary_file: Optional[str] = None,
    include_irish: bool = True,
    include_midweek: bool = False,
    include_archive: bool = True,
    project_root: Optional[str] = None,
    custom_files: Optional[Dict[str, str]] = None
) -> List[pd.DataFrame]:
    """
    Dynamically searches pattern arrays across active root files and any .xlsx files in archive/,
    returning a combined list of DataFrames tagged with an explicit 'Season' metadata column.

    Inputs:
        domain: Competition domain ("Men's", "Women's", or "Midweek").
        stat_type: Stat file category ("bat" or "bowl").
        primary_file: Optional path to an active scorecard file.
        include_irish: Whether to include Irish Cup / National Cup scorecards (Men's).
        include_midweek: Whether to include Midweek League scorecards.
        include_archive: Whether to scan and load historical files from the archive/ folder.
        project_root: Optional project root folder (defaults to current working directory).
        custom_files: Optional overrides mapping file keys.

    Outputs:
        List[pd.DataFrame]: Loaded dataframes each enriched with an explicit 'Season' column.

    Helper Apps:
        app.py, engine.py, starring_rules.py, tests/test_seasonal_rollover.py.
    """
    import engine as eng
    return eng.load_multi_season_scorecard_frames(
        domain=domain,
        stat_type=stat_type,
        primary_file=primary_file,
        include_irish=include_irish,
        include_midweek=include_midweek,
        include_archive=include_archive,
        project_root=project_root,
        custom_files=custom_files
    )
