# NCU Cricket Hub — System Architecture & Maintenance Guide

## Document Overview
This document serves as the master engineering and architectural maintenance reference for the **NCU Cricket Hub** application suite, specifically covering:
- [`app.py`](file:///c:/Users/Wylie/Downloads/Python%20Environment/app.py) — Starring Registry, Rule A10–A13 Administration & Word Document Generation
- [`stats_app.py`](file:///c:/Users/Wylie/Downloads/Python%20Environment/stats_app.py) — League & Cup Season Statistics, Batting/Bowling Averages & Performance Leaderboards
- [`finance_app.py`](file:///c:/Users/Wylie/Downloads/Python%20Environment/finance_app.py) — Registration Fee Auditing, Sharon Invoicing & Penalty Ledger Systems
- [`secretary_app.py`](file:///c:/Users/Wylie/Downloads/Python%20Environment/secretary_app.py) — Match Secretary Portal, Scorecard Ingestion & Quota Verification
- [`engine.py`](file:///c:/Users/Wylie/Downloads/Python%20Environment/engine.py) / [`starring_rules.py`](file:///c:/Users/Wylie/Downloads/Python%20Environment/starring_rules.py) — Core Business Logic, Calamine I/O, Player Disambiguation & Regulatory Engines

---

## 1. Architectural Principles & Deployment Isolation

### 1.1 Multi-App Independent Deployment Model
Each front-end application (`app.py`, `stats_app.py`, `finance_app.py`, and `secretary_app.py`) is deployed to production from **separate, self-contained Git repositories**.

```
                           +--------------------------------+
                           |       Production Target        |
                           +--------------------------------+
                                           |
      +--------------------+---------------+--------------------+--------------------+
      |                    |                                    |                    |
+-------------+    +-----------------+                  +-----------------+    +-----------------+
|   app.py    |    |  stats_app.py   |                  | finance_app.py  |    |secretary_app.py |
| (Starring & |    |  (Statistics &  |                  | (Invoicing &    |    |  (Scorecards &  |
|  Registry)  |    |    Averages)    |                  |   Fee Audits)   |    |    Auditing)    |
+-------------+    +-----------------+                  +-----------------+    +-----------------+
      |                    |                                    |                    |
      +--------------------+---------------+--------------------+--------------------+
                                           | (Local copy per repo)
                           +--------------------------------+
                           |  engine.py / starring_rules.py |
                           |    (Core Engine & Utilities)   |
                           +--------------------------------+
```

### 1.2 Isolation Constraints
1. **Zero Cross-App Dependencies:** Never import functions, classes, or globals between frontend scripts (e.g., `app.py` must never import from `finance_app.py` or `stats_app.py`).
2. **Self-Contained Backend Libraries:** Any helper functions, algorithms, or constants shared across applications must reside in that application repository's local copy of [`engine.py`](file:///c:/Users/Wylie/Downloads/Python%20Environment/engine.py) or [`starring_rules.py`](file:///c:/Users/Wylie/Downloads/Python%20Environment/starring_rules.py).
3. **No Dynamic Cross-App File Sharing:** Configuration paths, threshold stores, and assets must be resolved dynamically relative to the local repository root.

---

## 2. High-Performance Calamine Excel I/O Wrappers

### 2.1 The Calamine Engine Wrapper (`read_excel_calamine`)
Large Excel workbooks (e.g. multi-megabyte NV Play ball-by-ball scorecards, Sport80 revenue reports, and master starring registries) experience severe quadratic parsing bottlenecks when read using standard openpyxl or un-engined Pandas calls.

All spreadsheet ingestion across the workspace is standardized to route strictly through `eng.read_excel_calamine()`:

```python
def read_excel_calamine(io_target: Any, *args: Any, **kwargs: Any) -> Any:
    """
    Reads an Excel file utilizing the high-performance 'calamine' engine when
    available and appropriate, falling back cleanly to the standard pandas/openpyxl
    engine if calamine is unavailable, unsupported, or if an ExcelFile is supplied.
    If calamine fails on a seekable stream, resets stream offset via seek(0) before fallback.
    """
```

#### Key Implementation Characteristics:
- **Rust-Powered Parsing:** Leverages the native `calamine` Rust library, accelerating worksheet loading times by **10× to 25×** compared to standard openpyxl.
- **Fail-Safe Fallback:** If `calamine` is not installed or encounters an unsupported spreadsheet feature, the wrapper catches the exception, resets stream pointers via `seek(0)`, and falls back cleanly to Pandas' default engine.
- **Seekable Stream Safety:** When accepting `io.BytesIO` streams (e.g., file uploads or generated in-memory exports), the stream pointer is defensively positioned at `0` before and after invocation.
- **Elimination of Un-Engined Reads:** Direct calls to `pd.read_excel(filepath)` without specifying `read_excel_calamine()` are strictly prohibited across all frontend and backend modules.

### 2.2 Standardized Excel Formatting Protocol (Mandatory)
Every workbook generated or updated must implement the 3-step formatting protocol:

1. **Dark Navy Accent Header:**
   - Font: White bold text (`#FFFFFF`, `bold=True`).
   - Fill: Solid Dark Navy Blue fill (`#1F4E78`).
2. **Freeze Panes:**
   - Row 2 freeze panes applied to all worksheets (`ws.freeze_panes = 'A2'`).
3. **Dynamic Max-Length Column Auto-Fitting:**
   - Width computed dynamically: `ws.column_dimensions[col_letter].width = max(max_len + 2, 10)`.
4. **Data Integrity Safeguards:**
   - **Bowling Figures Preservation:** Columns containing scores or bowling figures (e.g., `'1-21'`, `'5/24'`) must be explicitly formatted as Text (`@`) to prevent Excel from auto-converting figures to calendar dates.
   - **Ghost Row Prevention:** Never write empty styled cells beyond actual data boundaries.

### 2.3 Dynamic Workspace Path Resolution
Hardcoded filesystem paths (e.g., `C:\Users\...` or relative assumptions that fail when launched from child directories) are replaced with dynamic directory anchors:

```python
WORKSPACE_ROOT = os.path.dirname(os.path.abspath(__file__))
starring_history_path = os.path.join(WORKSPACE_ROOT, "NCU_Club_Starring_History.xlsx")
```

---

## 3. Player Name Resolution & Disambiguation Pipeline

### 3.1 Multi-Tier Canonical Player Identity Resolution
Cricket scorecards from NV Play frequently feature nicknames, abbreviated names, typographical errors, and absent UUIDs, while Sport80 registries contain legal registered names.

Resolution is handled centrally by [`engine.resolve_player_from_row()`](file:///c:/Users/Wylie/Downloads/Python%20Environment/engine.py#L1549):

```
                        [Input: Scorecard / Stats Row]
                                      |
                                      v
                        +---------------------------+
                        |  1. Extract Player UUID   |
                        +---------------------------+
                                      |
                  +-------------------+-------------------+
                  | UUID in id_map?                       | No UUID / Not in id_map
                  v                                       v
      +-----------------------+               +-----------------------+
      | 2. UUID Match & Alias |               | 3. Name & Club Context|
      |    Sport80 Lookup     |               |    Fast-Path Index    |
      +-----------------------+               +-----------------------+
                  |                                       |
                  |                               +-------+-------+
                  |                               | Exact Match?  |
                  |                               +---------------+
                  |                               | Yes       | No (Fuzzy Fallback)
                  |                               v           v
                  |               +-------------------+   +-------------------+
                  |               | Club Verification |   | TheFuzz TokenSort |
                  |               | against Fixture   |   | (Score >= 90)     |
                  |               +-------------------+   +-------------------+
                  |                               |           |
                  +---------------+---------------+-----------+
                                  |
                                  v
                  +-------------------------------+
                  |  4. KNOWN_DUPLICATES Check    |
                  |  Append Club: "Name (Club)"   |
                  +-------------------------------+
                                  |
                                  v
                  +-------------------------------+
                  |  5. Celtic Casing Correction  |
                  |  (e.g., "McMaster", "O'Neill")|
                  +-------------------------------+
                                  |
                                  v
                  [Output: Canonical Player Record]
```

#### Pipeline Stages:
1. **Tier 1 — Player UUID Lookup:** Directly resolves NV Play UUID against `id_map`. Supports `prefer_nv_play_name` flag when statistical reports require scoreboard-recognized names rather than legal registration names.
2. **Tier 2 — Contextual Disambiguation (`KNOWN_DUPLICATES`):** Prevents duplicate player collisions (e.g., players with identical names in different clubs, or generational duplicates like Father/Son). When a resolved name matches `KNOWN_DUPLICATES`, it is automatically qualified with the club acronym: `f"{canonical_name} ({short_club})"`.
3. **Tier 3 — Fast-Path Exact Name Indexing:** Uses cached pre-computed dictionary indexing (`_ID_MAP_INDEX_CACHE`) over normalized names to bypass expensive string distance algorithms for 95%+ of lookups.
4. **Tier 4 — Heuristic Fuzzy Matching:** When exact match is missing, applies `thefuzz.process.extractOne()` with `token_sort_ratio >= 90`, followed by fixture and club verification against `comb_context` (team context + match group).
5. **Tier 5 — Celtic Casing Normalization:** Formats surnames through `fix_celtic_casing()` to ensure proper capitalisation of prefixes (`Mc`, `Mac`, `O'`, `Fitz`).

### 3.2 Standardized Name Sorting (`get_name_sort_key`)
To guarantee uniform alphabetical ordering across financial statements, club invoicing, and player registries:
- Defined centrally in [`engine.get_name_sort_key()`](file:///c:/Users/Wylie/Downloads/Python%20Environment/engine.py#L3288) and mirrored cleanly in [`finance_app.py`](file:///c:/Users/Wylie/Downloads/Python%20Environment/finance_app.py#L1161).
- Returns a stable `Tuple[str, str]` representing `(surname.lower(), christian_name.lower())`.
- Accurately splits structured series (`Last Name`, `First Name`, `Full Name`), handles single-word names, and strips extraneous whitespace/parenthetical notes.

---

## 4. Domain & Session-State Scoping Architecture

### 4.1 The Streamlit State Collision Problem
In Streamlit applications featuring multi-domain navigation (Men's, Women's, Midweek) and multi-club contexts, static widget keys create state leaks where selecting an option or inputting text in one view pollutes or overrides controls in another view.

### 4.2 Scoping Hierarchy

```
Global Scope
 │
 ├── Domain-Scoped Keys: f"{prefix}_{domain}"
 │    ├── e.g. registry_club_select_{domain}
 │    ├── e.g. doc_reg_{domain}
 │    └── e.g. avg_bat_{domain}
 │
 └── Compound-Scoped Keys: f"{prefix}_{domain}_{selected_club}"
      ├── e.g. a10_enable_override_{domain}_{selected_club}
      ├── e.g. a10_replace_tier_{domain}_{selected_club}
      ├── e.g. a10_player_out_{domain}_{selected_club}
      ├── e.g. a10_player_in_select_{domain}_{selected_club}
      ├── e.g. a10_override_reason_{domain}_{selected_club}
      └── e.g. a10_save_override_{domain}_{selected_club}
```

### 4.3 Scoping Patterns by Module

#### [`app.py`](file:///c:/Users/Wylie/Downloads/Python%20Environment/app.py) (Starring Registry & Administrative Suite)
- **Master Club Selector:** `key=f"registry_club_select_{domain}"` ensures Men's and Women's views preserve their independent club selections across tab switches.
- **Emergency Roster Modification Controls:** All emergency override checkboxes, tier dropdowns, player substitution selectors, text fallbacks, explanation inputs, and save buttons use `f"..._{domain}_{selected_club}"`. Draft replacements for one club never bleed into another club's form.
- **Domain Switch Detection:**
  ```python
  if 'doc_last_domain' not in st.session_state or st.session_state.doc_last_domain != domain:
      st.session_state.player_search_active = False
      st.session_state.doc_last_domain = domain
  ```

#### [`stats_app.py`](file:///c:/Users/Wylie/Downloads/Python%20Environment/stats_app.py) (Statistics & Averages)
- **Competition & Club Filters:** Scoped as `key=f"club_summary_comp_{active_domain}"` and `key=f"club_summary_club_{active_domain}"`.
- **File Ingestion Selectors:** Scoped per domain: `f"avg_reg_{domain}"`, `f"avg_alias_{domain}"`, `f"avg_bat_{domain}"`, `f"avg_bowl_{domain}"`, `f"avg_league_{domain}"`, `f"avg_cup_{domain}"`.
- **Statistical Threshold Sliders:** Dynamic tier thresholds dynamically keyed per league tier (e.g., `wp_runs`, `s1_wick`, `j1_runs`).

#### [`finance_app.py`](file:///c:/Users/Wylie/Downloads/Python%20Environment/finance_app.py) (Invoicing & Fee Auditing)
- **Download Buttons:** Download triggers scoped per club: `key=f"dl_comp_excel_{safe_sel_club}"`, `key=f"dl_comp_docx_{safe_sel_club}"`.
- **Audit File Inputs:** Scoped per domain: `f"unreg_reg_{domain}"`, `f"club_fines_bat_{domain}"`, `f"club_fines_revenue_{domain}"`.
- **Session Output Caching:** Audit outputs stored under explicit session keys (`st.session_state['audit_outputs']`, `st.session_state['generated_invoices']`) to prevent loss on pagination or rerenders.

---

## 5. Cricket Statistics Domain Rules & Math

All core statistical calculations are centralized in [`engine.py`](file:///c:/Users/Wylie/Downloads/Python%20Environment/engine.py) to ensure zero mathematical drift across applications:

### 5.1 Batting Averages (`calculate_batting_average`)
- **Formula:** $\text{Average} = \frac{\text{Total Runs}}{\text{Innings} - \text{Not Outs}}$
- **Zero Dismissal Rule:** If dismissals equal 0, the function **must never divide by zero**. It returns the total runs formatted with an asterisk (e.g. `'120*'`).
- **Formatting:** Numeric averages are formatted to exactly 2 decimal places (`f"{avg:.2f}"`).

### 5.2 Bowling Economy Rates (`calculate_economy_rate`)
- **Base-6 Fractional Overs Conversion:** In cricket, $3.2$ overs represents $3$ overs and $2$ legal balls ($3 + \frac{2}{6} = 3.333$ overs).
- **Formula:** $\text{Economy} = \frac{\text{Runs Conceded}}{\text{Balls Bowled} / 6.0}$
- Zero balls bowled returns `'0.00'`.

### 5.3 Bowling Averages (`calculate_bowling_average`)
- **Formula:** $\text{Average} = \frac{\text{Runs Conceded}}{\text{Wickets Taken}}$
- Zero wickets returns standard NCU dash `'-'`.

### 5.4 Best Bowling Sorting Key (`bb_sort_key`)
- Parses performance strings like `'5-24'` or `'3/15'` into sortable tuples `(wickets, -runs)`.
- Guarantees that higher wickets rank first, with fewer runs conceded breaking ties.

### 5.5 Match Format Classification (`classify_match_type`)
- Standardized classification routine in `engine.py:2598`.
- Segregates fixtures into `'League'`, `'Cup'`, `'T20'`, `'Irish'`, or `'Midweek League'`.
- Prevents Irish/Evara national cup fixtures from contaminating local NCU league/cup average calculations.

---

## 6. Developer Guidelines & Maintenance Workflow

### 6.1 Pre-Modification Backup Protocol (Rule 0)
Before making changes to any `.py` or configuration file, developers and automated agents must run a pre-flight backup compressing all active `.py` files into the root directory:
- **Format:** `NCU Cricket Hub app (Backup - YYYY-MM-DD_HHMM).zip`
- **Execution:** Silently generated via background inline script or shell command.

### 6.2 Regression Testing
The workspace contains a comprehensive test suite in `tests/`:
```powershell
python -m pytest tests/
```
- **Validation Standard:** All 122 unit tests across `test_app.py`, `test_engine.py`, `test_starring_rules.py`, `test_stats_app.py`, `test_audits.py`, `test_duplicates.py`, and `test_disambiguation.py` must pass with 0 failures and 0 errors before committing changes.

### 6.3 Code Quality & Style
- **Python Type Hints:** Explicit type annotations are required on all functions (e.g., `def calculate_average(runs: int, dismissals: int) -> float:`).
- **Structured Docstrings:** Include concise summaries specifying `Inputs`, `Outputs`, and `Helper Apps`.
- **No Dead Code:** Eliminate temporary debugging `print()` statements and commented-out legacy code prior to commit.
