# Antigravity Agent Configuration

## Active Model Strategies
- Primary: Gemini Flash 3.8 + Extended Thinking
- Failover (Exhausted Quota): Claude Sonnet + Extended Thinking

### 0. Pre-Modification Automation (Mandatory Backup)
- CRITICAL PRE-FLIGHT ACTION: BEFORE writing or patching any `.py` or configuration files in the workspace, you MUST programmatically generate a timestamped `.zip` backup of all python scripts.
- Execution Method: Silently run a background shell command or inline python snippet to compress all active `.py` files into a file structure in the root directory.
- Naming Convention: The backup file MUST be named exactly using this format: `NCU Cricket Hub app (Backup - YYYY-MM-DD_HHMM).zip` (substituting the current system date and time).
- Do not ask the user for confirmation to create the backup; execute it silently as an atomic precursor to any file modification task.


## Quota Optimization Rules
To maximize the remaining weekly limits for Claude/GPT models, all automated actions must adhere to these structural constraints:

### 1. Repository Scanning & File Ingestion
- Targeted reads only: Explicitly ask the user for permission before reading more than 3 SOURCE CODE files (`.py`) at once.
- Excel / Data Automation Allowed: The agent is fully permitted to read, parse, and modify `.xlsx` and `.csv` files automatically to maintain user workflows.
- CRITICAL EXCEL SAFETY: When reading `.xlsx` or `.csv` files, NEVER read the raw data rows into the chat context. Read only the table headers, structural schemas, or the specific rows required for the active script execution.
- Strictly ignore the following high-density folders:
  - `node_modules/`
  - `.git/`
  - `dist/` or `build/`

### 2. Terminal Executions & Error Handling
- Do NOT engage in recursive diagnostic loops. 
- Maximum loop depth: If a terminal command fails (e.g., test or build error), you are permitted exactly ONE auto-correction attempt. 
- If the second attempt fails, HALT immediately, print the log output to the chat window, and wait for human instruction. Do not repeatedly prompt yourself.

### 3. Execution Delegations
- **Claude Opus:** Restrict use exclusively to architectural outline design, schema definitions, and debugging complex logic blocks. 
- **Claude Sonnet:** Delegate all boilerplate generation, multi-file routine editing, dependency installs, and active file manipulation to Sonnet. 
- Force an auto-downgrade to Sonnet for file writing tasks after Opus has finalized a structural blueprint.

### 4. Media Constraints
- Deactivate automatic browser automation UI verification screenshots unless explicitly requested by the user.

### 5. Cricket Statistics Domain Rules
- **Batting Averages:** When calculating batting averages, the formula MUST divide Total Runs by (Innings minus Not Outs). If a player has 0 dismissals, return the Total Runs with an asterisk or handle as a string, never divide by zero.
- **Bowling Metrics:** Ensure bowling economy rate is calculated as `(Runs Conceded / Overs Bowled)`. Remember that 3.2 overs means 3 overs and 2 balls (which is 3.333 overs in decimal math). Convert fractional overs to base-6 before doing division.
- **Data Flags:** Keep track of string representations like "DNB" (Did Not Bat) or "sub" (Substitute). Never try to parse these strings as numeric integers during statistical aggregation loops.

### 6. Code Style and Maintenance
- **Type Hinting:** All new or refactored Python functions must include explicit Python type hints (e.g., `def calculate_average(runs: int, dismissals: int) -> float:`).
- **Docstrings:** Every modified function must contain a concise docstring explaining its inputs, output behavior, and which helper apps depend on it.
- **No Leftover Comments:** Do not leave commented-out "dead code" strings or temporary `print()` statements inside production modules after debugging is complete.

### 7. Post-Modification Verification
- **Automated Validation:** Immediately after writing a code patch and verifying the syntax, you are authorized to run `pytest` silently in the workspace terminal.
- **Regression Handling:** If any of the 45 unit tests fail after your edit, roll back the specific code change immediately using the backup ZIP or git checkout, print the test trace logs, and ask the user for guidance. Do not leave the codebase in a broken state.

### 8. Excel Generation and Aesthetic Design
- Follow the exact mandatory formatting protocols, text constraints, and color tokens defined inside `Gemini.md`.

