# Antigravity Workspace Guidelines & Rules

## Excel File Formatting Protocol (MANDATORY)
Every time an Excel workbook (\.xlsx\) is created, modified, or updated in this workspace, apply the following 3 formatting steps to every worksheet:

1. **Highlight the first row as the header row**:
   - Set font to **bold**: \cell.font = Font(..., bold=True)   - Apply a soft accent fill: \cell.fill = PatternFill(start_color='FFFFE0', end_color='FFFFE0', fill_type='solid')\ (Light Yellow)
2. **Freeze the top row**:
   - Set freeze panes at row 2: \ws.freeze_panes = 'A2'\ (or \worksheet.freeze_panes(1, 0)\) so the header row remains visible when scrolling.
3. **Make every column exactly the max width of the widest entry**:
   - Measure the longest entry in each column (header and all data rows) and set width with +2 padding:
     \ws.column_dimensions[col_letter].width = max(max_length + 2, 10)     (Ensures text, UUIDs, and numbers are fully readable with no clipping or \###\ truncation).

### Data Integrity Safeguards:
- **Best Bowling / Score Format Preservation**:
  Columns containing cricket bowling figures (e.g. '1-21'\, '5/24'\) must maintain text format (\@\) so Excel never misinterprets them as calendar dates.
- **Ghost Row Prevention**:
  Never save workbooks with millions of blank styled rows; prune empty rows to the actual data boundaries.
- **Regression Verification**:
  Ensure all unit tests pass (\python -m pytest tests/\) whenever modifying code or data structures.
