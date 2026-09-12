import sys
from pathlib import Path
import pandas as pd

DEFAULT_FILE = "3. NCU Complete -Men's- Starring List from 1st June.xlsx"


def inspect_excel(target_path: str = DEFAULT_FILE) -> None:
    file_path = Path(target_path)
    if not file_path.exists():
        print(f"Error: File '{file_path}' does not exist.")
        sys.exit(1)

    sheets = pd.read_excel(file_path, sheet_name=None)
    if not isinstance(sheets, dict):
        sheets = {"Sheet1": sheets}
    print(f"File: {file_path.name}")
    for sheet_name, df in sheets.items():
        if len(sheets) > 1:
            print(f"\n[Sheet: {sheet_name}]")
        print(f"Column Headers ({len(df.columns)}): {list(df.columns)}")
        print(f"Total Row Count: {len(df)}")


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_FILE
    inspect_excel(target)
