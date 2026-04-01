# tools/compare_output.py
# Compares VBA macro output vs Python module output.
#
# Usage:
#   python tools/compare_output.py \
#     --vba   path/to/vba_output.xlsm \
#     --python path/to/python_output.xlsm \
#     --output path/to/differences.xlsx   (optional)

import argparse
from pathlib import Path

import pandas as pd

SHEET_NAME = "Testergebnisse"
HEADER_ROW = 4   # row 5 in Excel = index 4 in pandas

def load_sheet(path: Path, sheet_name=SHEET_NAME, header=0,) -> pd.DataFrame:
    """Loads Testergebnisse sheet from an Excel file."""
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    return pd.read_excel(
        path,
        sheet_name=sheet_name,
        header=header,
        engine="openpyxl"
    )

def normalize(series: pd.Series) -> pd.Series:
    """Treats NaN, 'nan', 'none', '' as equivalent empty values."""
    return (
        series
        .fillna("")
        .astype(str)
        .str.replace(
            {
                "_x000D_": "",
                "\r\n": "\n",
                "\r": "\n",
                "nan": "",
                "none": "",
                "<na>": ""
            },
            regex=False,
        )
    )

def compare_outputs(
    df_vba:    pd.DataFrame,
    df_python: pd.DataFrame,
    key:       str,
) -> pd.DataFrame:
    common_cols = [
        col for col in df_vba.columns
        if col in df_python.columns and col != key
    ]

    print(f"\n{'='*60}")
    print(f"Rows:    VBA={len(df_vba)} | Python={len(df_python)}")
    print(f"Columns: {len(common_cols)} compared")
    print(f"{'='*60}\n")

    differences = []
    total_diffs = 0

    for col in common_cols:
        vba_col    = normalize(df_vba[col])
        python_col = normalize(df_python[col])

        diff_mask  = vba_col != python_col
        diff_count = diff_mask.sum()
        total_diffs += diff_count

        if diff_count == 0:
            print(f"✅ {col}")
        else:
            print(f"❌ {col} — {diff_count} difference(s)")

            for idx in df_vba[diff_mask].index:
                req_id = df_vba.loc[idx, key] if key in df_vba.columns else idx
                differences.append({
                    "Row":           idx ,
                    f"{key}":        req_id,
                    "Column":        col,
                    "VBA":           str(df_vba.loc[idx, col]),
                    "Python":        str(df_python.loc[idx, col]),
                })

    print(f"\n{'='*60}")
    print(f"Total differences: {total_diffs}")
    print(f"{'='*60}\n")

    return pd.DataFrame(differences)

def main():
    parser = argparse.ArgumentParser(
        description="Compare VBA macro output vs Python module output."
    )
    parser.add_argument("--vba",    required=True, help="Path to VBA output Excel file")
    parser.add_argument("--python", required=True, help="Path to Python output Excel file")
    parser.add_argument("--key",    default="ID der Anforderung",
                        help="Requirement ID column name (default: 'ID der Anforderung')")
    parser.add_argument("--output", default=None,
                        help="Optional: path to write differences Excel file")
    parser.add_argument("--sheet", default=SHEET_NAME,
                        help="Optional: Excel worksheet name to compare")
    args = parser.parse_args()

    vba_path    = Path(args.vba)
    python_path = Path(args.python)

    print(f"VBA file:    {vba_path.name}")
    print(f"Python file: {python_path.name}")
    print(f"Sheet:       {args.sheet} (header row {HEADER_ROW + 1})")

    df_vba    = load_sheet(vba_path)
    df_python = load_sheet(python_path)

    diff_df = compare_outputs(df_vba, df_python, args.key)

    if args.output and not diff_df.empty:
        out_path = Path(args.output)
        diff_df.to_excel(out_path, index=False)
        print(f"Differences written to: {out_path}")
    elif diff_df.empty:
        print("✅ No differences found — outputs are identical.")

if __name__ == "__main__":
    main()
