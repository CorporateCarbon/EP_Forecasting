#!/usr/bin/env python3
r"""
Combine all CSVs in a directory into a single Excel workbook with one sheet per file.

Default behavior:
  - Looks in the current working directory (or a provided directory).
  - Expects columns: Year, Month, Day (any casing); optionally drops 'Dec. Year' if present.
  - Creates a 'Date' column formatted like "YYYY, Mon DD".
  - Writes results to 'combined_output.xlsx' in the chosen directory.

Usage:
  # From the folder that contains your CSVs:
  python combine_csvs.py

  # Or specify a folder and/or output file:
  python combine_csvs.py "C:/path/to/csvs" -o my_combined.xlsx

  # Change the filename pattern (default: *.csv):
  python combine_csvs.py -p "*.CSV"
"""

import argparse
import sys
from pathlib import Path
import re
import pandas as pd


def sanitize_sheet_name(name: str) -> str:
    """Excel sheet names max 31 chars, cannot contain : \ / ? * [ ]"""
    cleaned = re.sub(r'[:\\/?*\[\]]', "_", name)
    return cleaned[:31]


def make_unique(name: str, existing: set) -> str:
    """Ensure the sheet name is unique by appending a counter if needed."""
    base = name
    suffix = 1
    while name in existing:
        tail = f"_{suffix}"
        name = (base[: 31 - len(tail)] + tail) if len(base) + len(tail) > 31 else base + tail
        suffix += 1
    existing.add(name)
    return name


def find_cols_case_insensitive(columns, target):
    lower_map = {c.lower(): c for c in columns}
    return lower_map.get(target.lower())


def build_date_column(df: pd.DataFrame) -> pd.Series:
    y_col = find_cols_case_insensitive(df.columns, "Year")
    m_col = find_cols_case_insensitive(df.columns, "Month")
    d_col = find_cols_case_insensitive(df.columns, "Day")
    if not all([y_col, m_col, d_col]):
        raise ValueError("Missing columns: Year, Month, Day")

    # Coerce to integers where possible, then back to strings for zero-padding
    y = pd.to_numeric(df[y_col], errors="coerce").astype("Int64").astype(str)
    m = pd.to_numeric(df[m_col], errors="coerce").astype("Int64").astype(str).str.zfill(2)
    d = pd.to_numeric(df[d_col], errors="coerce").astype("Int64").astype(str).str.zfill(2)

    dt = pd.to_datetime(y + m + d, format="%Y%m%d", errors="coerce")
    if dt.isna().any():
        bad_rows = int(dt.isna().sum())
        raise ValueError(f"Could not parse dates for {bad_rows} row(s). Check Year/Month/Day values.")
    return dt.dt.strftime("%Y, %b %d")


def process_file(csv_path: Path) -> pd.DataFrame:
    df = pd.read_csv(csv_path)
    df.insert(0, "Date", build_date_column(df))

    # Drop helper/date-construction columns if present
    to_drop = []
    for candidate in ["combine_date", "Year", "Month", "Day", "Dec. Year", "Dec Year"]:
        real = find_cols_case_insensitive(df.columns, candidate)
        if real:
            to_drop.append(real)
    if to_drop:
        df.drop(columns=to_drop, inplace=True, errors="ignore")
    return df


def main():
    parser = argparse.ArgumentParser(description="Combine CSVs into an Excel workbook (one sheet per file).")
    parser.add_argument("directory", nargs="?", default=".", help="Directory containing CSV files (default: .)")
    parser.add_argument("-o", "--output", default="combined_output.xlsx", help="Output Excel file name.")
    parser.add_argument("-p", "--pattern", default="*.csv", help='Glob pattern for files (default: "*.csv").')
    args = parser.parse_args()

    input_dir = Path(args.directory).expanduser().resolve()
    if not input_dir.is_dir():
        print(f"❌ Directory not found: {input_dir}", file=sys.stderr)
        sys.exit(1)

    csv_files = sorted(input_dir.glob(args.pattern))
    if not csv_files:
        print(f"⚠️  No files matched {args.pattern} in {input_dir}")
        sys.exit(0)

    output_path = (input_dir / args.output).resolve()

    file_dict = {}
    errors = []
    used_sheet_names = set()

    for csv_path in csv_files:
        try:
            df = process_file(csv_path)
            sheet_name = make_unique(sanitize_sheet_name(csv_path.stem), used_sheet_names)
            file_dict[sheet_name] = df
        except Exception as e:
            errors.append((csv_path.name, str(e)))

    if not file_dict:
        print("❌ No valid CSVs processed successfully.")
        if errors:
            print("\nErrors:")
            for name, msg in errors:
                print(f"  - {name}: {msg}")
        sys.exit(2)

    try:
        with pd.ExcelWriter(output_path) as writer:
            for sheet, df in file_dict.items():
                df.to_excel(writer, sheet_name=sheet, index=False)
    except Exception as e:
        print(f"❌ Failed to write Excel file: {e}", file=sys.stderr)
        sys.exit(3)

    print(f"✅ Wrote {len(file_dict)} sheet(s) to: {output_path}")
    if errors:
        print("\n⚠️  Some files were skipped due to errors:")
        for name, msg in errors:
            print(f"  - {name}: {msg}")


if __name__ == "__main__":
    main()
