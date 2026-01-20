#%%##
# Combine multiple CSV files from FullCAM 2024 into a single Excel workbook with each CSV as a separate sheet.
# FullCAM 2024 CSV have different header names to 2016

#Part 1 - Combine CSVs into Excel workbook
''''''''''''

import os
import glob
import pandas as pd
from datetime import datetime

# === Use existing folder path set earlier ===
#inputDir = output_folder  # Use the output folder where CSVs were saved
inputDir = r"C:\Users\GeorginaDoyle\github\EP_FullCAM\COALARA\Sch4_Project\FC24_Projectcsv"  # Change this to your directory
output_folder = inputDir  # You can change this if needed
os.chdir(inputDir)

# === Gather all CSV files ===
extension = 'csv'
filenames = glob.glob(f'*.{extension}')
file_dict = {}

# === Process each CSV file ===
for f in filenames:
    try:
        df = pd.read_csv(f)

        if {'Year (yr)', 'Month (mo)', 'Day of month (day)'}.issubset(df.columns):
            # Create and format Date column: "02/06/2022"
            df['Date'] = pd.to_datetime(
                df[['Year (yr)', 'Month (mo)', 'Day of month (day)']].rename(
                    columns={
                        'Year (yr)': 'year',
                        'Month (mo)': 'month',
                        'Day of month (day)': 'day'
                    }
                ),
                errors='coerce'
            )
            # Format as "DD/MM/YYYY" (e.g., '02/06/2022')
            df['Date'] = df['Date'].dt.strftime('%d/%m/%Y').astype(str)

            # Reorder and select desired columns
            desired_columns = [
                'Date',
                'C mass of trees  (tC/ha)',
                'CH4 emitted due to fire (tCH4/ha)',
                'C mass of forest debris  (tC/ha)',
                'C mass of forest products  (tC/ha)',
                'N2O emitted due to fire (tN2O/ha)'
            ]

            # Check if all required columns exist
            missing_cols = [col for col in desired_columns[1:] if col not in df.columns]
            if missing_cols:
                print(f"⚠️ Skipping {f}: missing columns: {missing_cols}")
                continue

            df = df[desired_columns]

            # Truncate sheet name to max 31 characters
            sheet_name = f.replace('.csv', '')[:31]
            file_dict[sheet_name] = df
        else:
            print(f"⚠️ Skipping {f}: missing Year/Month/Day columns")
    except Exception as e:
        print(f"⚠️ Skipping {f}: error — {e}")

# === Write to Excel ===
today_str = datetime.today().strftime('%Y-%m-%d')
combined_filename = f'combined_output_{today_str}.xlsx'
combined_output_path = os.path.join(output_folder, combined_filename)

if file_dict:
    with pd.ExcelWriter(combined_output_path) as writer:
        for sheet_name, df in file_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"✅ Combined Excel file created at: {combined_output_path}")
else:
    print("⚠️ No valid data to write.")

#%%##
# Section 7- QC the final combined file

from datetime import datetime
import os
import pandas as pd

# Generate dated filename
today_str = datetime.today().strftime('%Y-%m-%d')
combined_filename = f'combined_output_{today_str}.xlsx'
combined_output_path = os.path.join(output_folder, combined_filename)

# Write to Excel
with pd.ExcelWriter(combined_output_path) as writer:
    for sheet_name, df in file_dict.items():
        # Separate Date column and the rest
        if 'Date' in df.columns:
            date_col = df[['Date']]
            other_cols = df.drop(columns=['Date'])

            # Convert all other columns to numeric (force number format)
            for col in other_cols.columns:
                other_cols[col] = pd.to_numeric(other_cols[col], errors='coerce')

            # Recombine
            cleaned_df = pd.concat([date_col, other_cols], axis=1)
        else:
            # No Date column? Just try converting everything
            cleaned_df = df.apply(pd.to_numeric, errors='coerce')

        # Save cleaned sheet
        cleaned_df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"✅ Combined Excel file created with numeric formatting: {combined_output_path}")


#%%###
#PART 2 - COMBINE TWO EXCEL WORKBOOKS SIDE-BY-SIDE
# -*- coding: utf-8 -*-
# Combine two multi-sheet Excel workbooks (Baseline vs Project) into one workbook.
# For each matching sheet name, merge on 'Date' and output side-by-side metrics with labels.
#%%##
from pathlib import Path
from datetime import datetime
import pandas as pd
import re
import numpy as np
#%%##
# ---------- INPUTS ----------
baseline_path = r"C:\Users\GeorginaDoyle\github\EP_FullCAM\COALARA\Sch4_Baseline\FC24\combined_output_2025-11-13.xlsx"
project_path = r"C:\Users\GeorginaDoyle\github\EP_FullCAM\COALARA\Sch4_Project\FC24_Projectcsv\combined_output_2025-11-13.xlsx"

output_folder = r"C:\Users\GeorginaDoyle\github\EP_FullCAM\outputs"
output_name_prefix = "FC24_combined_output"

LABEL_A = "Baseline"
LABEL_B = "Project"
DATE_FMT = "%d/%m/%Y"  # dates already in this format
LAST_DAY = {
    1: 31, 2: 28, 3: 31, 4: 30, 5: 31, 6: 30,
    7: 31, 8: 31, 9: 30, 10: 31, 11: 30, 12: 31
}
# Canonical metric names and order for final output
METRIC_COLS = [
    "C mass of trees (tC/ha)",
    "CH4 emitted due to fire (tCH4/ha)",
    "C mass of forest debris (tC/ha)",
    "C mass of forest products (tC/ha)",
    "N2O emitted due to fire (tN2O/ha)",
]
#%%##
# ---------- HELPERS ----------
def _norm(s: str) -> str:
    """Collapse internal whitespace and strip."""
    return re.sub(r"\s+", " ", str(s)).strip()

def _normalise_headers(df: pd.DataFrame) -> pd.DataFrame:
    """Normalise headers and map known spacing variants to canonical names."""
    df = df.copy()
    df.columns = [_norm(c) for c in df.columns]
    ren = {}
    # Fix known double-space variants -> canonical
    map_variants = {
        "C mass of trees  (tC/ha)": "C mass of trees (tC/ha)",
        "C mass of forest debris  (tC/ha)": "C mass of forest debris (tC/ha)",
        "C mass of forest products  (tC/ha)": "C mass of forest products (tC/ha)",
    }
    for c in df.columns:
        if c in map_variants:
            ren[c] = map_variants[c]
    if ren:
        df = df.rename(columns=ren)
    return df
def _select_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """Return DataFrame with exactly METRIC_COLS in that order (missing -> NaN)."""
    out = {}
    for c in METRIC_COLS:
        if c in df.columns:
            out[c] = pd.to_numeric(df[c], errors="coerce")
        else:
            out[c] = pd.Series(np.nan, index=df.index, name=c)
    return pd.DataFrame(out)


def _coerce_to_month_end(date_series: pd.Series) -> pd.Series:
    dt = pd.to_datetime(date_series, errors="coerce", dayfirst=True)
    y = dt.dt.year
    m = dt.dt.month
    d = m.map(LAST_DAY)
    return pd.to_datetime(dict(year=y, month=m, day=d), errors="coerce")


def _prepare_side(df: pd.DataFrame, label: str) -> pd.DataFrame:
    df = _normalise_headers(df)
    if "Date" not in df.columns:
        raise ValueError("Input sheet missing 'Date' column.")

    # --- NORMALISE DATE TO LAST DAY OF MONTH ---
    raw_str = df["Date"].astype(str)
    key_dt  = _coerce_to_month_end(raw_str)      # merge key
    date_out = key_dt.dt.strftime(DATE_FMT)      # final display

    metrics = _select_metrics(df)
    metrics.columns = [f"{label} - {c}" for c in metrics.columns]

    out = pd.concat([
        key_dt.rename("_key"),
        date_out.rename("Date"),
        metrics
    ], axis=1)

    # Drop rows with all metrics NaN
    metric_cols = metrics.columns.tolist()
    out = out.dropna(how="all", subset=metric_cols)

    # Deduplicate on month-end key (keep last)
    out = out.drop_duplicates(subset=["_key"], keep="last").reset_index(drop=True)

    return out

def _merge_two_frames(a_df: pd.DataFrame, b_df: pd.DataFrame) -> pd.DataFrame:
    merged = a_df.merge(b_df, on="_key", how="outer", suffixes=("_a", "_b"))
    merged = merged.sort_values("_key").reset_index(drop=True)

    # Rebuild canonical Date column (already month-end)
    merged["Date"] = merged["_key"].dt.strftime(DATE_FMT)

    a_cols = [c for c in merged.columns if c.startswith("Baseline - ")]
    b_cols = [c for c in merged.columns if c.startswith("Project - ")]

    merged = merged[["Date"] + a_cols + b_cols]

    for c in a_cols + b_cols:
        merged[c] = pd.to_numeric(merged[c], errors="coerce")

    return merged


#%%##
# ---------- MAIN ----------
def main():
    out_dir = Path(output_folder)
    out_dir.mkdir(parents=True, exist_ok=True)
    xls_a = pd.ExcelFile(baseline_path)
    xls_b = pd.ExcelFile(project_path)

    sheets_a = set(xls_a.sheet_names)
    sheets_b = set(xls_b.sheet_names)
    common = sorted(sheets_a & sheets_b)
    missing_a = sorted(sheets_b - sheets_a)
    missing_b = sorted(sheets_a - sheets_b)

    if not common:
        raise SystemExit("No matching sheet names between the two workbooks.")

    ts = datetime.now().strftime("%Y-%m-%d")
    out_xlsx = out_dir / f"{output_name_prefix}_{ts}.xlsx"

    with pd.ExcelWriter(out_xlsx) as xw:
        for name in common:
            try:
                df_a = pd.read_excel(baseline_path, sheet_name=name)
                df_b = pd.read_excel(project_path, sheet_name=name)

                a_pre = _prepare_side(df_a, LABEL_A)
                b_pre = _prepare_side(df_b, LABEL_B)

                merged = _merge_two_frames(a_pre, b_pre)
                merged.to_excel(xw, sheet_name=name[:31], index=False)
                print(f"✓ Wrote sheet: {name}")
            except Exception as e:
                print(f"⚠️ Skipped '{name}': {e}")

    if missing_a:
        print(f"Note: sheets only in Project workbook (no Baseline match): {missing_a}")
    if missing_b:
        print(f"Note: sheets only in Baseline workbook (no Project match): {missing_b}")

    print(f"\n✅ Combined Excel created: {out_xlsx}")

if __name__ == "__main__":
    main()

# %%
