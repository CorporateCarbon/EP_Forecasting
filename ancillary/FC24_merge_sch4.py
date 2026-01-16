
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
baseline_path = r"C:\Users\GeorginaDoyle\github\EP_FullCAM\baseline_combined.xlsx"
project_path = r"C:\Users\GeorginaDoyle\github\EP_FullCAM\project_combined.xlsx"

output_folder = r"C:\Users\GeorginaDoyle\github\EP_FullCAM\outputs"
output_name_prefix = "combined_output"

LABEL_A = "Baseline"
LABEL_B = "Project"
DATE_FMT = "%d/%m/%Y"  # dates already in this format

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

def _prepare_side(df: pd.DataFrame, label: str) -> pd.DataFrame:
    df = _normalise_headers(df)
    if "Date" not in df.columns:
        raise ValueError("Input sheet missing 'Date' column.")
    # Ensure Date is string in DD/MM/YYYY, and build sortable helper
    date_str = df["Date"].astype(str)
    sort_dt = pd.to_datetime(date_str, errors="coerce", dayfirst=True)
    metrics  = _select_metrics(df)
    metrics.columns = [f"{label} - {c}" for c in metrics.columns]
    out = pd.concat([date_str.rename("Date"), sort_dt.rename("_sort") , metrics], axis=1)
    # Drop rows with all metrics NaN
    metric_cols = metrics.columns.tolist()
    out = out.dropna(how="all", subset=metric_cols)
    # Deduplicate by Date (keep last)
    out = out.drop_duplicates(subset=["Date"], keep="last").reset_index(drop=True)
    return out

def _merge_two_frames(a_df: pd.DataFrame, b_df: pd.DataFrame) -> pd.DataFrame:
    # Add suffixes so overlapping columns (_sort) are preserved
    merged = a_df.merge(b_df, on="Date", how="outer", suffixes=("_a", "_b"))

    # Combine the sortable date helpers (_sort_a, _sort_b)
    merged["_sort"] = merged["_sort_a"].where(merged["_sort_a"].notna(), merged["_sort_b"])
    merged = merged.drop(columns=["_sort_a", "_sort_b"]).sort_values("_sort").drop(columns="_sort").reset_index(drop=True)

    # Order columns: Date, Baseline metrics, Project metrics
    a_cols = [c for c in merged.columns if c.startswith("Baseline - ")]
    b_cols = [c for c in merged.columns if c.startswith("Project - ")]
    merged = merged[["Date"] + a_cols + b_cols]

    # Force numeric types for metrics
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
