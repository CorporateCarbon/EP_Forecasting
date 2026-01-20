
# Combine matched CSVs from two folders into one Excel (one sheet per base filename)
# Side-by-side columns: Date, <metrics from folder A>, <metrics from folder B>
#ombines pairs of matching CSVs from two folders into a single Excel workbook (one sheet per file); it builds Date from Year + Month (or Step In Year) and forces month‑end day; only includes the defined metric set; hard errors if none of the metrics exist.
# 
# %%##
import os
import glob
from pathlib import Path
import pandas as pd
from datetime import datetime

# ---- USER PARAMS -------------------------------------------------------------
baseline_folder = r"C:\Users\GeorginaDoyle\github\EP_FullCAM\Coalara_Schedule4_NewModel\Baseline_2022\outputs"       # e.g. output_baseline_folder
project_folder = r"C:\Users\GeorginaDoyle\Corporate Carbon Pty Ltd\Corporate Carbon - 04. CARBON DELIVERY\08. ERF Projects\Coalara Park Australian Sandalwood Plantation Project - AT\FullCAM\250904_Schedule4_FullCAM2016\ProjectScenario_FullCAM\csv_output"       # e.g. output_folder_B
output_folder = r"C:\Users\GeorginaDoyle\github\EP_FullCAM\outputs"
output_name_prefix = "combined_output"
DATE_FMT = "%d/%m/%Y"                # Example file shows DD/MM/YYYY. Use '%Y, %b %d' if preferred.

# Which columns to output (in this order) for each side
METRIC_COLS = [
    'C mass of trees (tC/ha)',
    'CH4 emitted due to fire (tCH4/ha)',
    'C mass of forest debris (tC/ha)',
    'C mass of forest products (tC/ha)',
    'N2O emitted due to fire (tN2O/ha)',
]

LABEL_A = "Baseline"
LABEL_B = "Project"

# Fixed month -> last day mapping (Feb=28 per spec)
LAST_DAY = {1:31, 2:28, 3:31, 4:30, 5:31, 6:30, 7:31, 8:31, 9:30, 10:31, 11:30, 12:31}
# -----------------------------------------------------------------------------


def _normalise_headers(df: pd.DataFrame) -> pd.DataFrame:
    # Strip spaces, fix casing, and standardise 'Step In Year' -> 'Month'
    newcols = {}
    for c in df.columns:
        c_clean = c.strip()
        if c_clean.lower() == "step in year":
            newcols[c] = "Month"
        else:
            newcols[c] = c_clean
    return df.rename(columns=newcols)


def _build_date(df: pd.DataFrame) -> pd.Series:
    if "Year" not in df.columns or "Month" not in df.columns:
        raise ValueError("Missing 'Year'/'Month' columns after header normalisation.")
    y = pd.to_numeric(df["Year"], errors="coerce").astype("Int64")
    m = pd.to_numeric(df["Month"], errors="coerce").astype("Int64")
    d = m.map(LAST_DAY).astype("Int64")
    return pd.to_datetime(dict(year=y, month=m, day=d), errors="coerce")


def _prepare_side(df: pd.DataFrame, label: str) -> pd.DataFrame:
    df = _normalise_headers(df).copy()
    # Build Date (last day of month) and format for display
    date = _build_date(df)
    out = pd.DataFrame({"Date": date.dt.strftime(DATE_FMT)})

    # Keep requested metrics that exist, then prefix all metric headers with label
    present = [c for c in METRIC_COLS if c in df.columns]
    if not present:
        raise ValueError(f"None of the required metric columns are present for side '{label}'.")
    metrics_prefixed = df[present].add_prefix(f"{label} - ")
    out = out.join(metrics_prefixed)

    # Drop invalid rows
    out = out[~out["Date"].isna()]
    out = out.dropna(how="all", subset=list(metrics_prefixed.columns))
    # Deduplicate Date (keep last if multiple rows per month)
    out = out.drop_duplicates(subset=["Date"], keep="last").reset_index(drop=True)
    return out


def combine_two_csvs_side_by_side(csv_a: Path, csv_b: Path) -> pd.DataFrame:
    a = pd.read_csv(csv_a)
    b = pd.read_csv(csv_b)

    a_pre = _prepare_side(a, LABEL_A)
    b_pre = _prepare_side(b, LABEL_B)

    # Outer-join on Date; columns are already label-prefixed so no name collisions
    merged = (
        a_pre.merge(b_pre, on="Date", how="outer")
        .assign(_sortdate=lambda d: pd.to_datetime(d["Date"], errors="coerce", dayfirst=True))
        .sort_values("_sortdate")
        .drop(columns="_sortdate")
        .reset_index(drop=True)
    )

    # Order columns: Date, Baseline metrics (in METRIC_COLS order), Project metrics (same order)
    a_cols = [f"{LABEL_A} - {c}" for c in METRIC_COLS if f"{LABEL_A} - {c}" in merged.columns]
    b_cols = [f"{LABEL_B} - {c}" for c in METRIC_COLS if f"{LABEL_B} - {c}" in merged.columns]
    ordered_cols = ["Date"] + a_cols + b_cols
    merged = merged[ordered_cols]
    return merged

#%%##
def main():
    Path(output_folder).mkdir(parents=True, exist_ok=True)

    def map_files(folder):
        return {Path(p).stem: Path(p) for p in glob.glob(str(Path(folder) / "*.csv"))}

    a_map = map_files(baseline_folder)
    b_map = map_files(project_folder)

    common = sorted(set(a_map).intersection(b_map))
    if not common:
        raise SystemExit("No matching CSV filenames across the two folders.")

    ts = datetime.now().strftime("%Y-%m-%d")
    out_xlsx = Path(output_folder) / f"{output_name_prefix}_{ts}.xlsx"

    with pd.ExcelWriter(out_xlsx) as xw:
        for stem in common:
            try:
                df_out = combine_two_csvs_side_by_side(a_map[stem], b_map[stem])
                sheet_name = stem[:31]  # Excel sheet name limit
                df_out.to_excel(xw, sheet_name=sheet_name, index=False)
                print(f"✓ Wrote sheet: {sheet_name}")
            except Exception as e:
                print(f"⚠️  Skipped '{stem}': {e}")

    print(f"\n✅ Combined Excel created: {out_xlsx}")


if __name__ == "__main__":
    main()
# %%
