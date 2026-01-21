# -*- coding: utf-8 -*-
"""
Shared helpers for merging Baseline/Project Schedule 4 workbooks.

Features:
- Normalizes header spacing.
- Fills missing metrics with NaN.
- Outer-joins and orders Baseline then Project.
- Tolerates multiple date formats by parsing with dayfirst/iso fallback.
"""
from __future__ import annotations

from datetime import datetime
import re
from pathlib import Path
from typing import Iterable, List, Optional

import numpy as np
import pandas as pd

# Canonical metric names and order for final output
METRIC_COLS: List[str] = [
    "C mass of trees (tC/ha)",
    "CH4 emitted due to fire (tCH4/ha)",
    "C mass of forest debris (tC/ha)",
    "C mass of forest products (tC/ha)",
    "N2O emitted due to fire (tN2O/ha)",
]

LAST_DAY = {
    1: 31,
    2: 28,
    3: 31,
    4: 30,
    5: 31,
    6: 30,
    7: 31,
    8: 31,
    9: 30,
    10: 31,
    11: 30,
    12: 31,
}


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
        "C mass of forest litter and deadwood  (tC/ha)": "C mass of forest debris (tC/ha)",
        "C mass of forest litter and deadwood (tC/ha)": "C mass of forest debris (tC/ha)",
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


def _parse_dates(date_series: pd.Series) -> pd.Series:
    """Parse dates with dayfirst first, then fall back to month-first/ISO."""
    s = date_series.astype(str).str.strip()
    dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
    mask = dt.isna()
    if mask.any():
        dt2 = pd.to_datetime(s[mask], errors="coerce", dayfirst=False)
        dt.loc[mask] = dt2
    return dt


def _build_date_key(date_series: pd.Series, normalize_month_end: bool) -> pd.Series:
    dt = _parse_dates(date_series)
    if normalize_month_end:
        y = dt.dt.year
        m = dt.dt.month
        d = m.map(LAST_DAY)
        dt = pd.to_datetime(dict(year=y, month=m, day=d), errors="coerce")
    return dt


def _format_key_dates(key: pd.Series, date_fmt: str) -> pd.Series:
    def _fmt(v) -> str:
        if isinstance(v, pd.Timestamp):
            return v.strftime(date_fmt)
        if isinstance(v, datetime):
            return v.strftime(date_fmt)
        return str(v)

    return key.map(_fmt)


def _prepare_side(
    df: pd.DataFrame,
    label: str,
    date_fmt: str,
    normalize_month_end: bool,
) -> pd.DataFrame:
    df = _normalise_headers(df)
    if "Date" not in df.columns:
        raise ValueError("Input sheet missing 'Date' column.")

    date_str = df["Date"].astype(str)
    key_dt = _build_date_key(date_str, normalize_month_end=normalize_month_end)
    key = key_dt.where(key_dt.notna(), date_str)

    metrics = _select_metrics(df)
    metrics.columns = [f"{label} - {c}" for c in metrics.columns]

    out = pd.concat(
        [
            key.rename("_key"),
            key_dt.rename("_sort"),
            metrics,
        ],
        axis=1,
    )

    # Drop rows with all metrics NaN
    metric_cols = metrics.columns.tolist()
    out = out.dropna(how="all", subset=metric_cols)

    # Deduplicate on key (keep last)
    out = out.drop_duplicates(subset=["_key"], keep="last").reset_index(drop=True)
    return out


def _merge_two_frames(
    a_df: pd.DataFrame,
    b_df: pd.DataFrame,
    date_fmt: str,
    label_a: str,
    label_b: str,
) -> pd.DataFrame:
    merged = a_df.merge(b_df, on="_key", how="outer", suffixes=("_a", "_b"))

    # Order by parsed date when available
    merged["_sort"] = merged["_sort_a"].where(merged["_sort_a"].notna(), merged["_sort_b"])
    merged = merged.drop(columns=["_sort_a", "_sort_b"]).sort_values("_sort").drop(columns="_sort")

    # Rebuild Date column from key
    merged["Date"] = _format_key_dates(merged["_key"], date_fmt)

    a_cols = [f"{label_a} - {c}" for c in METRIC_COLS if f"{label_a} - {c}" in merged.columns]
    b_cols = [f"{label_b} - {c}" for c in METRIC_COLS if f"{label_b} - {c}" in merged.columns]

    merged = merged[["Date"] + a_cols + b_cols].reset_index(drop=True)

    for c in a_cols + b_cols:
        merged[c] = pd.to_numeric(merged[c], errors="coerce")
    return merged


def merge_workbooks(
    baseline_path: str,
    project_path: str,
    output_folder: str,
    output_name_prefix: str,
    label_a: str = "Baseline",
    label_b: str = "Project",
    date_fmt: str = "%d/%m/%Y",
    normalize_month_end: bool = False,
) -> Path:
    """Merge two multi-sheet workbooks by sheet name with side-by-side metrics."""
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

                a_pre = _prepare_side(
                    df_a, label=label_a, date_fmt=date_fmt, normalize_month_end=normalize_month_end
                )
                b_pre = _prepare_side(
                    df_b, label=label_b, date_fmt=date_fmt, normalize_month_end=normalize_month_end
                )

                merged = _merge_two_frames(
                    a_pre, b_pre, date_fmt=date_fmt, label_a=label_a, label_b=label_b
                )
                merged.to_excel(xw, sheet_name=name[:31], index=False)
                print(f'Wrote sheet: {name}')
            except Exception as e:
                print(f"Skipped '{name}': {e}")

    if missing_a:
        print(f"Note: sheets only in Project workbook (no Baseline match): {missing_a}")
    if missing_b:
        print(f"Note: sheets only in Baseline workbook (no Project match): {missing_b}")

    print(f"\nCombined Excel created: {out_xlsx}")
    return out_xlsx
