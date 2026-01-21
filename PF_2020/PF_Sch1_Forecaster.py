# -*- coding: utf-8 -*-
"""
Schedule 1 ACCU Forecaster (Anniversary Year) — Robust xlwings runner
--------------------------------------------------------------------
Assumptions (per your confirmation):
- Summary!D30 = Calculated ACCUs for the current reporting period (NOT carbon stock)
- Summary inputs:
    B11 = RP start date
    B12 = RP end date
    B13 = "FullCAM date" (must correspond to a date present in CEA01!A:A in many workbooks)
- Optional (if present/desired):
    Calculations!A28 = Project carbon stock total (for reference only)

Key design points:
- Builds anniversary reporting periods from PROJECT_START_DATE.
- For each RP end, picks the nearest available FullCAM date from CEA01!A:A within tolerance.
- Forces CalculateFullRebuild twice (some workbooks need this) with retries if D30 is blank.
- Outputs a CSV:
    RP#, RP Start, RP End, FullCAM Date used, Calculated ACCUs (D30), Cumulative ACCUs, Carbon Stock (optional)
"""
#%%##
from __future__ import annotations

import os
import csv
from dataclasses import dataclass
from datetime import datetime, date, timedelta
from typing import Optional, List, Tuple

import xlwings as xw
from dateutil.relativedelta import relativedelta


# ---------------------------
# USER SETTINGS (EDIT THESE)
# ---------------------------
EXCEL_PATH = r"C:\Users\GeorginaDoyle\github\EP_Forecasting\PF_2020\PF_SCh1_FC24_cp.xlsx"

# Project start date (anniversary-based reporting)
PROJECT_START_DATE = date(2022, 6, 25)  # <-- EDIT if needed

# Forecast horizon in years (typically 25 for crediting period, but set as required)
HORIZON_YEARS = 25

# Tolerance when matching RP end date to an available FullCAM date in CEA01!A:A
FULLCAM_TOLERANCE_DAYS = 5

# If your workbook requires dates as TEXT (leading apostrophe), keep True.
WRITE_DATES_AS_TEXT = True

# Recalc behaviour
FULL_REBUILD_TIMES_PER_ATTEMPT = 2
RECALC_RETRIES = 2  # retries if D30 is still blank

# Output CSV
CSV_OUT = os.path.join(
    os.path.dirname(EXCEL_PATH),
    f"Sch1_FC24_Forecast_{datetime.now():%Y%m%d-%H%M%S}.csv"
)

# Optional: read carbon stock total for reference (can set to None to disable)
READ_CARBON_STOCK = True
CARBON_STOCK_CELL = ("Calculations", "A28")  # sheet, cell


# ---------------------------
# Helper functions
# ---------------------------
def parse_excel_date(v) -> Optional[date]:
    """Convert Excel cell value to date."""
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def safe_float(v) -> Optional[float]:
    try:
        if v is None or v == "":
            return None
        return float(v)
    except Exception:
        return None


def write_date_cell(ws, addr: str, d: date):
    """Write a date into a cell in the format the workbook expects."""
    if WRITE_DATES_AS_TEXT:
        ws.range(addr).number_format = "@"
        ws.range(addr).value = "'" + d.strftime("%d/%m/%Y")
    else:
        ws.range(addr).number_format = "dd/mm/yyyy"
        ws.range(addr).value = datetime(d.year, d.month, d.day)


def calculate_full_rebuild(wb, times: int = 2):
    """Force Excel full rebuild recalculation (often needed for complex workbooks)."""
    for _ in range(times):
        wb.app.api.CalculateFullRebuild()


def load_fullcam_dates(ws_cea) -> List[date]:
    """Load CEA01 column A into a list of dates."""
    col = ws_cea.range("A1").expand("down").value
    if not isinstance(col, list):
        col = [col]
    dates: List[date] = []
    for v in col:
        d = parse_excel_date(v)
        if d:
            dates.append(d)
    return dates


def find_nearest_fullcam_date(fullcam_dates: List[date], target: date, tolerance_days: int) -> Optional[date]:
    """Find a FullCAM date within +/- tolerance_days of target; prefer exact match."""
    if not fullcam_dates:
        return None

    # Exact match first
    if target in set(fullcam_dates):
        return target

    best = None
    best_abs = None
    for d in fullcam_dates:
        diff = abs((d - target).days)
        if diff <= tolerance_days:
            if best is None or diff < best_abs:
                best = d
                best_abs = diff
    return best


def build_anniversary_periods(project_start: date, years: int) -> List[Tuple[date, date]]:
    """
    Anniversary reporting periods:
      RP1: start = project_start
           end   = project_start + 1 year - 1 day
      RP2: start = project_start + 1 year
           end   = project_start + 2 years - 1 day
      etc.
    """
    periods: List[Tuple[date, date]] = []
    rp_start = project_start
    for _ in range(years):
        rp_end = (rp_start + relativedelta(years=1)) - timedelta(days=1)
        periods.append((rp_start, rp_end))
        rp_start = rp_start + relativedelta(years=1)
    return periods


@dataclass
class ForecastRow:
    rp_index: int
    rp_start: date
    rp_end: date
    fullcam_date_used: date
    calculated_accus: Optional[float]   # Summary!D30
    cumulative_accus: Optional[float]   # running sum of calculated_accus (None-safe)
    carbon_stock_total: Optional[float] # optional reference value


# ---------------------------
# Main runner
# ---------------------------
def run_forecast() -> str:
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"Excel file not found: {EXCEL_PATH}")

    periods = build_anniversary_periods(PROJECT_START_DATE, HORIZON_YEARS)

    with xw.App(visible=True, add_book=False) as app:
        wb = app.books.open(EXCEL_PATH)

        # Sheets
        ws_sum = wb.sheets["Summary"]
        ws_cea = wb.sheets["CEA01"]

        # Preload FullCAM dates once
        fullcam_dates = load_fullcam_dates(ws_cea)
        if not fullcam_dates:
            raise RuntimeError("No FullCAM dates found in CEA01 column A.")

        rows: List[ForecastRow] = []
        cum = 0.0
        cum_has_value = False  # track whether we've had any non-null outputs

        for i, (rp_start, rp_end) in enumerate(periods, start=1):
            # Find a valid FullCAM date near RP end (this is the key NULL fix)
            fc_date = find_nearest_fullcam_date(fullcam_dates, rp_end, FULLCAM_TOLERANCE_DAYS)

            # If RP end isn't present, also try RP end +1 day (common if monthly table is month-end)
            if fc_date is None:
                fc_date = find_nearest_fullcam_date(fullcam_dates, rp_end + timedelta(days=1), FULLCAM_TOLERANCE_DAYS)

            # If still None, fall back to nearest to RP start (less ideal, but better than NULL cascade)
            if fc_date is None:
                fc_date = find_nearest_fullcam_date(fullcam_dates, rp_start, FULLCAM_TOLERANCE_DAYS)

            if fc_date is None:
                # Can't run this RP reliably
                rows.append(ForecastRow(
                    rp_index=i,
                    rp_start=rp_start,
                    rp_end=rp_end,
                    fullcam_date_used=rp_end,
                    calculated_accus=None,
                    cumulative_accus=(cum if cum_has_value else None),
                    carbon_stock_total=None
                ))
                continue

            # Write inputs
            write_date_cell(ws_sum, "B11", rp_start)
            write_date_cell(ws_sum, "B12", rp_end)
            write_date_cell(ws_sum, "B13", fc_date)

            # Recalc with retries until D30 resolves
            accus = None
            carbon_stock = None

            for _attempt in range(RECALC_RETRIES + 1):
                calculate_full_rebuild(wb, times=FULL_REBUILD_TIMES_PER_ATTEMPT)

                accus = safe_float(ws_sum.range("D30").value)  # <-- your confirmed output

                if READ_CARBON_STOCK:
                    sheet, cell = CARBON_STOCK_CELL
                    carbon_stock = safe_float(wb.sheets[sheet].range(cell).value)

                if accus is not None:
                    break

            # Update cumulative
            if accus is not None:
                cum += accus
                cum_has_value = True

            rows.append(ForecastRow(
                rp_index=i,
                rp_start=rp_start,
                rp_end=rp_end,
                fullcam_date_used=fc_date,
                calculated_accus=accus,
                cumulative_accus=(cum if cum_has_value else None),
                carbon_stock_total=carbon_stock
            ))

        # Save workbook (optional but usually helpful)
        wb.save()

    # Write CSV
    with open(CSV_OUT, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow([
            "RP",
            "RP_Start",
            "RP_End",
            "FullCAM_Date_Used",
            "Calculated_ACCUs (Summary!D30)",
            "Cumulative_ACCUs",
            "Carbon_Stock_Total (optional)",
        ])
        for r in rows:
            w.writerow([
                r.rp_index,
                r.rp_start.isoformat(),
                r.rp_end.isoformat(),
                r.fullcam_date_used.isoformat(),
                r.calculated_accus,
                r.cumulative_accus,
                r.carbon_stock_total,
            ])

    return CSV_OUT


if __name__ == "__main__":
    out = run_forecast()
    print(f"✅ Schedule 1 anniversary forecast written to: {out}")

# %%
