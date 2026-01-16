# -*- coding: utf-8 -*-
"""
ACCU Forecast Runner (xlwings)
--------------------------------
- Iterates yearly RP dates
- Sets Summary!B11 (RP start), Summary!B12 (RP end), Summary!B13 (FullCAM date = RP date)
- Forces full recalc
- Reads Calculations!A28 (Project carbon stock total) and A86 (RP)
- Populates 'ACCU summary' table (tblACCU) for the matching Date row:
    Date | Carbon stock | Net abatement amount | Calculated ACCU | Issued ACCU | RP
- Writes a separate CSV of forecasts

Open the file and update:
- excel_path → your workbook path
- (Optional) output_start_date → first RP FullCAM date to simulate
- (Optional) project_start → used to compute a 30-year horizon if you don’t pass horizon_years

It assumes:

- ACCU summary has a Table named tblACCU with columns:
- Date, Carbon stock, Net abatement amount, Calculated ACCU, Issued ACCU, RP
- Your totals and RP are already calculated at Calculations!A28 and Calculations!A86 (as you described).

If your Calculated ACCU should apply a different rule than “equals Net abatement”, tweak the calc_accu line in write_tblaccu_row().


"""

#%%##
import os
import csv
from datetime import datetime
from dateutil.relativedelta import relativedelta
import xlwings as xw
#%%##
excel_path = r"C:\Users\GeorginaDoyle\Downloads\NEW_CP_PlForestry_FC2016_Calculator.xlsx"  # <-- update
out_dir = os.path.dirname(excel_path) or "."
csv_out = os.path.join(out_dir, f"ACCU_Forecast_{datetime.now():%Y%m%d-%H%M%S}.csv")


project_start = datetime(2022, 6, 25)      # used for 30-yr horizon if not provided
start_cea_row = 51                          # CEA01 row for first FullCAM (RP end) date
cea_step_rows = 12                          # +12 rows per year (monthly entries)

accu_summary_first_row = 4   # first output row in ACCU Summary

def ensure_datetime(value):
    if isinstance(value, datetime):
        return value
    if value is None:
        raise ValueError("Date cell is empty.")
    if isinstance(value, str):
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(value, fmt)
            except ValueError:
                continue
    raise ValueError(f"Could not parse date: {value!r}")

def set_summary_dates_as_text(ws_summary, start_dt, end_dt):
    # Summary!B11 (start), B12 (end), B13 (same as end) as TEXT 'dd/mm/yyyy
    for addr in ("B11", "B12", "B13"):
        ws_summary.range(addr).number_format = "@"
    ws_summary.range("B11").value = "'" + start_dt.strftime("%d/%m/%Y")
    ws_summary.range("B12").value = "'" + end_dt.strftime("%d/%m/%Y")
    ws_summary.range("B13").value = "'" + end_dt.strftime("%d/%m/%Y")

def get_d30_cumulative(ws_summary):
    val = ws_summary.range("D30").value
    try:
        return float(val) if val not in (None, "") else 0.0
    except Exception:
        return 0.0

def write_accu_summary_row(ws_accu, row_idx, rp_end_date, d30_cum, net_abatement, calc_accu, rp_label):
    """
    ACCU Summary column mapping (so A90->B<row> is numeric!):
      A: Date (RP end / FullCAM)
      B: Carbon Stock (D30 cumulative)   <-- numeric target for Calculations!A90
      C: Net Abatement (ΔD30)
      D: Calculated ACCU
      E: Issued ACCU (placeholder)
      F: RP label string (e.g., '2025-2026')
    """
    ws_accu.range(f"A{row_idx}").number_format = "dd/mm/yyyy"
    ws_accu.range(f"A{row_idx}").value = rp_end_date
    ws_accu.range(f"B{row_idx}").value = d30_cum
    ws_accu.range(f"C{row_idx}").value = net_abatement
    ws_accu.range(f"D{row_idx}").value = calc_accu
    ws_accu.range(f"E{row_idx}").value = None
    ws_accu.range(f"F{row_idx}").value = rp_label

def set_calculations_a90_to_stock(ws_calc, target_row):
    """Point Calculations!A90 at numeric carbon stock in 'ACCU Summary'!B<row>."""
    ws_calc.range("A90").formula = f"='ACCU Summary'!B{target_row}"

def run_forecast(excel_path, csv_out, project_start, horizon_years=None, visible=True):
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    with xw.App(visible=visible, add_book=False) as app:
        wb = app.books.open(excel_path)
        ws_sum  = wb.sheets["Summary"]
        ws_calc = wb.sheets["Calculations"]
        ws_cea  = wb.sheets["CEA01"]
        ws_accu = wb.sheets["ACCU Summary"]

        # Determine first RP from CEA01
        first_end = ensure_datetime(ws_cea.range(f"A{start_cea_row}").value)
        first_start = datetime(first_end.year - 1, 7, 1)

        if horizon_years is None:
            horizon_years = 30 - ((first_start - project_start).days // 365)

        prev_cum = None
        cea_row = start_cea_row
        accu_row = accu_summary_first_row

        with open(csv_out, "w", newline="") as f:
            w = csv.writer(f)
            w.writerow([
                "Project Name",           # A
                "RP Start",               # B
                "RP End (FullCAM)",       # C
                "D30 Cumulative",         # D
                "Net Abatement (ΔD30)",   # E
                "Calculated ACCU"         # F
            ])

            for _ in range(horizon_years):
                # RP end (FullCAM) from CEA01
                end_dt = ensure_datetime(ws_cea.range(f"A{cea_row}").value)
                start_dt = datetime(end_dt.year - 1, 7, 1)

                # Write Summary dates as TEXT and force full rebuild twice (mimic old behaviour)
                set_summary_dates_as_text(ws_sum, start_dt, end_dt)
                ws_cea.activate(); ws_sum.activate(); ws_calc.activate()
                wb.app.api.CalculateFullRebuild()
                wb.app.api.CalculateFullRebuild()

                # Read cumulative and compute deltas
                d30_cum = get_d30_cumulative(ws_sum)
                net_abatement = d30_cum - prev_cum if prev_cum is not None else d30_cum
                calc_accu = net_abatement

                # RP label (string) for human readability, stored in F
                rp_label = f"{start_dt.year}-{end_dt.year}"

                # Write the row (B gets numeric carbon stock)
                write_accu_summary_row(
                    ws_accu, accu_row,
                    rp_end_date=end_dt,
                    d30_cum=d30_cum,
                    net_abatement=net_abatement,
                    calc_accu=calc_accu,
                    rp_label=rp_label
                )

                # Now set A90 to the numeric carbon stock cell we just wrote
                set_calculations_a90_to_stock(ws_calc, accu_row)

                # CSV row
                proj_name = ws_sum.range("B3").value or "Unknown Project"
                w.writerow([proj_name, start_dt.date(), end_dt.date(), d30_cum, net_abatement, calc_accu])

                # advance
                prev_cum = d30_cum
                cea_row += cea_step_rows
                accu_row += 1

        wb.save()
    return csv_out

if __name__ == "__main__":
    print("Running forecast...")
    out_path = run_forecast(
        excel_path=excel_path,
        csv_out=csv_out,
        project_start=project_start,
        horizon_years=None,
        visible=True,
    )
    print(f"Done. CSV written to: {out_path}")
# %%
