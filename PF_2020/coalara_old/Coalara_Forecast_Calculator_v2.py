'''
Script: Coalara Park FullCAM FY-based ACCU forecast (no discounts)
Precision: values are kept as floats (no truncation).

Uses Financial Year periods (1 Jul → 30 Jun) like your Script 1

- Dynamically finds the FullCAM date for each FY as 01/07/YYYY or 02/07/YYYY on CEA01!A:A

After each recalculation, reads:
- Summary!D27 → Net abatement (reporting period)
-Summary!D28 → Previous reporting period carbon stock change
-Summary!D29 → Net carbon stock change (D27 - D28)
- Summary!D30 → Calculated ACCUs
'''

# %%FY-based loop with dynamic FullCAM date (01/07 or 02/07), no discounts
import os
import csv
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
import xlwings as xw

# ----------------------------
# User parameters
# ----------------------------
excel_path = r"C:\Users\GeorginaDoyle\Downloads\CP_forecast\Coalara_RP2_ACCU_Calc_NewFormulas.xlsx"
currentDate = datetime.now().strftime("%Y%m%d")
csvOut = rf"C:\Users\GeorginaDoyle\Downloads\CP_forecast\{currentDate}_OldMethod_FY_Summary.csv"

# Anchor dates (match Script 1 behaviour)
project_start = datetime(2021, 6, 25)
output_start = datetime(2023, 7, 1)  # FY 2023/24 start (1 July 2023), as per Script 1
# Compute how many FY periods remain to reach 25 years from project start
num_years = 25 - ((output_start - project_start).days // 365)

# ----------------------------
# Helpers
# ----------------------------
def ensure_text_dates(summary_sheet, start_dt, end_dt):
    """
    Force Summary!B11, B12, B13 as text, then write B11/B12 dates (dd/mm/yyyy) with leading apostrophe.
    B13 (FullCAM) is written separately after we find it.
    """
    for cell in ('B11', 'B12', 'B13'):
        summary_sheet.range(cell).number_format = "@"
    summary_sheet.range('B11').value = f"'{start_dt.strftime('%d/%m/%Y')}"
    summary_sheet.range('B12').value = f"'{end_dt.strftime('%d/%m/%Y')}"

def parse_maybe_excel_date(v):
    """Handle Excel date (datetime) or string 'dd/mm/yyyy' safely; return datetime.date or None."""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None

def find_fullcam_date_for_year(cea_sheet, year):
    """
    Search CEA01!A:A for a date equal to 01/07/<year> or 02/07/<year>.
    Returns a date object if found, else None.
    """
    target1 = date(year, 7, 1)
    target2 = date(year, 7, 2)

    # Read the entire used column A once
    colA = cea_sheet.range('A1').expand('down').value
    if not isinstance(colA, list):
        colA = [colA]

    # Scan for either target date
    for v in colA:
        d = parse_maybe_excel_date(v)
        if d in (target1, target2):
            return d
    return None

# ----------------------------
# Validate input workbook
# ----------------------------
if not os.path.exists(excel_path):
    raise FileNotFoundError(f"Excel file not found: {excel_path}")
print("Input file found.")

# ----------------------------
# Build FY date ranges: 1 Jul YYYY → 30 Jun (YYYY+1)
# ----------------------------
date_pairs = []
cur = output_start
for _ in range(max(0, num_years)):
    start = cur
    end = (start + relativedelta(years=1)) - relativedelta(days=1)  # 30 June next year
    date_pairs.append((start, end))
    cur = start + relativedelta(years=1)

print(f"Generated {len(date_pairs)} FY ranges starting {output_start.strftime('%d/%m/%Y')}.")

# ----------------------------
# Open workbook
# ----------------------------
app = xw.App(visible=True, add_book=False)
wb = app.books.open(excel_path)

try:
    summary = wb.sheets['Summary']
    cea = wb.sheets['CEA01']
    calcs = wb.sheets['Calculations']
except Exception as e:
    # Ensure we close Excel if sheets not found
    wb.close()
    app.quit()
    raise RuntimeError("Could not find required sheets: Summary, CEA01, Calculations") from e

rows_out = []

try:
    for start_dt, end_dt in date_pairs:
        # FY year label = starting year (e.g., 2023 for FY23/24)
        fy_year = start_dt.year

        # 1) Write reporting period (as text)
        ensure_text_dates(summary, start_dt, end_dt)

        # 2) Find the FullCAM date for this FY: 01/07/<start_year> or 02/07/<start_year>
        fc_date = find_fullcam_date_for_year(cea, fy_year)
        if fc_date is None:
            print(f"⚠️ No FullCAM date found for FY starting {start_dt.strftime('%d/%m/%Y')} "
                  f"(expected 01/07/{fy_year} or 02/07/{fy_year}). Skipping.")
            continue

        # 3) Write Summary!B13 (as text)
        summary.range('B13').number_format = "@"
        summary.range('B13').value = f"'{fc_date.strftime('%d/%m/%Y')}"

        # 4) Recalculate workbook
        calcs.activate()
        wb.app.api.CalculateFullRebuild()

        # 5) Read outputs from Summary
        #    D27 = Net abatement (reporting period)
        #    D28 = Previous reporting period carbon stock change
        #    D29 = Net carbon stock change (D27 - D28)
        #    D30 = Calculated ACCUs
        def read_float(cell_addr):
            v = summary.range(cell_addr).value
            try:
                return float(v)
            except (TypeError, ValueError):
                return None

        net_abatement = read_float('D27')
        prev_stock_change = read_float('D28')
        net_stock_change = read_float('D29')
        calc_accus = read_float('D30')

        # Keep row even if some values are None; caller can inspect gaps
        rows_out.append([
            f"FY{start_dt.year}/{(start_dt.year + 1) % 100:02d}",
            start_dt.strftime("%Y-%m-%d"),
            end_dt.strftime("%Y-%m-%d"),
            fc_date.strftime("%d/%m/%Y"),
            net_abatement,
            prev_stock_change,
            net_stock_change,
            calc_accus
        ])

        print(f"FY {fy_year}/{fy_year+1}: "
              f"NetAbate={net_abatement} | PrevStockΔ={prev_stock_change} | "
              f"NetStockΔ={net_stock_change} | ACCUs={calc_accus}")

    # ----------------------------
    # Write CSV
    # ----------------------------
    os.makedirs(os.path.dirname(csvOut), exist_ok=True)
    with open(csvOut, 'w', newline='') as f:
        w = csv.writer(f)
        w.writerow([
            "Financial Year",
            "Start Date",
            "End Date",
            "FullCAM Date",
            "Net Abatement (D27)",
            "Prev Stock Change (D28)",
            "Net Stock Change (D29)",
            "Calculated ACCUs (D30)",
        ])
        w.writerows(rows_out)

    print(f"✅ Processing complete. FY Summary saved to {csvOut}")

finally:
    # Always close workbook/app
    wb.close()
    app.quit()

# %%
