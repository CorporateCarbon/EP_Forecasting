# v1 - forecast calculator for Coalara Permanent Stand Method
#%%##
import os
import csv
from datetime import datetime
from dateutil.relativedelta import relativedelta
import xlwings as xw
#%%##
   # excel_path = r"C:\Users\GeorginaDoyle\Downloads\FullCAM_Test\NewMethod(rp03)\Rp03_Calculator_test.xlsx"
   # csvOut = rf"C:\Users\GeorginaDoyle\Downloads\FullCAM_Test\NewMethod(rp03)\{currentDate}_Test_AbatementSummary.csv"

# === Setup ===
#excel_path = r"C:\Users\GeorginaDoyle\Downloads\Coalara_RP3_ACCU_Calc_v3.xlsx"
#current_date = datetime.now().strftime("%Y%m%d")
#csv_out = rf"C:\Users\GeorginaDoyle\Downloads\{current_date}_New_YearlySummary.csv"

currentDate = datetime.now().strftime("%Y%m%d")
excel_path = r"C:\Users\GeorginaDoyle\Downloads\FullCAM_Test\NewMethod(rp03)\Rp03_Calculator_test.xlsx"
csvOut = rf"C:\Users\GeorginaDoyle\Downloads\FullCAM_Test\NewMethod(rp03)\{currentDate}_Test2_AbatementSummary.csv"

#%%##
if not os.path.exists(excel_path):
    raise FileNotFoundError(f"Excel file not found: {excel_path}")

print("Input file found.")

# === Generate 25 years of reporting periods ===
project_start = datetime(2021, 6, 25)
num_years = 25
date_pairs = []

current_date_obj = project_start
for _ in range(num_years):
    start = current_date_obj.replace(day=1)
    end = (start + relativedelta(years=1)) - relativedelta(days=1)
    date_pairs.append((start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")))
    current_date_obj += relativedelta(years=1)

print(f"Generated {len(date_pairs)} date ranges.")

# === Open Excel workbook ===
app = xw.App(visible=True, add_book=False)
wb = app.books.open(excel_path)

try:
    summary_sheet = wb.sheets['Summary']
    cea_sheet = wb.sheets['CEA01']
    calc_sheet = wb.sheets['Calculations']
except Exception as e:
    raise RuntimeError("Could not find required sheets in workbook.") from e

# === Loop through years ===
accumulated_accu = 0
row_idx = 2  # FullCAM dates start in CEA01!A3
summary_rows = []

for year_idx, (start_str, end_str) in enumerate(date_pairs):
    try:
        start_dt = datetime.strptime(start_str, "%Y-%m-%d")
        end_dt = datetime.strptime(end_str, "%Y-%m-%d")

        # Force Excel cells as text
        for cell in ['B11','B12','B13']:
            summary_sheet.range(cell).number_format = "@"

        summary_sheet.range('B11').value = f"'{start_dt.strftime('%d/%m/%Y')}"
        summary_sheet.range('B12').value = f"'{end_dt.strftime('%d/%m/%Y')}"

        fullcam_date = cea_sheet.range(f"A{row_idx}").value
        if not isinstance(fullcam_date, datetime):
            fullcam_date = datetime.strptime(str(fullcam_date), "%d/%m/%Y")

        summary_sheet.range('B13').value = f"'{fullcam_date.strftime('%d/%m/%Y')}"

        # Trigger recalculation
        calc_sheet.activate()
        wb.app.api.CalculateFullRebuild()

        # Extract Carbon Stock (Summary!D30)
        current_stock = summary_sheet.range('D30').value
        if current_stock is None:
            print(f"⚠️ No value in D30 for {start_str}")
            row_idx += 12
            continue

        current_stock = float(current_stock)

        # === New Logic ===
        if year_idx == 0:
            # Base year override
            net_abatement = 2247.62
        else:
            net_abatement = current_stock - accumulated_accu

        calc_accu = round(net_abatement * 0.75, 2)
        issued_accu = round(max(0, calc_accu - accumulated_accu), 2)

        # Update total issued
        accumulated_accu += issued_accu

        summary_rows.append([
            start_dt.strftime("%Y"),
            start_str, end_str, fullcam_date.strftime("%d/%m/%Y"),
            round(current_stock, 2), round(net_abatement, 2),
            calc_accu, issued_accu
        ])

        print(f"{start_dt.year}: Stock={current_stock:.2f}, "
              f"Net={net_abatement:.2f}, CalcACCU={calc_accu}, Issued={issued_accu}")

    except Exception as loop_err:
        print(f"❌ Error in year {start_str}: {loop_err}")

    row_idx += 12  # move to next July in CEA01

# === Write results to CSV ===
with open(csvOut, 'vw', newline='') as f:
    writer = csv.writer(f)
    writer.writerow([
        "Year", "Start Date", "End Date", "FullCam Date",
        "Carbon Stock (D30)", "Net Abatement", "Calculated ACCU", "Issued ACCU"
    ])
    writer.writerows(summary_rows)

# Cleanup
wb.close()
app.quit()
print(f"✅ Processing complete. Yearly Summary saved to {csv_out}")

# %%
#V2 - Forecast Calulator for Coalara - Rotation Method (discount forest products) 
import os
import csv
from datetime import datetime
from dateutil.relativedelta import relativedelta
import xlwings as xw
#%%##
# === Setup ===
excel_path = r"C:\Users\GeorginaDoyle\Downloads\241114_Coalara_RP2_ACCU_Calc.xlsx"
current_date = datetime.now().strftime("%Y%m%d")
csv_out = rf"C:\Users\GeorginaDoyle\Downloads\{current_date}_OldMeth_YearlySummary.csv"

if not os.path.exists(excel_path):
    raise FileNotFoundError(f"Excel file not found: {excel_path}")

print("Input file found.")

# === Generate 25 years of reporting periods ===
project_start = datetime(2021, 6, 25)
num_years = 25
date_pairs = []

current_date_obj = project_start
for _ in range(num_years):
    start = current_date_obj.replace(day=1)
    end = (start + relativedelta(years=1)) - relativedelta(days=1)
    date_pairs.append((start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")))
    current_date_obj += relativedelta(years=1)

print(f"Generated {len(date_pairs)} date ranges.")

# === Open Excel workbook ===
app = xw.App(visible=True, add_book=False)
wb = app.books.open(excel_path)

try:
    summary_sheet = wb.sheets['Summary']
    cea_sheets = [s for s in wb.sheets if s.name.lower().startswith("cea")]
    calc_sheet = wb.sheets['Calculations']
except Exception as e:
    raise RuntimeError("Could not find required sheets in workbook.") from e

# === Loop through years ===
accumulated_accu = 0
row_idx = 3  # FullCAM dates start in row 3 of each CEA sheet
summary_rows = []

for year_idx, (start_str, end_str) in enumerate(date_pairs):
    try:
        start_dt = datetime.strptime(start_str, "%Y-%m-%d")
        end_dt = datetime.strptime(end_str, "%Y-%m-%d")

        # Force Excel cells as text
        for cell in ['B11','B12','B13']:
            summary_sheet.range(cell).number_format = "@"

        summary_sheet.range('B11').value = f"'{start_dt.strftime('%d/%m/%Y')}"
        summary_sheet.range('B12').value = f"'{end_dt.strftime('%d/%m/%Y')}"

        # Use FullCAM date from CEA01 to drive Summary
        fullcam_date = cea_sheets[0].range(f"A{row_idx}").value
        if not isinstance(fullcam_date, datetime):
            fullcam_date = datetime.strptime(str(fullcam_date), "%d/%m/%Y")

        summary_sheet.range('B13').value = f"'{fullcam_date.strftime('%d/%m/%Y')}"

        # Trigger recalculation
        calc_sheet.activate()
        wb.app.api.CalculateFullRebuild()

        # Extract Carbon Stock (Summary!D30)
        current_stock = summary_sheet.range('D30').value
        if current_stock is None:
            print(f"⚠️ No value in D30 for {start_str}")
            row_idx += 12
            continue
        current_stock = float(current_stock)

        # === NEW: subtract any values in col E of CEA sheets (same row) ===
        deduction_total = 0.0
        for cea in cea_sheets:
            val = cea.range(f"E{row_idx}").value
            try:
                deduction_total += float(val) if val is not None else 0.0
            except Exception:
                continue
        adjusted_stock = current_stock - deduction_total

        # === Abatement logic ===
        if year_idx == 0:
            net_abatement = 2247.62  # fixed base
        else:
            net_abatement = adjusted_stock - accumulated_accu

        calc_accu = round(net_abatement * 0.75, 2)
        issued_accu = round(max(0, calc_accu - accumulated_accu), 2)

        # Update running total
        accumulated_accu += issued_accu

        summary_rows.append([
            start_dt.strftime("%Y"),
            start_str, end_str, fullcam_date.strftime("%d/%m/%Y"),
            round(current_stock, 2), round(deduction_total, 2),
            round(adjusted_stock, 2), round(net_abatement, 2),
            calc_accu, issued_accu
        ])

        print(f"{start_dt.year}: Stock={current_stock:.2f}, Deduct={deduction_total:.2f}, "
              f"AdjStock={adjusted_stock:.2f}, Net={net_abatement:.2f}, "
              f"CalcACCU={calc_accu}, Issued={issued_accu}")

    except Exception as loop_err:
        print(f"❌ Error in year {start_str}: {loop_err}")

    row_idx += 12  # move to next July in CEA sheets

# === Write results to CSV ===
with open(csv_out, 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow([
        "Year", "Start Date", "End Date", "FullCam Date",
        "Carbon Stock (D30)", "CEA Deduction (ΣE)", "Adjusted Stock",
        "Net Abatement", "Calculated ACCU", "Issued ACCU"
    ])
    writer.writerows(summary_rows)

# Cleanup
wb.close()
app.quit()
print(f"✅ Processing complete. Yearly Summary saved to {csv_out}")

# %%
