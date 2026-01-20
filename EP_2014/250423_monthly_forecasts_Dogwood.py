import os
import xlwings as xw
import csv
from datetime import datetime
from dateutil.relativedelta import relativedelta
    
currentDate = datetime.now().strftime("%Y%m%d")

#adjust file path - add getuser() to make it more dynamic
excel_path = "C:/Users/EmilyHoward/Corporate Carbon Pty Ltd/Corporate Carbon - 04. CARBON DELIVERY/03. Procedures/ACCU forecasting/Monthly Forecasts/EP/Dogwood/Calculator/02. Dogwood Project Calculations - 2023 - Release.xlsx"
csvOut = f"C:/Users/EmilyHoward/Corporate Carbon Pty Ltd/Corporate Carbon - 04. CARBON DELIVERY/03. Procedures/ACCU forecasting/Monthly Forecasts/EP/Dogwood/{currentDate}_AbatementSummary12.csv"

if not os.path.exists(excel_path):
    raise FileNotFoundError(f"Excel file not found: {excel_path}")
print("Input file found.")

#adjust output start date depending on when you want to start forecasting from 
outputStartDate = datetime(2019, 1, 1)
projectStart = datetime(2011, 7, 1)

#leave num_years if you want to generate forecasts on a yearly basis 
#change to num_months if you want to generate forecasts on a monthly basis
#and adjust for how long you want to generate forecasts for
num_years = 20#4 - ((outputStartDate - projectStart).days // 365)
#num_months = 13#300 - ((outputStartDate - projectStart).days // 30)
print(f"Generating data for {num_years} months...")

datePairs = []
currentDate = outputStartDate
for _ in range(num_years):
    start = currentDate.replace(day=1)
    end = (start + relativedelta(years=1)) - relativedelta(days=1)
    datePairs.append((start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")))
    currentDate += relativedelta(years=1)
print(f"Generated {len(datePairs)} date ranges.")

with open(csvOut, 'w', newline='') as myCSV:
    writer = csv.writer(myCSV)
    writer.writerow(['Project Name', 'Start Date', 'End Date', 
                        'RAW_extract', 'C_Abatement'])

    app = xw.App(visible=True, add_book=False)
    wb = app.books.open(excel_path)
    print(f"Workbook loaded: {wb.name}")
    print("Sheets:", [s.name for s in wb.sheets])

    try:
        cover = wb.sheets['Cover']
        report = wb.sheets['Abatement - Report 5']
        jin = wb.sheets['JIN_CEA_01']
        hid = wb.sheets['HID_CEA_01']
        lis = wb.sheets['LIS_CEA_01']
        ste1 = wb.sheets['STE_CEA_01']
        ste2 = wb.sheets['STE_CEA_02']
        ste3 = wb.sheets['STE_CEA_03']

    except KeyError as e:
        raise KeyError(f"Sheet not found: {e}")
    

    project_age = (outputStartDate - projectStart).days // 365

    prev_result = None

    #you'll need to adjust these indices based on which row you want to start from in the relevant sheets 
    #(they are all relating to FullCAM outputs)
    x = 91-24
    x2 = 90-24
    y = 129-24
    z = 146-24

    for start_str, end_str in datePairs:
            try:
                start_dt = datetime.strptime(start_str, "%Y-%m-%d")
                end_dt = datetime.strptime(end_str, "%Y-%m-%d")

                c5 = report.range('C5').value
                jin_e = jin.range(f'E{x}').value
                jin_f = jin.range(f'F{x}').value
                jin_c = (jin_e + jin_f) * c5

                c6 = report.range('C6').value
                hid_e = hid.range(f'E{y}').value
                hid_f = hid.range(f'F{y}').value
                hid_c = (hid_e + hid_f) * c6

                c7 = report.range('C7').value
                lis_e = lis.range(f'E{z}').value
                lis_f = lis.range(f'F{z}').value
                lis_c = (lis_e + lis_f) * c7

                c8 = report.range('C8').value
                ste1_e = ste1.range(f'E{x2}').value
                ste1_f = ste1.range(f'F{x2}').value
                ste1_c = (ste1_e + ste1_f) * c8

                c9 = report.range('C9').value
                ste2_e = ste2.range(f'E{x2}').value
                ste2_f = ste2.range(f'F{x2}').value
                ste2_c = (ste2_e + ste2_f) * c9

                c10 = report.range('C10').value
                ste3_e = ste3.range(f'E{x2}').value      
                ste3_f = ste3.range(f'F{x2}').value
                ste3_c = (ste3_e + ste3_f) * c10

                result = jin_c + hid_c + lis_c + ste1_c + ste2_c + ste3_c
                times = 44/12
                newResult = result*times

                print(f"{start_str} → {end_str}: {newResult}")
                print(f"JIN: {jin_c}, HID: {hid_c}, LIS: {lis_c}, STE1: {ste1_c}, STE2: {ste2_c}, STE3: {ste3_c}")
                print(f"Total result: {result}, Adjusted result: {newResult}")

                if newResult is None:
                    print(f"No result, skipping.")
                    x += 12
                    x2 += 12
                    y += 12
                    z += 12 
                    continue

                projName = cover.range('B15').value or 'Unknown Project'
                newResultInt = int(newResult)
                prevResultInt = int(prev_result) if prev_result is not None else 0
                delta = newResultInt - prevResultInt -3.891384-1.42-25.30
                delta = delta*0.75

                writer.writerow([projName, start_str, end_str, 
                                 newResultInt, delta if prev_result else newResultInt])
                prev_result = newResult
            
            except Exception as loop_err:
                print(f"Error in period {start_str}–{end_str}: {type(loop_err).__name__}: {loop_err}")
            x += 12
            x2 += 12
            y += 12
            z += 12 

    print(f"Processing complete. Output saved to: {csvOut}")
    wb.close()
    app.quit()
    
