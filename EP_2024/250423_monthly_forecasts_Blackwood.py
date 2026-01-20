import os
import xlwings as xw
import csv
from datetime import datetime
from dateutil.relativedelta import relativedelta
    
currentDate = datetime.now().strftime("%Y%m%d")
excel_path = "C:/Users/EmilyHoward/Corporate Carbon Pty Ltd/Corporate Carbon - 04. CARBON DELIVERY/03. Procedures/ACCU forecasting/Monthly Forecasts/EP/Blackwood/Calculator/250428_Blackwood_ACCU_Calcul_NO Treat SIMUL.xlsx"
csvOut = f"C:/Users/EmilyHoward/Corporate Carbon Pty Ltd/Corporate Carbon - 04. CARBON DELIVERY/03. Procedures/ACCU forecasting/Monthly Forecasts/EP/Blackwood/{currentDate}_AbatementSummary_Bwood1.csv"

if not os.path.exists(excel_path):
    raise FileNotFoundError(f"Excel file not found: {excel_path}")
print("Input file found.")

outputStartDate = datetime(2023, 10, 1)
projectStart = datetime(2022, 5, 4)
num_years = 25 - ((outputStartDate - projectStart).days // 365)
#num_months = 24 #300 - ((outputStartDate - projectStart).days // 30)
print(f"Generating data for {num_years} years...")

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
        result = wb.sheets['Abatement']
    except KeyError as e:
        raise KeyError(f"Sheet not found: {e}")
    

    project_age = (outputStartDate - projectStart).days // 365

    prev_result = None

    for i, datePair in enumerate(datePairs[:num_years]): 
        cover.range('B20').value = datePair[0]
        cover.range('B21').value = datePair[1]
        print(f"Setting dates: Start month: {datePair[0]}, End month: {datePair[1]}")
        wb.app.calculate()

        newResult = result.range('I17').value #NetAbatement C16
        print(f"New result for {datePair[0]} to {datePair[1]}: {newResult}")

        projName = 'Blackwood Biodiversity and Carbon Project'

        if newResult is not None:
            try:
                newResultInt = int(newResult)
                prevResultInt = int(prev_result) if prev_result is not None else 0
            except ValueError:
                newResultInt = 0
                prevResultInt = 0

        period_start = datetime.strptime(datePair[0], "%Y-%m-%d")
        years_since_start = (period_start - projectStart).days // 365
        delta = newResultInt - prevResultInt
        delta = delta * 0.75

        if prev_result is None:
            writer.writerow([projName, datePair[0], datePair[1], newResultInt, newResultInt])
        else:
            writer.writerow([projName, datePair[0], datePair[1], newResultInt, delta])

        prev_result = newResult

    wb.close()
    app.quit()

