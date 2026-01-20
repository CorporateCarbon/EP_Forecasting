#%%##
try:
    import os
    import xlwings as xw
    import csv
    from datetime import datetime
    from dateutil.relativedelta import relativedelta
    
    currentDate = datetime.now().strftime("%Y%m%d")
    excel_path = r"C:\Users\GeorginaDoyle\Corporate Carbon Pty Ltd\Corporate Carbon - 04. CARBON DELIVERY\08. ERF Projects\Coalara Park Australian Sandalwood Plantation Project - AT\FullCAM\250911_Schedule4_FullCAM2024\Chedule4_FC24_Forecast.xlsx"
    csvOut = r"C:\Users\GeorginaDoyle\Corporate Carbon Pty Ltd\Corporate Carbon - 04. CARBON DELIVERY\08. ERF Projects\Coalara Park Australian Sandalwood Plantation Project - AT\FullCAM\250911_Schedule4_FullCAM2024\Chedule4_FC24_Forecast-2026RP.csv"

    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel file not found: {excel_path}")
    print("Input file found.")

    outputStartDate = datetime(2025, 7, 2)
    projectStart = datetime(2021, 6, 25)
    num_years = 30 - ((outputStartDate - projectStart).days // 365)
    #num_months = 300 - ((outputStartDate - projectStart).days // 30)
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
        writer.writerow(['Project Name', 'Start Date', 'End Date', 'FullCam Date',
                         'RAW_extract', 'C_Abatement'])

        app = xw.App(visible=True, add_book=False)
        wb = app.books.open(excel_path)
        print(f"Workbook loaded: {wb.name}")
        print("Sheets:", [s.name for s in wb.sheets])

        try:
            table = wb.sheets['Summary']
            sheet = wb.sheets['CEA01']
        except KeyError as e:
            raise KeyError(f"Sheet not found: {e}")
        

        project_age = (outputStartDate - projectStart).days // 365
        
        prev_result = None
        x = 51 # This needs to be adjusted to the start of the calculations (it refers to the date column in the CEA01 sheet)

        for start_str, end_str in datePairs:
            try:
                start_dt = datetime.strptime(start_str, "%Y-%m-%d")
                end_dt = datetime.strptime(end_str, "%Y-%m-%d")

                # Force Excel to treat these as literal text values 
                table.range('B11').number_format = "@"
                table.range('B12').number_format = "@"
                table.range('B13').number_format = "@"

                table.range('B11').value = f"'{start_dt.strftime('%d/%m/%Y')}"
                table.range('B12').value = f"'{end_dt.strftime('%d/%m/%Y')}"

                # Read and format FullCam date
                fullcam_date = sheet.range(f'A{x}').value
                if not isinstance(fullcam_date, datetime):
                    fullcam_date = datetime.strptime(fullcam_date, "%d/%m/%Y")

                table.range('B13').value = f"'{fullcam_date.strftime('%d/%m/%Y')}"

                print(f"{start_str} → {end_str} | FullCam: {fullcam_date.strftime('%Y-%m-%d')}")

                sheet.activate()
                table.activate()
                wb.sheets['Calculations'].activate()
                wb.app.api.CalculateFullRebuild()

                wb.app.api.CalculateFullRebuild()
                newResult = table.range('D30').value
                print(f"D30: {newResult}")

                if newResult is None:
                    print(f"No value in D30 for {start_str}, skipping.")
                    x += 1
                    continue

                projName = table.range('B3').value or 'Unknown Project'
                newResultInt = int(newResult)
                prevResultInt = int(prev_result) if prev_result is not None else 0
                delta = newResultInt - prevResultInt

                years_since_start = (start_dt - projectStart).days // 365
                # discount_rate = next((r for y, r in discount_schedule if years_since_start >= y), 0.0)
                # discounted_abatement = round(delta * (1 - discount_rate), 2)
                # writer.writerow([projName, start_str, end_str, fullcam_date,
                #                  newResultInt])
                writer.writerow([projName, start_str, end_str, fullcam_date,
                                  newResultInt, delta if prev_result else newResultInt])
                prev_result = newResult

            except Exception as loop_err:
                print(f"Error in period {start_str}–{end_str}: {type(loop_err).__name__}: {loop_err}")
            x += 12

        print(f"Processing complete. Output saved to: {csvOut}")
        wb.close()
        app.quit()

except Exception as e:
    print(f"\n An error occurred: {type(e).__name__}: {e}")
# %%
