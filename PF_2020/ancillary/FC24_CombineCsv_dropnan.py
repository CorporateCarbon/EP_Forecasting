#%%##
# Combine multiple CSV files from FullCAM 2024 into a single Excel workbook.
# Drops rows that contain NaN values in Date or metric columns.

import os
import glob
import pandas as pd
from datetime import datetime

# === Use existing folder path set earlier ===
inputDir = r"C:\Users\GeorginaDoyle\Corporate Carbon Pty Ltd\Corporate Carbon - 04. CARBON DELIVERY\08. ERF Projects\Coalara Park Australian Sandalwood Plantation Project - AT\FullCAM\September 2025 Reforecast\250910_Schedule1_FullCAM2024\Output"  # Change this to your directory
output_folder = r"C:\Users\GeorginaDoyle\Downloads"  # You can change this if needed
os.chdir(inputDir)

# === Gather all CSV files ===
extension = "csv"
filenames = glob.glob(f"*.{extension}")
file_dict = {}

# === Process each CSV file ===
for f in filenames:
    try:
        df = pd.read_csv(f)

        if {"Year (yr)", "Month (mo)", "Day of month (day)"}.issubset(df.columns):
            df = df.copy()
            df["Date"] = pd.to_datetime(
                df[["Year (yr)", "Month (mo)", "Day of month (day)"]].rename(
                    columns={
                        "Year (yr)": "year",
                        "Month (mo)": "month",
                        "Day of month (day)": "day",
                    }
                ),
                errors="coerce",
            )

            # Remap alternate column name to canonical
            df = df.rename(
                columns={
                    "C mass of forest litter and deadwood  (tC/ha)": "C mass of forest debris  (tC/ha)"
                }
            )

            desired_columns = [
                "Date",
                "C mass of trees  (tC/ha)",
                "CH4 emitted due to fire (tCH4/ha)",
                "C mass of forest debris  (tC/ha)",
                "C mass of forest products  (tC/ha)",
                "N2O emitted due to fire (tN2O/ha)",
            ]

            missing_cols = [col for col in desired_columns[1:] if col not in df.columns]
            if missing_cols:
                print(f"Skipping {f}: missing columns: {missing_cols}")
                continue

            df = df[desired_columns]

            # Coerce metrics to numeric and drop rows with any NaN values
            for col in desired_columns[1:]:
                df[col] = pd.to_numeric(df[col], errors="coerce")

            before = len(df)
            df = df.dropna(subset=desired_columns, how="any")
            dropped = before - len(df)
            if dropped:
                print(f"Note: dropped {dropped} rows with NaN values in {f}")

            df["Date"] = df["Date"].dt.strftime("%d/%m/%Y").astype(str)

            sheet_name = f.replace(".csv", "")[:31]
            file_dict[sheet_name] = df
        else:
            print(f"Skipping {f}: missing Year/Month/Day columns")
    except Exception as e:
        print(f"Skipping {f}: error {e}")

# === Write to Excel ===
today_str = datetime.today().strftime("%Y-%m-%d")
combined_filename = f"combined_output_{today_str}.xlsx"
combined_output_path = os.path.join(output_folder, combined_filename)

if file_dict:
    with pd.ExcelWriter(combined_output_path) as writer:
        for sheet_name, df in file_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"Combined Excel file created at: {combined_output_path}")
else:
    print("No valid data to write.")

# %%
