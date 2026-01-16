#%%##
# Combine multiple FullCAM CSVs into a single Excel workbook (one sheet per CSV)
# - Supports FullCAM 2024 headers: Year, Step In Year, Day of year (day)
# - Supports FullCAM 2016 headers: Year (yr), Month (mo), Day of month (day)
# - Builds Date as DD/MM/YYYY
# - Coerces data columns to numeric in a QC pass
#%%##
import os
import glob
import pandas as pd
from datetime import datetime, timedelta
#%%##
# ========= Settings =========
# If you've already set output_folder elsewhere, you can re-use it.
inputDir = r"C:\Users\GeorginaDoyle\github\EP_FullCAM\COALARA\Sch4_Baseline\FC24"  # <— change if needed
output_folder = inputDir
os.chdir(inputDir)
#%%#
# ========= Helpers =========
def build_date_column(df: pd.DataFrame) -> pd.Series:
    """
    Return a 'Date' Series (DD/MM/YYYY) from either:
      (A) 2016: Year (yr), Month (mo), Day of month (day)
      (B) 2024: Year, Step In Year, Day of year (day)
    """
    # Case A: 2016 headers
    cols_2016 = {'Year (yr)', 'Month (mo)', 'Day of month (day)'}
    if cols_2016.issubset(df.columns):
        tmp = df[['Year (yr)', 'Month (mo)', 'Day of month (day)']].rename(
            columns={'Year (yr)': 'year', 'Month (mo)': 'month', 'Day of month (day)': 'day'}
        )
        dt_series = pd.to_datetime(tmp, errors='coerce')  # this is already a Series
        return dt_series.dt.strftime('%d/%m/%Y').fillna('')

    # Case B: 2024 headers
    cols_2024 = {'Year', 'Step In Year', 'Day of year (day)'}
    if cols_2024.issubset(df.columns):
        year = pd.to_numeric(df['Year'], errors='coerce').astype('Int64')
        doy  = pd.to_numeric(df['Day of year (day)'], errors='coerce').round().astype('Int64')

        def _from_doy(y, d):
            if pd.isna(y) or pd.isna(d) or d < 1:
                return pd.NaT
            try:
                base = datetime(int(y), 1, 1)
                return base + timedelta(days=int(d) - 1)
            except Exception:
                return pd.NaT

        # pd.to_datetime(list_of_py_datetimes) returns a DatetimeIndex -> wrap to Series
        di = pd.to_datetime([_from_doy(y, d) for y, d in zip(year, doy)], errors='coerce')
        dt_series = pd.Series(di, index=df.index)
        return dt_series.dt.strftime('%d/%m/%Y').fillna('')

    # Fallback: no recognizable date columns
    return pd.Series([''] * len(df), index=df.index)


# ========= Gather CSVs =========
filenames = glob.glob('*.csv')
file_dict = {}
#%%##
# ========= Process each CSV =========
for f in filenames:
    try:
        df = pd.read_csv(f)

        # Build Date column
        date_series = build_date_column(df)
        if (date_series != '').any():
            df.insert(0, 'Date', date_series)  # put Date as first column

        # Desired columns (keep if exist; otherwise just keep whatever is present)
        preferred_cols = [
            'Date',
            'C mass of trees  (tC/ha)',
            'C mass of forest litter and deadwood  (tC/ha)',
            'C mass of forest products  (tC/ha)',
            'CH4 emitted due to fire (tCH4/ha)',
            'N2O emitted due to fire (tN2O/ha)',
        ]

        # Rename the litter/deadwood column for output
        old_name = 'C mass of forest litter and deadwood  (tC/ha)'
        new_name = 'C mass of forest debris (tC/ha)'
        if old_name in df.columns:
            df = df.rename(columns={old_name: new_name})
            # ensure preferred_cols refers to the renamed column
            preferred_cols = [new_name if c == old_name else c for c in preferred_cols]
        existing = [c for c in preferred_cols if c in df.columns]
        if 'Date' in existing:
            # keep Date + preferred existing + any other columns that might be needed later
            out_df = df[existing].copy()
        else:
            out_df = df.copy()

        # Truncate sheet name to Excel's 31-char limit
        sheet_name = os.path.splitext(os.path.basename(f))[0][:31]
        file_dict[sheet_name] = out_df

    except Exception as e:
        print(f"⚠️ Skipping {f}: error — {e}")

# ========= Write to Excel (raw pass) =========
today_str = datetime.today().strftime('%Y-%m-%d')
combined_filename = f'combined_output_{today_str}.xlsx'
combined_output_path = os.path.join(output_folder, combined_filename)

if file_dict:
    with pd.ExcelWriter(combined_output_path) as writer:
        for sheet_name, df in file_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"✅ Combined Excel file created at: {combined_output_path}")
else:
    print("⚠️ No valid data to write.")

#%%##
# Section: QC pass – ensure numeric formatting for non-Date columns

# Re-open the data in-memory and coerce non-Date columns to numeric
if file_dict:
    with pd.ExcelWriter(combined_output_path) as writer:
        for sheet_name, df in file_dict.items():
            if 'Date' in df.columns:
                date_col = df[['Date']]
                other_cols = df.drop(columns=['Date']).copy()
                for col in other_cols.columns:
                    other_cols[col] = pd.to_numeric(other_cols[col], errors='coerce')
                cleaned_df = pd.concat([date_col, other_cols], axis=1)
            else:
                cleaned_df = df.apply(pd.to_numeric, errors='coerce')

            cleaned_df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"✅ Combined Excel file created with numeric formatting: {combined_output_path}")
else:
    print("⚠️ QC skipped – no data present.")
