# Coalara Forecaster Scripts - Functional Comparison

## Overview
7 scripts in `PF_2020/COALARA/Coalara_Forecaster/` with distinct purposes and methodologies.

---

## Script Comparison Matrix

| Script | Purpose | Core Methodology | Data Source | Output Type | Key Features |
|--------|---------|------------------|-------------|------------|--------------|
| **accu_forecast_runner.py** | ACCU summary table population via xlwings | Reads D30 (carbon stock), writes to Excel table, feeds A90 | Excel workbook (Summary, CEA01, ACCU Summary sheets) | CSV + Excel table (tblACCU) | Writes directly to Excel ACCU Summary table; uses 30-yr horizon |
| **Coalara_Calculatorv2.py** | Yearly abatement summary with CEA deductions | Fixed base year + incremental calculation with deductions | Excel (Summary, CEA01, Calculations) | CSV | v2 logic: subtracts CEA deduction values; fixed 2247.62 base year |
| **Coalara_Forecast_Calculator_MyVersion.py** | FY-based forecast (old method) | Financial Year (1 Jul - 30 Jun) period; dynamic FullCAM date lookup | Excel (Summary, CEA01, Calculations) | CSV | Reads D27-D30; searches CEA01 for 01/07 or 02/07 dates; FY labels |
| **Coalara_Forecast_Calculator.py** | 25-yr baseline forecast (simple incremental) | Annual periods (Jan-Dec); reads D30 carbon stock; computes delta | Excel (Summary, CEA01, Calculations) | CSV | Simplest logic; fixed 25-year horizon; no discounts; delta-only |
| **CoalaraCalc_OG.py** | Original 30-yr forecast template | Similar to Calculator.py but 30-year horizon | Excel (Summary, CEA01, Calculations) | CSV | Project age calculation; starts at row 51 (CEA01); 30-yr limit |
| **MERGE_FC24_FULLCAM_SCHEDULE4.py** | Merge baseline vs project FullCAM outputs | Pandas-based merge on Date; normalizes headers; combines metrics | 2× Excel multi-sheet workbooks | Excel XLSX | Compares Baseline vs Project side-by-side; handles header variants; metric selection |
| **PF_Schedule4_Forecaster.py** | Parametric carbon projection (no Excel) | Mathematical model: CP growth from CBASE→CLT over 180 months; DPP discount | Hardcoded parameters (CBASE, CLT, DPP, RP schedule) | CSV | Pure Python; no Excel dependency; CO2-e conversion; area scaling; first-year deduction |

---

## Detailed Functional Differences

### **1. accu_forecast_runner.py**
- **Data Flow**: Excel workbook → xlwings API → CSV + back to Excel table
- **Unique Aspect**: 
  - Writes results directly to Excel `ACCU Summary` table (tblACCU)
  - Sets `Calculations!A90` formula to point at written carbon stock values
  - Two-way interaction: reads D30, writes to table, feeds back into calculations
- **Reporting Granularity**: Annual (30-year horizon, +12 rows per year in CEA01)
- **Key Values Read**: 
  - Summary!D30 (Project carbon stock total)
  - Summary!A86 (RP label)

### **2. Coalara_Calculatorv2.py** 
- **Data Flow**: Excel workbook → xlwings → CSV
- **Unique Aspect**: 
  - Implements "Rotation Method" (CEA sheet deductions)
  - Subtracts forestry product deductions from column E of CEA sheets before net abatement calc
  - Base year override: Year 0 = 2247.62 (hardcoded)
  - Incremental ACCU calculation: `calc_accu = net_abatement × 0.75`
- **Reporting Granularity**: Annual (25 years, +12 rows per CEA01 row)
- **Outputs**: Stock, deduction, adjusted stock, net abatement, calculated ACCU, issued ACCU

### **3. Coalara_Forecast_Calculator_MyVersion.py** (MyVersion)
- **Data Flow**: Excel workbook → xlwings → CSV
- **Unique Aspect**: 
  - **Financial Year** periods (1 Jul - 30 Jun, not calendar year)
  - Dynamic FullCAM date lookup: searches CEA01 column A for 01/07/YYYY or 02/07/YYYY
  - Reads **D27, D28, D29, D30** separately (Net abatement, previous stock change, net stock change, ACCUs)
  - Output start date: 2023-07-01 (FY 2023/24 start)
  - Computes remaining years from project start to reach 25-year horizon
- **Reporting Granularity**: FY periods (dynamic count based on horizon)
- **Output Headers**: FY label, date range, FullCAM date, D27, D28, D29, D30

### **4. Coalara_Forecast_Calculator.py** (Baseline)
- **Data Flow**: Excel workbook → xlwings → CSV
- **Unique Aspect**: 
  - **Simplest logic** among Excel-based scripts
  - Fixed **25-year** horizon from project start (2022-06-25)
  - Annual periods: Jan 1 - Dec 31 (not FY)
  - Reads only **D30** (no intermediate calculations)
  - Computes delta (`RAW_extract - prev_result`) for abatement
  - Increments CEA01 row by 12 per year (monthly data grain)
- **Key Variables**: Simple delta calculation of D30

### **5. CoalaraCalc_OG.py** (Original)
- **Data Flow**: Excel workbook → xlwings → CSV
- **Unique Aspect**: 
  - Similar to Calculator.py but **30-year** horizon
  - Output start date can be specified (e.g., 2025-07-02)
  - Computes project age from start date
  - Starts CEA01 row at 51 (may align with forecast data vs historical)
  - **Discount schedule commented out** (retained for reference but unused)
- **Use Case**: Longer forecasting horizon; potentially for projects mid-lifecycle

### **6. MERGE_FC24_FULLCAM_SCHEDULE4.py**
- **Data Flow**: 2× Excel files (Baseline + Project) → Pandas DataFrames → Excel XLSX
- **Unique Aspect**: 
  - **No xlwings**; pure Pandas merge
  - Compares **Baseline vs Project** scenarios side-by-side
  - Normalizes column headers (e.g., double-space variants → canonical)
  - Selects 5 canonical metrics (C mass trees, CH4/N2O fire, forest debris, forest products)
  - Outer join on Date; deduplicates by date (keep last); sorts by helper date column
  - Output format: `Date | Baseline - Metric_1 | Baseline - Metric_2 | ... | Project - Metric_1 | ...`
- **Key Assumption**: Both workbooks have matching sheet names and Date column

### **7. PF_Schedule4_Forecaster.py** (Pure Python)
- **Data Flow**: Parameters → Math model → CSV
- **Unique Aspect**: 
  - **No Excel dependency** (pure Python, no xlwings)
  - Carbon pool (CP) growth model: `CP(n) = CBASE + (n/180) × (CLT - CBASE) × DPP`
  - Parameters set in script:
    - `CBASE`: baseline carbon mass (value from B21)
    - `CLT`: long-term carbon mass (value from B28)
    - `PERMANENCE_YEARS`: 25 (DPP = 0.75) or else 1.0
    - `CEA_START`: 2021-06-25
    - `FIRST_RP_START/END`: 2025-10-31 – 2026-06-30
    - `FIRST_YEAR_DEDUCTION`: 12000 ACCUs
  - Handles:
    - CO2-e conversion (multiply by 44/12 if flag set)
    - Area scaling (per-ha → total if AREA_HA set)
    - Month-based computation (up to 180 months to reach CLT)
  - Output: RP index, start/end dates, months elapsed, CP (C & CO2-e), annual/cumulative credits (raw & adjusted)
- **Precision**: Keeps floats (no truncation)

---

## Comparison by Use Case

### **If you need to...**

| Need | Script(s) |
|------|-----------|
| Read current Excel state & write to ACCU table | **accu_forecast_runner.py** |
| Forecast with forest product deductions | **Coalara_Calculatorv2.py** |
| Use Financial Year reporting periods (Jul-Jun) | **Coalara_Forecast_Calculator_MyVersion.py** |
| Simple 25-yr calendar-year baseline | **Coalara_Forecast_Calculator.py** |
| Extended 30-yr horizon forecast | **CoalaraCalc_OG.py** |
| Compare Baseline vs Project FullCAM outputs | **MERGE_FC24_FULLCAM_SCHEDULE4.py** |
| Model carbon growth without Excel | **PF_Schedule4_Forecaster.py** |

---

## Summary of Key Differences

| Dimension | Range |
|-----------|-------|
| **Horizon** | 25 years (most) vs 30 years (CoalaraCalc_OG, accu_forecast_runner) |
| **Period Type** | Calendar year (Jan-Dec) vs Financial Year (Jul-Jun, MyVersion only) |
| **Excel Dependency** | 6 scripts use xlwings; 1 pure Python (PF_Schedule4_Forecaster) |
| **Data Flow** | Read-only Excel (5) vs Two-way Excel (accu_forecast_runner) vs Multi-workbook merge (MERGE_FC24) |
| **Abatement Logic** | Delta of D30 vs Net abatement with deductions vs Pure mathematical model |
| **Output Focus** | Forecast CSV mostly; MERGE outputs side-by-side comparison |
| **Discount/Deduction** | None (baseline) vs CEA deductions (v2) vs First-year only (PF) vs Commented (OG) |

---

## Column Mapping (Excel Reads)

All xlwings-based scripts read from standard sheets:

| Sheet | Key Cells |
|-------|-----------|
| **Summary** | B11 (RP start), B12 (RP end), B13 (FullCAM date), D27-D30 (metrics), B3 (project name) |
| **CEA01** | Column A (FullCAM dates), Column E (deductions, v2 only) |
| **Calculations** | A86 (RP label, used by accu_runner), A90 (set by accu_runner to reference ACCU Summary) |
| **ACCU Summary** | tblACCU table (written by accu_forecast_runner) |

---

## Version/Status Indicators

- **"OG"** = Original (CoalaraCalc_OG.py) — baseline template
- **"v2"** = Version 2 with rotation method (Coalara_Calculatorv2.py)
- **"MyVersion"** = FY-based variant (Coalara_Forecast_Calculator_MyVersion.py)
- **"_Base"** = Simple baseline (Coalara_Forecast_Calculator.py)
- **"ACCU runner"** = Two-way Excel interaction (accu_forecast_runner.py)
- **"PF"** = Portable Forestry standalone model (PF_Schedule4_Forecaster.py)
- **"MERGE"** = Comparative multi-workbook analysis (MERGE_FC24_FULLCAM_SCHEDULE4.py)
