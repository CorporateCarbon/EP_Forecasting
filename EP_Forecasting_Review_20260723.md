# EP_Forecasting — Code & Logic Review

**Date:** 23 July 2026
**Reviewer:** Sanija (with Claude)
**Repo:** CorporateCarbon/EP_Forecasting
**Purpose:** Confirm the forecaster is calculating and outputting correct ACCU forecasts, and assess readiness to build into the platform.

---

## 1. What the tool does

Two Tk-GUI workflows over shared engine modules:

1. **Forecast** — `EP_Forecast_Runner.py` → `Ep_Forecast_Engine.py`. Opens a method calculator workbook in Excel (via xlwings/COM), and for each reporting period (RP) writes the RP number + end date into the `Forecast_script_helper` sheet, forces a recalculation, and reads back `ACCUs Realised`. Outputs an aggregated workbook (Name, Registry ID, RP, RP-Start, RP-End, ACCUs Realised).
2. **Add to inventory** — `add_to_inv_UI.py` → `add_forecast_to_inventory.py`. Cleans the Monday.com Master Inventory + Declared Projects Portfolio exports, deletes existing inventory rows for the ERF after a cutoff, appends the new forecast rows, and writes a "forecast delta" audit workbook.

Three methods are supported: **EP_2014** and **EP_2024** (Environmental Plantings), and **PF_2020** (Plantation Forestry, Sch1/Sch4). The current generalized engine is fed by the EP calculators; PF_2020 still runs through its own standalone forecaster scripts.

## 2. Is it calculating correctly? — Yes, the core logic is sound

The abatement maths lives inside the calculator workbooks (`Forecast_script_helper` sheet), and was verified across all three EP test calculators (Blackwood, Devon Park, Dogwood) — they are identical:

- **Net Abatement** = Carbon mass per CEA × 44/12 (correct CO₂-e conversion).
- **Total Abatement** (`B31`) = Carbon stock at end of RP − carbon stock at end of previous RP (the correct period delta).
- **ACCUs Realised** (`B36`) = `Total − (Permanence×Total + Buffer×Total)` with **Permanence = 0.20** and **Buffer = 0.05**, i.e. a net **×0.75**.

This matches the discount applied in the older per-project scripts (`delta * 0.75`), but is now correctly split into its two regulatory components (20% permanence-period discount for 25-year permanence + 5% risk-of-reversal buffer) and lives in the workbook rather than hard-coded in Python. **Conclusion: the forecast logic is correct.**

### 2a. Independent validation (numbers reproduced from scratch)

To confirm this beyond reading the formulas, the Blackwood forecast was **re-derived independently in Python** — straight from the raw FullCAM carbon time series in each CEA sheet, with no reference to the calculator's own formulas — and compared to the tool's own expected-output workbook (`Blackwood_Forecast_OUTPUT_EXAMPLE.xlsx`). The re-derivation used only: stock = Σ(CEA area × (C mass of trees + forest debris)) × 44/12, then (stock at RP end − stock at previous RP end) × 0.75.

The result was an **exact match to the last decimal on every reporting period tested**:

| RP | Reproduced | Tool output | Difference |
|---|---|---|---|
| 3 | 1053.337 | 1053.337 | 0.000000 |
| 4 | 1119.209 | 1119.209 | 0.000000 |
| 5 | 1123.273 | 1123.273 | 0.000000 |
| 6 | 1085.989 | 1085.989 | 0.000000 |
| 7 | 1034.404 | 1034.404 | 0.000000 |

The validation script is saved in the repo as `validate_ep_forecast.py` (runs with openpyxl, no Excel required). This is strong evidence the tool is producing correct forecasts — and, since the whole calculation reproduces in ~30 lines of pure Python, it also demonstrates the maths can be re-implemented in code for the platform without desktop Excel.

## 3. Bugs found and fixed (on branch `review-fixes-20260723`)

1. **Fixed-RP-count mode crash (functional bug).** In `run_engine`, `final_rp_end` is only computed in full-lifecycle mode. In "forecast N RPs" mode it stayed `None`, and the loop's last-RP override set the final RP's end date to `None` (crash / bad output). Fixed so the override only applies in full-lifecycle mode; fixed-count mode now keeps the normal RP cadence.
2. **`finally`-block `NameError` risk.** The cleanup block referenced `calc_mode_prev`, which was commented out — only masked by a bare `except`. Now defined as `None` and guarded, so cleanup can't swallow a real error.
3. **Dead/broken methods removed.** `final_rp_end_from_project_end` and `rp_end_from_start` were defined without `self`, unused, and non-functional. Removed.

All changes are syntax-checked. No business logic was altered.

## 4. Issues NOT changed — need a decision or method sign-off

- **Discount is additive, not multiplicative.** `Total − (0.20×Total + 0.05×Total)` = ×0.75. Applying the buffer then permanence sequentially would be ×0.95×0.80 = ×0.76. The difference is small but should be confirmed against the current ACCU scheme rules by a method expert. *(Flagged for verification — not corrected.)*
- **Dead GUI control:** the "discount abatement" flag is collected in the runner but never passed to the engine and does nothing (the discount is in the workbook). Recommend removing it from the GUI to avoid confusion.
- **No automated tests.** There is a `test_inputs_&_Outputs/` folder with a Blackwood expected-output example. Recommend a regression test that runs the engine on the sample calculators and diffs against the expected output.
- **PF_2020 not unified.** Plantation Forestry still uses separate Sch1/Sch4 forecaster scripts with a different sheet/cell layout, so the "one engine, three methods" story only holds for the two EP methods today.
- **Legacy per-project scripts** (`EP_2014/…`, `EP_2024/…`) contain hard-coded user paths (e.g. `C:/Users/EmilyHoward/…`) and fixed cell refs (Devon Park reads `I10`, Blackwood `I17`). They are superseded by the engine — recommend archiving them to avoid accidental use.
- **Destructive inventory writes.** `add_forecast_to_inventory` deletes-then-appends rows by ERF + cutoff. Always run against a copy of the Master Inventory. (Timely given the recent Duff folder-deletion incident — worth a backup/undo safeguard.)

## 5. Platform-readiness recommendation

The **calculation logic is correct and now well understood**, which is the important thing. The main blocker for building this into the platform is the **hard dependency on desktop Excel via xlwings/pywin32 COM** — it's Windows-only, needs a visible Excel instance, and can't run headless or server-side at scale.

Because the actual maths is simple — *(carbon-stock delta) × 44/12 × (1 − 0.20 − 0.05)* per RP — the recommendation is to **re-implement the abatement calculation as pure Python/SQL** that reads the FullCAM outputs directly (the CSVs now in the `fullcam_files` bucket), removing Excel from the loop. The Excel calculators would remain the source of truth / analyst tool, but the platform would reproduce their logic in code. That makes forecasting reproducible, testable, and serverable.

**Suggested next steps:** (1) method expert confirms the 0.75 discount convention; (2) port the EP abatement formula to code with a regression test against the sample outputs; (3) decide how PF_2020 folds into the unified engine; (4) archive the legacy per-project scripts.
