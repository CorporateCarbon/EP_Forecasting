# EP_Forecasting

## Overview

ACCU (Australian Carbon Credit Unit) forecasting tooling for Corporate Carbon Group.
It reads FullCAM carbon-abatement calculator workbooks and projects ACCUs across ERF
reporting periods (RPs), then merges those forecasts into the Master Inventory.

- **EP** = Environmental Plantings method (confirmed: `EP2014`/`EP2024` calculator templates in `EP_2014/`, `EP_2024/`).
- **PF** = Plantation Forestry method (inferred: `PF_2020/` holds `PF_Sch1`/`PF_Sch4` forecasters and calculators). Marked inferred.
- Entity: Corporate Carbon Group; per-project scripts reference "Corporate Carbon Pty Ltd" paths (see `EP_2014/…Devon_Park.py`).

## Tech stack

- Python 3.12 (default targeted by `venv_setup_wizard.py`).
- **xlwings + pywin32** — drives a live Excel (COM) instance to recalculate calculator workbooks (`Ep_Forecast_Engine.py`, per-project scripts). Requires Excel installed on Windows.
- **openpyxl** — file-level Excel read/write for the inventory step (`add_forecast_to_inventory.py`, `helpers/clean_mi_export.py`).
- **pandas**, **python-dateutil**, **tkinter/ttk** (GUIs), **customtkinter** (declared dependency).
- No `requirements.txt`; the dependency list lives in `venv_setup_wizard.py` (`REQUIRED_PACKAGES`).

## Two workflows (each a Tk GUI over an engine module)

1. **Forecast**: `EP_Forecast_Runner.py` (GUI) → `Ep_Forecast_Engine.py` (`run_engine`).
   Opens the calculator workbook, iterates RPs writing inputs into the `Forecast_script_helper`
   sheet (col A labels → col B), forces `app.calculate()`, reads `ACCUs Realised`, and writes an
   aggregated output workbook. Project name/registry come from `Forecast_script_helper!A1/B1`.
2. **Add to inventory**: `add_to_inv_UI.py` (GUI) → `add_forecast_to_inventory.py`
   (`add_forecast_to_inventory`). Cleans the Monday.com Master Inventory + Declared Projects
   Portfolio exports (`helpers/clean_mi_export.py`), deletes existing inventory rows for the ERF
   whose RP-End is after the first forecast RP-Start (cutoff), appends new forecast rows enriched
   from the portfolio, and writes a "forecast delta" audit workbook (Summary / Old / Keep-vs-Replace / New sheets).

## Key files & directories

- `EP_Forecast_Runner.py`, `Ep_Forecast_Engine.py` — forecast GUI + engine.
- `add_to_inv_UI.py`, `add_forecast_to_inventory.py` — inventory GUI + engine.
- `helpers/clean_mi_export.py` — trims header junk above the `Name` row in Monday exports.
- `venv_setup_wizard.py` — self-bootstrapping venv installer for non-technical users.
- `EP_2014/`, `EP_2024/`, `PF_2020/` — older per-project standalone forecast scripts + calculator templates (predecessors to the generalized engine; `PF_2020/ancillary/` has CSV-merge utilities).
- `test_inputs_&_Outputs/` — sample calculator inputs and example outputs for both workflows. `Test1` is an example `.xlsx` (not a directory).

## Run

```bash
python venv_setup_wizard.py        # first-time setup (creates .venv, installs deps, relaunches)
python EP_Forecast_Runner.py       # forecasting GUI
python add_to_inv_UI.py            # add-forecast-to-inventory GUI
```
There is no automated test suite; validate against copies of files in `test_inputs_&_Outputs/`.

## Critical business rules (do not change without care)

- **EOM date indexing**: FullCAM outputs are end-of-month. `EP_Forecast_Runner._validate_and_build_config`
  subtracts **1 day** from the entered start date so a RP nominally `1/1/2025` is calculated from `31/12/2024`
  (previous EOM). ERF RPs don't overlap; the -1 day bridges adjacent periods. See `README.md`.
- **Permanence/risk discount**: per-project EP scripts multiply the period abatement delta by `0.75` (25% held back).
- **Inventory ID**: generated as `f-<Name4>-<yymmdd>-<accus>` for `Forecasted` rows lacking an ID, with name-prefix exceptions (`_generate_inventory_id`).
- **Derived inventory dates**: RP-Start written as forecast start +1 day; Forecasted Submission Date = RP-End +2 days; Date - Total Amount = RP-End +92 days.
- **Inventory writes are destructive** (delete-then-append by ERF + cutoff). Always run against a copy of the Master Inventory workbook.

## Conventions / notes

- Keep GUI and engine modules separate (GUI files only validate input and call the engine).
- Excel is launched **visible** in `Ep_Forecast_Engine.run_engine`; the `finally` block references `calc_mode_prev` which may be unset if calc-mode toggling stays commented out — check before relying on it.
- Windows-only due to xlwings/pywin32/Excel COM dependency.
