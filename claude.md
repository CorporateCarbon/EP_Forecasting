# EP_Forecasting

## Project Overview

ACCU forecasting tool for Emissions Projection (EP) abatement projects. Processes FullCAM end-of-month (EOM) outputs to generate ACCU forecasts aligned to ERF reporting periods, and adds forecast results to the Master Inventory.

## Repository Structure

```
EP_Forecasting/
├── EP_Forecast_Runner.py          # Main GUI runner application
├── Ep_Forecast_Engine.py          # Forecast calculation engine
├── add_forecast_to_inventory.py   # Add forecast results to Master Inventory
├── add_to_inv_UI.py               # GUI for inventory addition
├── venv_setup_wizard.py           # Virtual environment setup for first-time users
├── Test1/                         # Test data/outputs
├── README.md
└── claude.md
```

## Key Scripts

| Script | Purpose |
|--------|---------|
| `EP_Forecast_Runner.py` | Main GUI — configure and run EP forecasting |
| `Ep_Forecast_Engine.py` | Core forecast calculation from FullCAM EOM outputs |
| `add_forecast_to_inventory.py` | Write forecast results to Master Inventory |
| `add_to_inv_UI.py` | GUI wrapper for inventory addition step |
| `venv_setup_wizard.py` | Helper to set up Python venv for first-time users |

## Dependencies

Key libraries:
- `pandas` — data processing
- `python-dateutil` — date manipulation
- `xlwings` or `openpyxl` — Excel I/O
- `tkinter` (built-in) — GUI framework

## How to Run

For first-time setup:
```bash
python venv_setup_wizard.py
```

Run the forecasting GUI:
```bash
python EP_Forecast_Runner.py
```

## Critical Date Handling Note

> **FullCAM outputs are End-of-Month (EOM).** The system automatically indexes entered dates by **-1 day** to capture the previous EOM.

For example:
- Entered date: `1/1/2025`
- System interprets as: `31/12/2024`

This is intentional — ERF reporting periods run from EOM to EOM, but do not overlap. The -1 day adjustment bridges the gap between adjacent reporting periods. **Always account for this when adding or modifying date logic in `Ep_Forecast_Engine.py`.**

## Development Notes

- The -1 day EOM adjustment is a core assumption — document any changes to this logic clearly.
- `EP_Forecast_Runner.py` handles GUI; `Ep_Forecast_Engine.py` handles calculation — keep these separate.
- `add_forecast_to_inventory.py` writes to the Master Inventory — test with a copy of the workbook before running on production data.
- `Test1/` contains test data for validating the forecast engine — use for regression testing.
- `venv_setup_wizard.py` is for non-technical users.
