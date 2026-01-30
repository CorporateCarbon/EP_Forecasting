# add_forecast_to_inventory.py  (updated to use a separate Declared Projects Portfolio workbook)

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta, date
from pathlib import Path
from typing import Any, Dict, List, Optional

from openpyxl import load_workbook, Workbook


@dataclass
class AppConfig:
    forecast_workbook: str
    master_inventory_workbook: str
    declared_project_portfolio_workbook: str
    save_master_inventory_output: str
    save_forecast_delta_output: str


REQUIRED_INVENTORY_HEADERS: List[str] = [
    "Name", "Subitems", "Registry ID", "Inventory ID", "Status", "Issuance Status",
    "Days Since (Status)", "Date - Total Amount", "Total Amount (ACCUs)",
    "Realised Amount (ACCUs to CCG)", "Date - Realised Amount", "Forecasted Submission Date",
    "Unit Type", "Application ID", "Reporting Period - Start", "Reporting Period - End",
    "RP", "Delay Flag", "Data Source", "Data Update Date", "Declared Projects Portfolio",
    "Project Number", "Entity", "Proponents", "Methodology", "Business Unit", "Project Stage",
    "Operational Model", "Fee Model", "Number", "Unit", "Item ID", "Key",
]

PORTFOLIO_FIELDS: List[str] = [
    "Name",
    "Subitems",
    "Registry ID",
    "Project ID",          # written to inventory as Project Number
    "Methodology",
    "Project Stage",
    "Proponents",
    "Business Unit",
    "Operational Model",
    "Fee Model",
    "Entity",
    "Number",
    "Unit",
]

FORECAST_TO_INVENTORY_MAP: Dict[str, str] = {
    "RP Number": "RP",
    "RP Start (EOM)": "Reporting Period - Start",
    "RP End (EOM)": "Reporting Period - End",
    "ACCUs Realised": "Total Amount (ACCUs)",
}


def _norm(x: Any) -> str:
    return "" if x is None else str(x).strip()


def _lower_norm(x: Any) -> str:
    return _norm(x).lower()


def _build_header_map(ws, header_row: int = 1) -> Dict[str, int]:
    mapping: Dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=col).value
        key = _lower_norm(v)
        if key and key not in mapping:
            mapping[key] = col
    return mapping


def _ensure_headers(ws, required_headers: List[str]) -> Dict[str, int]:
    header_map = _build_header_map(ws, 1)

    if not header_map:
        for i, h in enumerate(required_headers, start=1):
            ws.cell(row=1, column=i).value = h
        return _build_header_map(ws, 1)

    next_col = ws.max_column + 1
    for h in required_headers:
        key = _lower_norm(h)
        if key not in header_map:
            ws.cell(row=1, column=next_col).value = h
            header_map[key] = next_col
            next_col += 1
    return header_map


def _find_sheet_by_keywords(wb, keywords: List[str]):
    for name in wb.sheetnames:
        lname = name.lower()
        if all(k.lower() in lname for k in keywords):
            return wb[name]
    return None


def _find_row_by_value(ws, header_map: Dict[str, int], header_name: str, value: Any) -> Optional[int]:
    key = _lower_norm(header_name)
    if key not in header_map:
        raise ValueError(f"Could not find column '{header_name}' in sheet '{ws.title}'.")

    target = _norm(value)
    if not target:
        return None

    col = header_map[key]
    for r in range(2, ws.max_row + 1):
        if _norm(ws.cell(row=r, column=col).value) == target:
            return r
    return None


def _to_datetime(v: Any) -> Optional[datetime]:
    if v is None:
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, str):
        s = v.strip()
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(s, fmt)
            except Exception:
                pass
    return None


def _write(ws, headers: Dict[str, int], row: int, header_name: str, value: Any) -> None:
    key = _lower_norm(header_name)
    if key not in headers:
        return
    ws.cell(row=row, column=headers[key]).value = value


def add_forecast_to_inventory(config: AppConfig) -> None:
    run_date = datetime.now().date()  # date only

    # ---- Forecast workbook ----
    f_wb = load_workbook(config.forecast_workbook, data_only=True)
    f_ws = f_wb[f_wb.sheetnames[0]]

    erf = _norm(f_ws["F1"].value)
    if not erf:
        raise ValueError("ERF value was blank. Expected it in cell F1 of the forecast workbook.")

    f_headers = _build_header_map(f_ws, 1)
    for needed in FORECAST_TO_INVENTORY_MAP.keys():
        if _lower_norm(needed) not in f_headers:
            raise ValueError(f"Forecast sheet is missing required column header '{needed}' in row 1.")

    # ---- Declared Projects Portfolio workbook (NEW INPUT) ----
    p_wb = load_workbook(config.declared_project_portfolio_workbook, data_only=True)
    p_ws = _find_sheet_by_keywords(p_wb, ["declared", "portfolio"]) or p_wb[p_wb.sheetnames[0]]
    p_headers = _build_header_map(p_ws, 1)

    if "registry id" not in p_headers:
        raise ValueError(f"Declared Projects Portfolio sheet '{p_ws.title}' is missing 'Registry ID' column.")

    p_row = _find_row_by_value(p_ws, p_headers, "Registry ID", erf)
    if p_row is None:
        raise ValueError(f"Could not find ERF '{erf}' in Declared Projects Portfolio 'Registry ID' column.")

    # ---- Master inventory workbook ----
    m_wb = load_workbook(config.master_inventory_workbook)
    inventory_ws = _find_sheet_by_keywords(m_wb, ["inventory"]) or m_wb[m_wb.sheetnames[0]]
    i_headers = _ensure_headers(inventory_ws, REQUIRED_INVENTORY_HEADERS)

    # Iterate forecast rows and append inventory rows
    rp_col = f_headers[_lower_norm("RP Number")]
    rows_written = 0

    for r in range(2, f_ws.max_row + 1):
        rp_val = f_ws.cell(row=r, column=rp_col).value
        if _norm(rp_val) == "":
            continue

        out_row = inventory_ws.max_row + 1

        # portfolio -> inventory
        for field in PORTFOLIO_FIELDS:
            src_key = _lower_norm(field)
            if src_key not in p_headers:
                raise ValueError(f"Declared Projects Portfolio missing required column '{field}'.")

            val = p_ws.cell(row=p_row, column=p_headers[src_key]).value

            dest_field = field
            if src_key == "project id":
                dest_field = "Project Number"

            _write(inventory_ws, i_headers, out_row, dest_field, val)

        # forecast -> inventory
        for f_col_name, inv_col_name in FORECAST_TO_INVENTORY_MAP.items():
            val = f_ws.cell(row=r, column=f_headers[_lower_norm(f_col_name)]).value
            _write(inventory_ws, i_headers, out_row, inv_col_name, val)

        # derived dates based on RP End
        rp_end_val = f_ws.cell(row=r, column=f_headers[_lower_norm("RP End (EOM)")]).value
        rp_end_dt = _to_datetime(rp_end_val)
        if rp_end_dt is None:
            raise ValueError(f"Row {r}: RP End (EOM) is missing or not a date.")

        _write(inventory_ws, i_headers, out_row, "Forecasted Submission Date", (rp_end_dt + timedelta(days=2)).date())
        _write(inventory_ws, i_headers, out_row, "Date - Total Amount", (rp_end_dt + timedelta(days=92)).date())

        _write(inventory_ws, i_headers, out_row, "Status", "Forecasted")
        _write(inventory_ws, i_headers, out_row, "Data Update Date", run_date)

        rows_written += 1

    # Save master inventory output
    out_master = str(Path(config.save_master_inventory_output))
    Path(out_master).parent.mkdir(parents=True, exist_ok=True)
    m_wb.save(out_master)

    # Save forecast delta output (simple log)
    delta_wb = Workbook()
    delta_ws = delta_wb.active
    delta_ws.title = "Forecast Delta"
    delta_ws["A1"] = "Run Date"
    delta_ws["B1"] = "ERF (Registry ID)"
    delta_ws["C1"] = "Rows Added"
    delta_ws["A2"] = run_date
    delta_ws["B2"] = erf
    delta_ws["C2"] = rows_written

    out_delta = str(Path(config.save_forecast_delta_output))
    Path(out_delta).parent.mkdir(parents=True, exist_ok=True)
    delta_wb.save(out_delta)


def run_process(config: AppConfig) -> None:
    add_forecast_to_inventory(config)
