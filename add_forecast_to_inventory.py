from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta, date
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

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

def _to_datetime(v):
    """
    ALWAYS returns a date or None.
    Handles Excel dates, datetimes, ISO strings, and Excel-serial numbers.
    """
    if v is None:
        return None

    if isinstance(v, datetime):
        return v.date()

    if isinstance(v, date):
        return v

    # Excel serial date
    if isinstance(v, (int, float)):
        try:
            return (datetime(1899, 12, 30) + timedelta(days=int(v))).date()
        except Exception:
            return None

    if isinstance(v, str):
        s = v.strip()
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(s, fmt).date()
            except Exception:
                pass

    return None


def _write(ws, headers: Dict[str, int], row: int, header_name: str, value: Any) -> None:
    key = _lower_norm(header_name)
    if key not in headers:
        return
    ws.cell(row=row, column=headers[key]).value = value

def _row_as_list(ws, row: int, max_col: int) -> List[Any]:
    return [ws.cell(row=row, column=c).value for c in range(1, max_col + 1)]

def _append_table(ws_out, headers: List[str], rows: List[List[Any]]) -> None:
    for c, h in enumerate(headers, start=1):
        ws_out.cell(row=1, column=c).value = h
    for r_idx, row_vals in enumerate(rows, start=2):
        for c, v in enumerate(row_vals, start=1):
            ws_out.cell(row=r_idx, column=c).value = v


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
            raise ValueError(f"Forecast sheet missing required header '{needed}' in row 1.")

    # Determine cutoff = first forecast RP Start (minimum RP Start among forecast rows)
    rp_start_col = f_headers[_lower_norm("RP Start (EOM)")]
    rp_num_col = f_headers[_lower_norm("RP Number")]

    forecast_rows: List[Dict[str, Any]] = []
    cutoff_start_dt: Optional[datetime] = None

    for r in range(2, f_ws.max_row + 1):
        rp_val = f_ws.cell(row=r, column=rp_num_col).value
        if _norm(rp_val) == "":
            continue

        start_val = f_ws.cell(row=r, column=rp_start_col).value
        start_dt = _to_datetime(start_val)
        if start_dt is None:
            raise ValueError(f"Row {r}: RP Start (EOM) missing or not a date.")

        if cutoff_start_dt is None or start_dt < cutoff_start_dt:
            cutoff_start_dt = start_dt

        forecast_rows.append({
            "row_index": r,
            "RP Number": f_ws.cell(row=r, column=f_headers[_lower_norm("RP Number")]).value,
            "RP Start (EOM)": start_dt,
            "RP End (EOM)": _to_datetime(f_ws.cell(row=r, column=f_headers[_lower_norm("RP End (EOM)")]).value),
            "ACCUs Realised": f_ws.cell(row=r, column=f_headers[_lower_norm("ACCUs Realised")]).value,
        })

    if cutoff_start_dt is None:
        raise ValueError("No forecast data rows found (RP Number column was empty).")

    # ---- Declared Projects Portfolio workbook ----
    p_wb = load_workbook(config.declared_project_portfolio_workbook, data_only=True)
    p_ws = _find_sheet_by_keywords(p_wb, ["declared", "portfolio"]) or p_wb[p_wb.sheetnames[0]]
    p_headers = _build_header_map(p_ws, 1)

    if "registry id" not in p_headers:
        raise ValueError(f"Declared Projects Portfolio sheet '{p_ws.title}' missing 'Registry ID' column.")

    p_row = _find_row_by_value(p_ws, p_headers, "Registry ID", erf)
    if p_row is None:
        raise ValueError(f"Could not find ERF '{erf}' in Declared Projects Portfolio 'Registry ID'.")

    # ---- Master inventory workbook ----
    m_wb = load_workbook(config.master_inventory_workbook)
    inventory_ws = _find_sheet_by_keywords(m_wb, ["inventory"]) or m_wb[m_wb.sheetnames[0]]
    i_headers = _ensure_headers(inventory_ws, REQUIRED_INVENTORY_HEADERS)

    # REQUIRED columns for deletion logic
    if _lower_norm("Registry ID") not in i_headers:
        raise ValueError("Inventory sheet missing 'Registry ID' column.")
    if _lower_norm("Reporting Period - End") not in i_headers:
        raise ValueError("Inventory sheet missing 'Reporting Period - End' column.")

    inv_registry_col = i_headers[_lower_norm("Registry ID")]
    inv_rpend_col = i_headers[_lower_norm("Reporting Period - End")]

    # --- Snapshot: all inventory rows where ERF matches Registry ID (BEFORE deletion) ---
    inventory_snapshot_headers = [inventory_ws.cell(row=1, column=c).value for c in range(1, inventory_ws.max_column + 1)]
    inv_match_rows_before: List[List[Any]] = []
    for r in range(2, inventory_ws.max_row + 1):
        if _norm(inventory_ws.cell(row=r, column=inv_registry_col).value) == erf:
            inv_match_rows_before.append(_row_as_list(inventory_ws, r, inventory_ws.max_column))

    # --- Delete rows: Registry ID == ERF AND Reporting Period - End > cutoff_start_dt ---
    deleted_rows: List[List[Any]] = []
    kept_rows: List[List[Any]] = []

    # iterate bottom-up so deletes don't shift later rows
    for r in range(inventory_ws.max_row, 1, -1):
        if _norm(inventory_ws.cell(row=r, column=inv_registry_col).value) != erf:
            continue

        rp_end_dt = _to_datetime(inventory_ws.cell(row=r, column=inv_rpend_col).value)
        # If RP End is blank/unparseable, treat as NOT deletable (kept)
        if rp_end_dt is None:
            kept_rows.append(_row_as_list(inventory_ws, r, inventory_ws.max_column))
            continue

        if rp_end_dt > cutoff_start_dt:
            deleted_rows.append(_row_as_list(inventory_ws, r, inventory_ws.max_column))
            inventory_ws.delete_rows(r, 1)
        else:
            kept_rows.append(_row_as_list(inventory_ws, r, inventory_ws.max_column))

    # kept_rows/deleted_rows were collected bottom-up; reverse for human readability
    kept_rows.reverse()
    deleted_rows.reverse()

    # --- Append new inventory rows for every forecast row ---
    rows_written = 0

    for fr in forecast_rows:
        if fr["RP End (EOM)"] is None:
            raise ValueError(f"Row {fr['row_index']}: RP End (EOM) missing or not a date.")

        out_row = inventory_ws.max_row + 1

        # portfolio -> inventory
        for field in PORTFOLIO_FIELDS:
            src_key = _lower_norm(field)
            if src_key not in p_headers:
                raise ValueError(f"Declared Projects Portfolio missing required column '{field}'.")

            val = p_ws.cell(row=p_row, column=p_headers[src_key]).value
            dest_field = "Project Number" if src_key == "project id" else field
            _write(inventory_ws, i_headers, out_row, dest_field, val)

        # forecast -> inventory
        _write(inventory_ws, i_headers, out_row, "RP", fr["RP Number"])
        _write(inventory_ws, i_headers, out_row, "Reporting Period - Start", fr["RP Start (EOM)"].date())
        _write(inventory_ws, i_headers, out_row, "Reporting Period - End", fr["RP End (EOM)"].date())
        _write(inventory_ws, i_headers, out_row, "Total Amount (ACCUs)", fr["ACCUs Realised"])

        # derived
        rp_end_dt = fr["RP End (EOM)"]
        _write(inventory_ws, i_headers, out_row, "Forecasted Submission Date", (rp_end_dt + timedelta(days=2)).date())
        _write(inventory_ws, i_headers, out_row, "Date - Total Amount", (rp_end_dt + timedelta(days=92)).date())

        # fixed
        _write(inventory_ws, i_headers, out_row, "Status", "Forecasted")
        _write(inventory_ws, i_headers, out_row, "Data Update Date", run_date)

        rows_written += 1

    # --- Save master inventory output ---
    out_master = str(Path(config.save_master_inventory_output))
    Path(out_master).parent.mkdir(parents=True, exist_ok=True)
    m_wb.save(out_master)

    # --- Forecast delta workbook with 3 sheets ---
    delta_wb = Workbook()

    # Sheet 1: Summary
    s1 = delta_wb.active
    s1.title = "Summary"
    s1["A1"] = "Run Date"
    s1["B1"] = "ERF (Registry ID)"
    s1["C1"] = "Cutoff (First RP Start)"
    s1["D1"] = "Rows Deleted"
    s1["E1"] = "Rows Added"
    s1["A2"] = run_date
    s1["B2"] = erf
    s1["C2"] = cutoff_start_dt.date()
    s1["D2"] = len(deleted_rows)
    s1["E2"] = rows_written

    # Sheet 2: All ERF-matching inventory rows (before deletion)
    s2 = delta_wb.create_sheet("Inventory ERF Rows")
    _append_table(s2, inventory_snapshot_headers, inv_match_rows_before)

    # Sheet 3: Not deleted + forecasts that replace deleted
    # We'll store:
    #   - kept inventory rows (those NOT deleted)
    #   - then a blank line
    #   - then a small forecast table
    s3 = delta_wb.create_sheet("Keep vs Forecast Replace")

    # kept rows section
    s3["A1"] = "KEPT inventory rows (Registry ID == ERF, RP End <= cutoff)"
    if kept_rows:
        # headers on row 2, data row 3+
        for c, h in enumerate(inventory_snapshot_headers, start=1):
            s3.cell(row=2, column=c).value = h
        for i, row_vals in enumerate(kept_rows, start=3):
            for c, v in enumerate(row_vals, start=1):
                s3.cell(row=i, column=c).value = v
        start_forecast_block_row = 3 + len(kept_rows) + 2
    else:
        start_forecast_block_row = 4

    # forecast section
    s3.cell(row=start_forecast_block_row, column=1).value = "FORECAST rows to be written (replacement set)"
    forecast_headers = ["RP Number", "RP Start (EOM)", "RP End (EOM)", "ACCUs Realised"]
    for c, h in enumerate(forecast_headers, start=1):
        s3.cell(row=start_forecast_block_row + 1, column=c).value = h

    for i, fr in enumerate(forecast_rows, start=0):
        rr = start_forecast_block_row + 2 + i
        s3.cell(row=rr, column=1).value = fr["RP Number"]
        s3.cell(row=rr, column=2).value = fr["RP Start (EOM)"]
        s3.cell(row=rr, column=3).value = fr["RP End (EOM)"]
        s3.cell(row=rr, column=4).value = fr["ACCUs Realised"]

    # Save delta output
    out_delta = str(Path(config.save_forecast_delta_output))
    Path(out_delta).parent.mkdir(parents=True, exist_ok=True)
    delta_wb.save(out_delta)


def run_process(config: AppConfig) -> None:
    add_forecast_to_inventory(config)
