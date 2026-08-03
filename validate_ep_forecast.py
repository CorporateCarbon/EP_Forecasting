#!/usr/bin/env python3
"""
Independent validation of the EP forecaster.

Rebuilds ACCUs Realised for Blackwood from scratch, straight from the raw FullCAM
carbon time series in each CEA sheet -- with no reference to the calculator's own
formulas -- and compares to the tool's expected output workbook.

Method reproduced:
    stock(date)   = sum over CEAs of  area_ha * (C mass of trees + forest debris) * 44/12
    total_abate   = stock(RP end) - stock(previous RP end)
    ACCUs Realised = total_abate * (1 - 0.20 permanence - 0.05 buffer)   # = *0.75

Run from the repo root:  python validate_ep_forecast.py
Requires: openpyxl.  (No Excel needed.)
"""
import os, openpyxl

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                    "test_inputs_&_Outputs", "forecaster_input_&_output")
CALC = os.path.join(BASE, "260121_BLK_Calculator_Updated.xlsx")
EXPECTED = os.path.join(BASE, "Blackwood_Forecast_OUTPUT_EXAMPLE.xlsx")
DISCOUNT = 1 - 0.20 - 0.05  # 0.75


def main():
    wb = openpyxl.load_workbook(CALC, data_only=True)

    ab = wb["Abatement"]
    areas = {}
    for r in range(5, 12):
        name, area = ab.cell(r, 1).value, ab.cell(r, 3).value
        if name:
            areas[name] = area

    def series(sheet):
        m = {}
        for r in range(10, sheet.max_row + 1):
            y, mo = sheet.cell(r, 1).value, sheet.cell(r, 2).value
            trees, deb = sheet.cell(r, 5).value, sheet.cell(r, 6).value
            if isinstance(y, (int, float)) and isinstance(mo, (int, float)) and isinstance(trees, (int, float)):
                m[(int(y), int(mo))] = (trees or 0) + (deb or 0)
        return m

    cea_series = {name: series(wb[name]) for name in areas}

    def stock(year, month):
        tot = 0.0
        for name, area in areas.items():
            v = cea_series[name].get((year, month))
            if v is None:
                return None
            tot += area * v * 44 / 12
        return tot

    # expected output
    ewb = openpyxl.load_workbook(EXPECTED, data_only=True)
    ews = ewb["Aggregated"]
    print(f"{'RP':>3} {'reproduced':>12} {'tool output':>12} {'diff':>12}")
    ok = True
    for r in range(2, ews.max_row + 1):
        rp = ews.cell(r, 3).value
        end = ews.cell(r, 5).value
        exp = ews.cell(r, 6).value
        if rp is None or end is None or exp is None:
            continue
        s_end, s_prev = stock(end.year, end.month), stock(end.year - 1, end.month)
        if s_end is None or s_prev is None:
            print(f"{rp:>3}  (no FullCAM data for this RP window)")
            continue
        mine = (s_end - s_prev) * DISCOUNT
        diff = mine - exp
        if abs(diff) > 1e-6:
            ok = False
        print(f"{rp:>3} {mine:12.3f} {exp:12.3f} {diff:12.6f}")
    print("\nRESULT:", "MATCH - forecaster validated" if ok else "MISMATCH - investigate")


if __name__ == "__main__":
    main()
