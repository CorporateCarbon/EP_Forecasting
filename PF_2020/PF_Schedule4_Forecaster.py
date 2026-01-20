# -*- coding: utf-8 -*-
#%%##
from __future__ import annotations
from dataclasses import dataclass
from datetime import date, datetime
from typing import List, Tuple
from dateutil.relativedelta import relativedelta
import csv
#%%#
# ------------------ PARAMETERS (EDIT THESE) ------------------
CBASE = 596.78 #value from b21
CLT   = 184419.56   #Value from B28
PERMANENCE_YEARS = 25
DPP = 0.75 if PERMANENCE_YEARS == 25 else 1.0

CEA_START = date(2021, 6, 25)
FIRST_RP_START = date(2025, 10, 31)
FIRST_RP_END   = date(2026, 6, 30)
MODELLING_END_RP_END = date(2046, 6, 30)

FIRST_YEAR_DEDUCTION = 12821.0  # in ACCUs (CO2-e)
APPLY_C_TO_CO2_CONVERSION = True
AREA_HA = None  # multiply outputs by area if CBASE/CLT are per-ha

#%%##
# ------------------ CORE LOGIC ------------------
def months_completed(d1: date, d2: date) -> int:
    if d2 < d1:  # ensure order
        return -months_completed(d2, d1)
    y = d2.year - d1.year
    m = d2.month - d1.month
    total = y * 12 + m
    if d2.day < d1.day:
        total -= 1
    return total

def cp_at_months(cbase: float, clt: float, n_months: int, dpp: float) -> float:
    n = min(max(n_months, 0), 180)
    return cbase + (n/180.0) * (clt - cbase) * dpp


@dataclass
class RPResult:
    rp_index: int
    rp_start: date
    rp_end: date
    n_months: int
    cp_c: float
    annual_credit_c: float
    annual_credit_c_adj: float


def generate_rp_schedule(first_start: date, first_end: date, last_end: date) -> List[Tuple[date, date]]:
    schedule = []
    rp_start = first_start
    rp_end = first_end
    while rp_end <= last_end:
        schedule.append((rp_start, rp_end))
        rp_start = rp_start + relativedelta(years=1)
        rp_end   = rp_end   + relativedelta(years=1)
    return schedule

def forecast() -> List[RPResult]:
    schedule = generate_rp_schedule(FIRST_RP_START, FIRST_RP_END, MODELLING_END_RP_END)
    results: List[RPResult] = []
    prev_cp_c = None
    cum_c_adj = 0.0

    for i, (rp_start, rp_end) in enumerate(schedule, start=1):
        n = months_completed(CEA_START, rp_end)
        cp_c_raw = cp_at_months(CBASE, CLT, n, DPP)
        annual_c_raw = cp_c_raw if prev_cp_c is None else (cp_c_raw - prev_cp_c)

        # Apply first-year deduction (in same units as annual_c_raw) if applicable
        annual_after_first_deduction = annual_c_raw
        if i == 1 and FIRST_YEAR_DEDUCTION:
            annual_after_first_deduction = annual_c_raw - FIRST_YEAR_DEDUCTION

        # Apply permanence reduction (0.75) if PERMANENCE_YEARS == 25
        if PERMANENCE_YEARS == 25:
            annual_c_adj = annual_after_first_deduction * 0.75
        else:
            annual_c_adj = annual_after_first_deduction

        cum_c_adj += annual_c_adj

        results.append(RPResult(
            rp_index=i,
            rp_start=rp_start, rp_end=rp_end,
            n_months=n,
            cp_c=cp_c_raw,
            annual_credit_c=annual_c_raw,
            annual_credit_c_adj=annual_c_adj,
        ))

        prev_cp_c = cp_c_raw

    return results

def write_csv(path: str, rows: List[RPResult]) -> str:
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow([
            "RP","RP_Start","RP_End","Months_since_start",
            "CP_C","Annual_Credit_C","Annual_Credit_C_Adj",
        ])
        for r in rows:
            w.writerow([
                r.rp_index, r.rp_start.isoformat(), r.rp_end.isoformat(), r.n_months,
                round(r.cp_c,6), round(r.annual_credit_c,6), round(r.annual_credit_c_adj,6)
            ])
    return path
#%%##
def main():
    rows = forecast()
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    out = f"accu_15yr_forecast_{ts}.csv"
    path = write_csv(out, rows)

    print(f"\nSaved: {path}\n")
    
if __name__ == '__main__':
    main()

# %%
