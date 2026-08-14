from __future__ import annotations

import calendar
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Tuple

import xlwings as xw
from dateutil.relativedelta import relativedelta
import time
import win32com.client as win32

from helpers.excel_paths import open_workbook, save_workbook


# ---------------- Config ----------------
@dataclass
class EngineConfig:
    starting_rp_number: int
    rp_length_months: int
    start_year: int
    start_month: int
    start_day: int

    forecast_full_lifecycle: bool
    forecast_number_of_rps: int | None

    input_calculator_file: str
    save_aggregated_output: str

    # PATCH (2026-08-14): explicit CER-registered crediting-period end.
    # ISO 'YYYY-MM-DD' string, or None to auto-read from the calculator
    # (Forecast_script_helper!D:'Crediting Period End' -> col E). This is the
    # AUTHORITATIVE forecast horizon; the old project_start+25yr inference is a
    # fallback only. No ACCUs are forecast for any RP ending after this date.
    crediting_period_end: str | None = None


# ---------------- Date helpers ----------------
def month_end(dt: datetime) -> datetime:
    last_day = calendar.monthrange(dt.year, dt.month)[1]
    return datetime(dt.year, dt.month, last_day)


def add_months_month_end(dt: datetime, months: int) -> datetime:
    # shift by N months then coerce to month-end
    return month_end(dt + relativedelta(months=months))


def excel_serial_to_datetime(val: float) -> datetime:
    # Excel serial date (1900 system): day 1 = 1900-01-01, but Excel has the 1900 leap-year bug.
    # xlwings usually returns datetime already, but this is a safe fallback.
    return datetime(1899, 12, 30) + timedelta(days=float(val))


# ---------------- Engine ----------------
class ForecastEngineXL:
    TARGET_SHEET = "Forecast_script_helper"

    # Column A labels
    LABEL_CURRENT_RP = "Current RP"
    LABEL_RP_END_YEAR = "Current RP End Year"
    LABEL_RP_END_MONTH = "current rp end month"
    LABEL_RP_END_DAY = "current rp end day"
    LABEL_RP_LENGTH = "RP Length"
    LABEL_ACCUS_REALISED = "ACCUs Realised"

    def __init__(self, book: xw.Book):
        self.book = book
        self.ws = self.book.sheets[self.TARGET_SHEET]

        # Build label index (Column A)
        self.label_row: Dict[str, int] = self._index_labels_col_a()

        # Validate required labels exist
        for required in (
            self.LABEL_CURRENT_RP,
            self.LABEL_RP_END_YEAR,
            self.LABEL_RP_END_MONTH,
            self.LABEL_RP_END_DAY,
            self.LABEL_RP_LENGTH,
            self.LABEL_ACCUS_REALISED,
        ):
            if self._norm(required) not in self.label_row:
                raise ValueError(f"Could not find label '{required}' in column A of '{self.TARGET_SHEET}'.")

    @staticmethod
    def _norm(x: Any) -> str:
        if x is None:
            return ""
        return str(x).strip().lower()

    @staticmethod
    def _strip_if_str(v: Any) -> Any:
        return v.strip() if isinstance(v, str) else v

    def _get_sheet_case_insensitive(self, wanted_name: str) -> xw.Sheet:
        """
        Returns a sheet matching wanted_name (case-insensitive).
        Raises a clear error if not found.
        """
        try:
            return self.book.sheets[wanted_name]
        except Exception:
            wanted_norm = wanted_name.strip().lower()
            for sh in self.book.sheets:
                if str(sh.name).strip().lower() == wanted_norm:
                    return sh
        raise ValueError(f"Could not find sheet '{wanted_name}' (case-insensitive) in workbook.")

    def get_project_metadata(self) -> tuple[Any, Any]:
        """
        Returns:
          project_name -> Calculator!A1
          registry_id  -> Calculator!B1
        """
        calc = self._get_sheet_case_insensitive("Forecast_script_helper")
        project_name = calc.range("A1").value
        registry_id = calc.range("B1").value
        print("Found name and Registry ID")
        return project_name, registry_id

    def _index_labels_col_a(self) -> Dict[str, int]:
        """
        Read column A values down to the last used cell and map normalized label -> row number.
        """
        # Get contiguous used range down from A1 (fast). If there are gaps, you can replace with used_range logic.
        colA = self.ws.range("A1:A300").value

        mapping: Dict[str, int] = {}
        if not isinstance(colA, list):
            colA = [colA]

        for idx, v in enumerate(colA, start=1):  # idx is row number
            key = self._norm(v)
            if key and key not in mapping:
                mapping[key] = idx
        return mapping

    def get_project_start_date(self) -> datetime:
        """
        Find 'Project Start Date' in column D and return corresponding column E value.
        """
        # Read col D down
        colD = self.ws.range("D1:D300").value
        if not isinstance(colD, list):
            colD = [colD]

        for idx, label in enumerate(colD, start=1):
            if label and str(label).strip().lower() == "project start date":
                val = self.ws.range((idx, 5)).value  # column E
                if isinstance(val, datetime):
                    return val
                if isinstance(val, (int, float)):
                    return excel_serial_to_datetime(val)
                raise ValueError("Project Start Date in column E is not a valid date.")
        raise ValueError("Could not find 'Project Start Date' in column D.")

    def get_crediting_period_end(self) -> datetime | None:
        """
        PATCH (2026-08-14): read the CER-registered crediting-period END date from
        'Crediting Period End' in column D -> column E of Forecast_script_helper.
        Returns None if the label is absent (caller then applies the +25yr fallback
        WITH A WARNING). This decouples the horizon from the 25-yr permanence proxy,
        which was the root cause of the Dogwood RP21-23 over-run.
        """
        colD = self.ws.range("D1:D300").value
        if not isinstance(colD, list):
            colD = [colD]
        for idx, label in enumerate(colD, start=1):
            if label and str(label).strip().lower() == "crediting period end":
                val = self.ws.range((idx, 5)).value  # column E
                if isinstance(val, datetime):
                    return val
                if isinstance(val, (int, float)):
                    return excel_serial_to_datetime(val)
                raise ValueError("Crediting Period End in column E is not a valid date.")
        return None

    def write_inputs_and_get_accus(
        self,
        rp_number: int,
        rp_end_date: datetime,
        rp_length_months: int,
    ) -> Tuple[datetime, Any]:
        """
        Writes inputs next to labels in column A (into column B), forces calc, returns:
          (rp_end_date datetime, ACCUs value in column B next to 'ACCUs Realised')
        """
        # Lookup row numbers
        r_rp = self.label_row[self._norm(self.LABEL_CURRENT_RP)]
        r_y = self.label_row[self._norm(self.LABEL_RP_END_YEAR)]
        r_m = self.label_row[self._norm(self.LABEL_RP_END_MONTH)]
        r_d = self.label_row[self._norm(self.LABEL_RP_END_DAY)]
        r_len = self.label_row[self._norm(self.LABEL_RP_LENGTH)]
        r_acc = self.label_row[self._norm(self.LABEL_ACCUS_REALISED)]

        # Write to column B (col=2) with whitespace stripping
        self.ws.range((r_rp, 2)).value = self._strip_if_str(rp_number)
        self.ws.range((r_y, 2)).value = self._strip_if_str(rp_end_date.year)
        self.ws.range((r_m, 2)).value = self._strip_if_str(rp_end_date.month)
        self.ws.range((r_d, 2)).value = self._strip_if_str(rp_end_date.day)
        self.ws.range((r_len, 2)).value = self._strip_if_str(rp_length_months)

        # Force calculation (critical)
        self.book.app.calculate()

        # Read ACCUs (column B next to label)
        accus_val = self.ws.range((r_acc, 2)).value

        return rp_end_date, accus_val


# ---------------- Runner ----------------
def run_engine(config: EngineConfig) -> None:
    # Ensure output folder exists
    out_path = Path(config.save_aggregated_output)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    print("1")
    # Start Excel (hidden)
    try:
        app = xw.App(visible=True, add_book=False)   # <-- start visible
        time.sleep(1)                                # <-- give Excel time to create window
        print("Excel started OK")
    except Exception as e:
        print("FAILED to start Excel")
        print(type(e), e)
        raise

    app.display_alerts = False
    app.screen_updating = False
    print("1.5")
    calc_mode_prev = None  # kept defined so the finally-block restore never NameErrors
    # calc_mode_prev = app.calculation
    # app.calculation = "manual"  # faster; we explicitly calculate each iteration
    print("2")
    try:
        # Open calculator workbook once
        book = open_workbook(app, config.input_calculator_file)
        engine = ForecastEngineXL(book)
        rp_len = int(config.rp_length_months)
        print("3")

        # PATCH (2026-08-14): resolve the crediting-period end ONCE, from an explicit
        # source, and apply it as a hard cap in BOTH modes (see the loop below).
        # Priority: config override -> calculator 'Crediting Period End' -> +25yr fallback.
        crediting_end = None
        if config.crediting_period_end:
            crediting_end = datetime.strptime(config.crediting_period_end, "%Y-%m-%d")
        else:
            crediting_end = engine.get_crediting_period_end()
        if crediting_end is None:
            raw_project_start = engine.get_project_start_date()
            crediting_end = (raw_project_start + relativedelta(years=25)).replace(day=1) - timedelta(days=1)
            print(f"WARNING: no 'Crediting Period End' found; falling back to project_start+25yr "
                  f"-> {crediting_end.date()}. This is the 25-yr assumption, NOT the registered "
                  f"crediting period. Add 'Crediting Period End' to Forecast_script_helper to make it authoritative.")
        crediting_end = month_end(crediting_end) if crediting_end.day >= 28 else crediting_end
        print(f"Crediting-period end (cap): {crediting_end.date()}")

        # final_rp_end always known now; used to truncate the last RP to a partial period.
        final_rp_end = crediting_end

        # Upper bound on RP count. The in-loop clamp is what actually enforces the cap,
        # so fixed-count mode can never over-run the crediting period either.
        current_end = month_end(datetime(config.start_year, config.start_month, config.start_day))
        months_to_end = (final_rp_end.year - current_end.year) * 12 + (final_rp_end.month - current_end.month)
        n_rps_by_end = months_to_end // rp_len + (1 if months_to_end % rp_len != 0 else 0)
        print("4")
        if config.forecast_number_of_rps is not None:
            # honour the user's count, but never beyond the crediting period
            n_rps = min(int(config.forecast_number_of_rps), n_rps_by_end)
        else:
            n_rps = n_rps_by_end

        # Create aggregated workbook (also via xlwings so saving is easy)
        out_book = app.books.add()
        out_sheet = out_book.sheets[0]
        out_sheet.name = "Aggregated"
        print("5")
        # Headers
        project_name, registry_id = engine.get_project_metadata()

        out_sheet.range("A1").value = [
            "Name",
            "Registry ID",
            "RP",
            "Reporting Period - Start",
            "Reporting Period - End",
            "ACCUs Realised",
        ]

        # Starting dates
        start_rp_num = int(config.starting_rp_number)
        current_rp_end = month_end(datetime(config.start_year, config.start_month, config.start_day))
        current_rp_start = datetime(config.start_year, config.start_month, config.start_day)        
        print("6")
        # Loop RPs
        rows_written = 0
        last_rp_end = None
        for i in range(n_rps):
            print("Loop start")
            rp_num = start_rp_num + i
            this_rp_len = rp_len
            next_rp_end = current_rp_start + relativedelta(months=this_rp_len)

            # PATCH (2026-08-14): AUTHORITATIVE crediting-period clamp, applied in EVERY
            # mode. If this RP would end on/after the crediting-period end, truncate it to
            # a PARTIAL period ending exactly on that date, emit it, then STOP. This is the
            # single guard that prevents forecasting ACCUs beyond the crediting period
            # (the Dogwood RP21-23 defect) and produces the correct partial final RP
            # (e.g. RP20 = 1 Jan - 30 Jun 2036) regardless of lifecycle/fixed-count.
            is_final = False
            if next_rp_end >= crediting_end:
                next_rp_end = crediting_end
                this_rp_len = (
                    (crediting_end.year - current_rp_start.year) * 12
                    + (crediting_end.month - current_rp_start.month)
                )
                is_final = True

            rp_end_dt, accus = engine.write_inputs_and_get_accus(
                rp_number=rp_num,
                rp_end_date=next_rp_end,
                rp_length_months=this_rp_len,
            )

            # Write row (row index in Excel = i+2)
            row = i + 2
            out_sheet.range((row, 1)).value = project_name
            out_sheet.range((row, 2)).value = registry_id
            out_sheet.range((row, 3)).value = rp_num
            out_sheet.range((row, 4)).value = current_rp_start
            out_sheet.range((row, 5)).value = rp_end_dt
            out_sheet.range((row, 6)).value = accus
            rows_written += 1
            last_rp_end = rp_end_dt

            # advance
            current_rp_start = next_rp_end
            current_rp_end = next_rp_end

            if is_final:
                print(f"Reached crediting-period end {crediting_end.date()} at RP {rp_num}; stopping.")
                break

        # PATCH (2026-08-14): QC assertion - no emitted RP may end after the crediting period.
        # Cheap regression guard against a silent over-run.
        if last_rp_end is not None and last_rp_end > crediting_end:
            raise AssertionError(
                f"QC FAIL: final RP end {last_rp_end.date()} is after crediting-period end "
                f"{crediting_end.date()}. Aborting to avoid crediting beyond the registered period."
            )
        print(f"QC OK: {rows_written} RP(s) written; last RP ends {last_rp_end.date() if last_rp_end else 'n/a'} "
              f"(<= crediting end {crediting_end.date()}).")

        # Save output
        save_workbook(out_book, out_path)
        out_book.close()
        print("loop close")

        # Save RAW output (calculator state after all RPs)

        # Optionally save calculator copy or just close
        book.close()

    finally:
        # restore settings and quit excel
        try:
            if calc_mode_prev is not None:
                app.calculation = calc_mode_prev
        except Exception:
            pass
        app.quit()
