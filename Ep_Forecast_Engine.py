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
    save_raw_output: str


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

    def final_rp_end_from_project_end(project_end_date):
        # last day of the month before project_end_date's month
        first_of_end_month = project_end_date.replace(day=1)
        return first_of_end_month - timedelta(days=1)

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

    def rp_end_from_start(rp_start, rp_len_months):
        return rp_start + relativedelta(months=rp_len_months)

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
    # calc_mode_prev = app.calculation
    # app.calculation = "manual"  # faster; we explicitly calculate each iteration
    print("2")
    try:
        # Open calculator workbook once
        book = app.books.open(config.input_calculator_file)
        engine = ForecastEngineXL(book)
        final_rp_end = None
        # Decide n_rps
        rp_len = int(config.rp_length_months)
        print("3")
        if config.forecast_number_of_rps is not None:
            n_rps = int(config.forecast_number_of_rps)
        else:
            raw_project_start = engine.get_project_start_date()
            project_end = raw_project_start + relativedelta(years=25)

            final_rp_end = project_end.replace(day=1) - timedelta(days=1)

            current_end = month_end(datetime(config.start_year, config.start_month, config.start_day))
            months_to_end = (final_rp_end.year - current_end.year) * 12 + (final_rp_end.month - current_end.month)
            print("4")
            n_rps = months_to_end // rp_len
            if months_to_end % rp_len != 0:
                n_rps += 1

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
        for i in range(n_rps):
            print("Loop start")
            rp_num = start_rp_num + i
            next_rp_end = current_rp_start + relativedelta(months=rp_len)


            # Correctly manage final RP, RP length needs to be adjusted and end date correctly entered.
            if i == n_rps - 1:
                next_rp_end = final_rp_end   # override the last RP end
                rp_len = (
                    (final_rp_end.year - current_rp_start.year) * 12
                    + (final_rp_end.month - current_rp_start.month)
                )

            rp_end_dt, accus = engine.write_inputs_and_get_accus(
                rp_number=rp_num,
                rp_end_date=next_rp_end,
                rp_length_months=rp_len,
            )

            # Write row (row index in Excel = i+2)
            row = i + 2
            out_sheet.range((row, 1)).value = project_name
            out_sheet.range((row, 2)).value = registry_id
            out_sheet.range((row, 3)).value = rp_num
            out_sheet.range((row, 4)).value = current_rp_start
            out_sheet.range((row, 5)).value = rp_end_dt
            out_sheet.range((row, 6)).value = accus

            # advance
            current_rp_start = next_rp_end
            current_rp_end = next_rp_end

        # Save output
        out_book.save(str(out_path))
        out_book.close()
        print("loop close")

        # Save RAW output (calculator state after all RPs)
        raw_out = Path(config.save_raw_output)
        raw_out.parent.mkdir(parents=True, exist_ok=True)
        book.save(str(raw_out))

        # Optionally save calculator copy or just close
        book.close()

    finally:
        # restore settings and quit excel
        try:
            app.calculation = calc_mode_prev
        except Exception:
            pass
        app.quit()
