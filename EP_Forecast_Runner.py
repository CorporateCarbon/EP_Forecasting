import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from dataclasses import dataclass
from datetime import date
from Ep_Forecast_Engine import EngineConfig, run_engine


@dataclass
class AppConfig:
    starting_rp_number: int
    rp_length_months: int
    start_year: int
    start_month: int
    start_day: int
    discount_abatement: bool

    forecast_full_lifecycle: bool
    forecast_number_of_rps: int | None

    input_calculator_file: str
    save_raw_output: str
    save_aggregated_output: str


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Reporting Period Tool")
        self.geometry("680x470")
        self.resizable(False, False)

        self._build_ui()

    def _build_ui(self):
        pad = 10

        container = ttk.Frame(self, padding=pad)
        container.pack(fill="both", expand=True)

        # --- Inputs ---
        inputs = ttk.LabelFrame(container, text="Inputs", padding=pad)
        inputs.pack(fill="x", padx=0, pady=(0, pad))

        # Variables
        self.var_starting_rp = tk.StringVar(value="")
        self.var_rp_length = tk.StringVar(value="")
        self.var_start_year = tk.StringVar(value="")
        self.var_start_month = tk.StringVar(value="")
        self.var_start_day = tk.StringVar(value="")

        self.var_discount = tk.BooleanVar(value=False)

        # NEW: forecast controls
        self.var_forecast_full_lifecycle = tk.BooleanVar(value=False)
        self.var_forecast_num_rps = tk.StringVar(value="")

        self.var_input_file = tk.StringVar(value="")
        self.var_save_raw = tk.StringVar(value="")
        self.var_save_agg = tk.StringVar(value="")

        def add_labeled_entry(parent, row, label, var, width=18):
            ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=4)
            e = ttk.Entry(parent, textvariable=var, width=width)
            e.grid(row=row, column=1, sticky="w", pady=4)
            return e

        add_labeled_entry(inputs, 0, "Starting RP Number", self.var_starting_rp)
        add_labeled_entry(inputs, 1, "RP Length (Months)", self.var_rp_length)

        # Start date row
        ttk.Label(inputs, text="Start Date (Y / M / D)").grid(row=2, column=0, sticky="w", pady=4)
        date_frame = ttk.Frame(inputs)
        date_frame.grid(row=2, column=1, sticky="w", pady=4)

        ttk.Entry(date_frame, textvariable=self.var_start_year, width=6).pack(side="left")
        ttk.Label(date_frame, text=" / ").pack(side="left")
        ttk.Entry(date_frame, textvariable=self.var_start_month, width=4).pack(side="left")
        ttk.Label(date_frame, text=" / ").pack(side="left")
        ttk.Entry(date_frame, textvariable=self.var_start_day, width=4).pack(side="left")

        # Discount checkbox
        ttk.Label(
            inputs,
            text="Dates must match FullCAM Schedule (End of month). Above details are for first RP you are forecasting."
        ).grid(row=3, column=1, sticky="w", pady=6)

        # NEW: Forecast lifecycle checkbox
        ttk.Checkbutton(
            inputs,
            text="Forecast full Project Life Cycle",
            variable=self.var_forecast_full_lifecycle,
            command=self._toggle_forecast_num_rps
        ).grid(row=4, column=1, sticky="w", pady=6)

        # NEW: Forecast Number of RPs entry (disabled when lifecycle is checked)
        self.entry_forecast_num_rps = add_labeled_entry(
            inputs, 5, "Forecast Number of RPs", self.var_forecast_num_rps, width=18
        )

        # initialize enabled/disabled state correctly
        self._toggle_forecast_num_rps()

        # --- File pickers ---
        files = ttk.LabelFrame(container, text="Files", padding=pad)
        files.pack(fill="x", padx=0, pady=(0, pad))

        def add_file_picker(row, label, var, browse_cmd, button_text="Browse…"):
            ttk.Label(files, text=label).grid(row=row, column=0, sticky="w", pady=6)
            entry = ttk.Entry(files, textvariable=var, width=58)
            entry.grid(row=row, column=1, sticky="w", pady=6)
            ttk.Button(files, text=button_text, command=browse_cmd).grid(row=row, column=2, sticky="w", padx=6, pady=6)

        add_file_picker(
            0,
            "Input Calculator File",
            self.var_input_file,
            self._browse_input_file
        )
        add_file_picker(
            1,
            "Save Raw Output",
            self.var_save_raw,
            self._browse_save_raw,
            button_text="Save as…"
        )
        add_file_picker(
            2,
            "Save Aggregated Output",
            self.var_save_agg,
            self._browse_save_agg,
            button_text="Save as…"
        )

        # --- Actions ---
        actions = ttk.Frame(container)
        actions.pack(fill="x")

        ttk.Button(actions, text="Run", command=self._on_run).pack(side="right", padx=(6, 0))
        ttk.Button(actions, text="Quit", command=self.destroy).pack(side="right")

        self.status = tk.StringVar(value="Ready.")
        ttk.Label(container, textvariable=self.status).pack(anchor="w", pady=(pad, 0))

    def _toggle_forecast_num_rps(self):
        """Disable Forecast Number of RPs if forecasting full lifecycle."""
        if self.var_forecast_full_lifecycle.get():
            self.entry_forecast_num_rps.state(["disabled"])
        else:
            self.entry_forecast_num_rps.state(["!disabled"])

    # ---------- Browse handlers ----------
    def _browse_input_file(self):
        path = filedialog.askopenfilename(
            title="Select Input Calculator File",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm *.xls"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.var_input_file.set(path)

    def _browse_save_raw(self):
        path = filedialog.asksaveasfilename(
            title="Save Raw Output As",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.var_save_raw.set(path)

    def _browse_save_agg(self):
        path = filedialog.asksaveasfilename(
            title="Save Aggregated Output As",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.var_save_agg.set(path)

    # ---------- Run / validation ----------
    def _on_run(self):
        try:
            config = self._validate_and_build_config()
        except ValueError as e:
            messagebox.showerror("Invalid input", str(e))
            return

        try:
            self.status.set("Running...")
            self.update_idletasks()

            run_process(config)  # <-- plug in your logic

            self.status.set("Done.")
            messagebox.showinfo("Success", "Processing completed successfully.")
        except Exception as e:
            self.status.set("Error.")
            messagebox.showerror("Error", f"Something went wrong:\n\n{e}")

    def _validate_and_build_config(self) -> AppConfig:
        def as_int(name: str, s: str) -> int:
            s = (s or "").strip()
            if s == "":
                raise ValueError(f"{name} is required.")
            try:
                return int(s)
            except Exception:
                raise ValueError(f"{name} must be an integer.")

        starting_rp = as_int("Starting RP Number", self.var_starting_rp.get())
        rp_months = as_int("RP Length (Months)", self.var_rp_length.get())
        y = as_int("Start Year", self.var_start_year.get())
        m = as_int("Start Month", self.var_start_month.get())
        d = as_int("Start Day", self.var_start_day.get())

        if rp_months <= 0:
            raise ValueError("RP Length (Months) must be greater than 0.")

        try:
            _ = date(y, m, d)
        except Exception:
            raise ValueError("Start Date is not a valid calendar date (check Y/M/D).")

        forecast_full = bool(self.var_forecast_full_lifecycle.get())

        # Only require Forecast Number of RPs if NOT forecasting full lifecycle
        forecast_num = None
        if not forecast_full:
            forecast_num = as_int("Forecast Number of RPs", self.var_forecast_num_rps.get())
            if forecast_num <= 0:
                raise ValueError("Forecast Number of RPs must be greater than 0.")

        input_file = self.var_input_file.get().strip()
        save_raw = self.var_save_raw.get().strip()
        save_agg = self.var_save_agg.get().strip()

        if not input_file:
            raise ValueError("Input Calculator File is required.")
        if not save_raw:
            raise ValueError("Save Raw Output path is required.")
        if not save_agg:
            raise ValueError("Save Aggregated Output path is required.")

        return AppConfig(
            starting_rp_number=starting_rp,
            rp_length_months=rp_months,
            start_year=y,
            start_month=m,
            start_day=d,
            discount_abatement=bool(self.var_discount.get()),
            forecast_full_lifecycle=forecast_full,
            forecast_number_of_rps=forecast_num,
            input_calculator_file=input_file,
            save_raw_output=save_raw,
            save_aggregated_output=save_agg,
        )


def run_process(config: AppConfig):
    engine_config = EngineConfig(
        starting_rp_number=config.starting_rp_number,
        rp_length_months=config.rp_length_months,
        start_year=config.start_year,
        start_month=config.start_month,
        start_day=config.start_day,
        forecast_full_lifecycle=config.forecast_full_lifecycle,
        forecast_number_of_rps=config.forecast_number_of_rps,
        input_calculator_file=config.input_calculator_file,
        save_aggregated_output=config.save_aggregated_output,
    )
    run_engine(engine_config)



if __name__ == "__main__":
    app = App()
    app.mainloop()
