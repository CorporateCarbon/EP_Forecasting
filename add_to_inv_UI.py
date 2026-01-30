import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from dataclasses import dataclass


@dataclass
class AppConfig:
    forecast_workbook: str
    master_inventory_workbook: str
    save_master_inventory_output: str
    save_forecast_delta_output: str


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Forecast Runner")
        self.geometry("720x320")
        self.resizable(False, False)

        self._build_ui()

    def _build_ui(self):
        pad = 10

        container = ttk.Frame(self, padding=pad)
        container.pack(fill="both", expand=True)

        # --- Files ---
        files = ttk.LabelFrame(container, text="Files", padding=pad)
        files.pack(fill="x", padx=0, pady=(0, pad))

        self.var_forecast_file = tk.StringVar(value="")
        self.var_master_inventory_file = tk.StringVar(value="")
        self.var_save_master_inventory = tk.StringVar(value="")
        self.var_save_forecast_delta = tk.StringVar(value="")

        def add_file_picker(row, label, var, browse_cmd, button_text="Browse…"):
            ttk.Label(files, text=label).grid(row=row, column=0, sticky="w", pady=6)
            entry = ttk.Entry(files, textvariable=var, width=60)
            entry.grid(row=row, column=1, sticky="w", pady=6)
            ttk.Button(files, text=button_text, command=browse_cmd).grid(
                row=row, column=2, sticky="w", padx=6, pady=6
            )

        add_file_picker(
            0,
            "Forecast Workbook",
            self.var_forecast_file,
            self._browse_forecast_file,
        )
        add_file_picker(
            1,
            "Master Inventory Workbook",
            self.var_master_inventory_file,
            self._browse_master_inventory_file,
        )
        add_file_picker(
            2,
            "Save Output Master Inventory",
            self.var_save_master_inventory,
            self._browse_save_master_inventory,
            button_text="Save as…",
        )
        add_file_picker(
            3,
            "Save Forecast Delta",
            self.var_save_forecast_delta,
            self._browse_save_forecast_delta,
            button_text="Save as…",
        )

        # --- Actions ---
        actions = ttk.Frame(container)
        actions.pack(fill="x")

        ttk.Button(actions, text="Run", command=self._on_run).pack(side="right", padx=(6, 0))
        ttk.Button(actions, text="Quit", command=self.destroy).pack(side="right")

        self.status = tk.StringVar(value="Ready.")
        ttk.Label(container, textvariable=self.status).pack(anchor="w", pady=(pad, 0))

    # ---------- Browse handlers ----------
    def _browse_forecast_file(self):
        path = filedialog.askopenfilename(
            title="Select Forecast Workbook",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm *.xls"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.var_forecast_file.set(path)

    def _browse_master_inventory_file(self):
        path = filedialog.askopenfilename(
            title="Select Master Inventory Workbook",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm *.xls"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.var_master_inventory_file.set(path)

    def _browse_save_master_inventory(self):
        path = filedialog.asksaveasfilename(
            title="Save Output Master Inventory As",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.var_save_master_inventory.set(path)

    def _browse_save_forecast_delta(self):
        path = filedialog.asksaveasfilename(
            title="Save Forecast Delta As",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.var_save_forecast_delta.set(path)

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

            run_process(config)  # <-- plug in your logic here

            self.status.set("Done.")
            messagebox.showinfo("Success", "Processing completed successfully.")
        except Exception as e:
            self.status.set("Error.")
            messagebox.showerror("Error", f"Something went wrong:\n\n{e}")

    def _validate_and_build_config(self) -> AppConfig:
        forecast_file = self.var_forecast_file.get().strip()
        master_file = self.var_master_inventory_file.get().strip()
        save_master = self.var_save_master_inventory.get().strip()
        save_delta = self.var_save_forecast_delta.get().strip()

        if not forecast_file:
            raise ValueError("Forecast Workbook is required.")
        if not master_file:
            raise ValueError("Master Inventory Workbook is required.")
        if not save_master:
            raise ValueError("Save Output Master Inventory path is required.")
        if not save_delta:
            raise ValueError("Save Forecast Delta path is required.")

        return AppConfig(
            forecast_workbook=forecast_file,
            master_inventory_workbook=master_file,
            save_master_inventory_output=save_master,
            save_forecast_delta_output=save_delta,
        )


def run_process(config: AppConfig):
    """
    Hook your real processing logic here.

    You will typically:
      - load forecast workbook (config.forecast_workbook)
      - load master inventory workbook (config.master_inventory_workbook)
      - produce updated master inventory -> save to config.save_master_inventory_output
      - produce forecast delta -> save to config.save_forecast_delta_output
    """
    # Example placeholder:
    # from your_engine import run_forecast_update
    # run_forecast_update(
    #     forecast_path=config.forecast_workbook,
    #     master_inventory_path=config.master_inventory_workbook,
    #     save_master_inventory_path=config.save_master_inventory_output,
    #     save_forecast_delta_path=config.save_forecast_delta_output,
    # )
    pass


if __name__ == "__main__":
    app = App()
    app.mainloop()
