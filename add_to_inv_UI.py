# add_to_inv_UI.py  (updated)

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from dataclasses import dataclass

# IMPORTANT: this imports the engine behind the UI
from add_forecast_to_inventory import run_process


@dataclass
class AppConfig:
    forecast_workbook: str
    master_inventory_workbook: str
    declared_project_portfolio_workbook: str
    save_master_inventory_output: str
    save_forecast_delta_output: str


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Add Forecast to Inventory")
        self.geometry("760x360")
        self.resizable(False, False)
        self._build_ui()

    def _build_ui(self):
        pad = 10

        container = ttk.Frame(self, padding=pad)
        container.pack(fill="both", expand=True)

        files = ttk.LabelFrame(container, text="Files", padding=pad)
        files.pack(fill="x", padx=0, pady=(0, pad))

        self.var_forecast_file = tk.StringVar(value="")
        self.var_master_inventory_file = tk.StringVar(value="")
        self.var_declared_portfolio_file = tk.StringVar(value="")
        self.var_save_master_inventory = tk.StringVar(value="")
        self.var_save_forecast_delta = tk.StringVar(value="")

        def add_file_picker(row, label, var, browse_cmd, button_text="Browse…"):
            ttk.Label(files, text=label).grid(row=row, column=0, sticky="w", pady=6)
            ttk.Entry(files, textvariable=var, width=60).grid(row=row, column=1, sticky="w", pady=6)
            ttk.Button(files, text=button_text, command=browse_cmd).grid(
                row=row, column=2, sticky="w", padx=6, pady=6
            )

        add_file_picker(0, "Forecast Workbook", self.var_forecast_file, self._browse_forecast_file)
        add_file_picker(1, "Master Inventory Workbook", self.var_master_inventory_file, self._browse_master_inventory_file)

        # NEW INPUT
        add_file_picker(2, "Declared Projects Portfolio", self.var_declared_portfolio_file, self._browse_declared_portfolio_file)

        add_file_picker(
            3,
            "Save Output Master Inventory",
            self.var_save_master_inventory,
            self._browse_save_master_inventory,
            button_text="Save as…",
        )
        add_file_picker(
            4,
            "Save Forecast Delta",
            self.var_save_forecast_delta,
            self._browse_save_forecast_delta,
            button_text="Save as…",
        )

        actions = ttk.Frame(container)
        actions.pack(fill="x")

        ttk.Button(actions, text="Run", command=self._on_run).pack(side="right", padx=(6, 0))
        ttk.Button(actions, text="Quit", command=self.destroy).pack(side="right")

        self.status = tk.StringVar(value="Ready.")
        ttk.Label(container, textvariable=self.status).pack(anchor="w", pady=(pad, 0))

    def _browse_forecast_file(self):
        path = filedialog.askopenfilename(
            title="Select Forecast Workbook",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if path:
            self.var_forecast_file.set(path)

    def _browse_master_inventory_file(self):
        path = filedialog.askopenfilename(
            title="Select Master Inventory Workbook",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if path:
            self.var_master_inventory_file.set(path)

    def _browse_declared_portfolio_file(self):
        path = filedialog.askopenfilename(
            title="Select Declared Projects Portfolio Workbook",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if path:
            self.var_declared_portfolio_file.set(path)

    def _browse_save_master_inventory(self):
        path = filedialog.asksaveasfilename(
            title="Save Output Master Inventory As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.var_save_master_inventory.set(path)

    def _browse_save_forecast_delta(self):
        path = filedialog.asksaveasfilename(
            title="Save Forecast Delta As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.var_save_forecast_delta.set(path)

    def _on_run(self):
        try:
            config = self._validate_and_build_config()
        except ValueError as e:
            messagebox.showerror("Invalid input", str(e))
            return

        try:
            self.status.set("Running...")
            self.update_idletasks()

            run_process(config)  # calls add_forecast_to_inventory

            self.status.set("Done.")
            messagebox.showinfo("Success", "Processing completed successfully.")
        except Exception as e:
            self.status.set("Error.")
            messagebox.showerror("Error", f"Something went wrong:\n\n{e}")

    def _validate_and_build_config(self) -> AppConfig:
        forecast_file = self.var_forecast_file.get().strip()
        master_file = self.var_master_inventory_file.get().strip()
        portfolio_file = self.var_declared_portfolio_file.get().strip()
        save_master = self.var_save_master_inventory.get().strip()
        save_delta = self.var_save_forecast_delta.get().strip()

        if not forecast_file:
            raise ValueError("Forecast Workbook is required.")
        if not master_file:
            raise ValueError("Master Inventory Workbook is required.")
        if not portfolio_file:
            raise ValueError("Declared Projects Portfolio Workbook is required.")
        if not save_master:
            raise ValueError("Save Output Master Inventory path is required.")
        if not save_delta:
            raise ValueError("Save Forecast Delta path is required.")

        return AppConfig(
            forecast_workbook=forecast_file,
            master_inventory_workbook=master_file,
            declared_project_portfolio_workbook=portfolio_file,
            save_master_inventory_output=save_master,
            save_forecast_delta_output=save_delta,
        )


if __name__ == "__main__":
    app = App()
    app.mainloop()
