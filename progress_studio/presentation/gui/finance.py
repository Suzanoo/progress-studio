from __future__ import annotations

from datetime import date
from pathlib import Path
import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from progress_studio.services.financial_forecast_workbook_service import (
    FinanceWorkbookAnalysis,
    FinancialForecastWorkbookService,
)


class FinanceFrame(ttk.Frame):
    """Desktop access to the accepted finance workbook lifecycle.

    Financial records stay in Excel. Like the application runner, background
    work returns through a queue; only the Tk thread touches widgets/dialogs.
    """

    def __init__(self, master, service=None) -> None:
        super().__init__(master, style="Surface.TFrame")
        self.service = service or FinancialForecastWorkbookService()
        self.workbook_var = tk.StringVar()
        self.opening_var = tk.StringVar()
        self.actuals_var = tk.StringVar()
        self.currency_var = tk.StringVar()
        self.capacity_var = tk.StringVar(value="100")
        self.status_var = tk.StringVar(value="Select a saved Progress Studio .xlsx workbook, then Check Workbook.")
        self.result_var = tk.StringVar()
        self._analysis: FinanceWorkbookAnalysis | None = None
        self._output_path: Path | None = None
        self._busy = False
        self._worker: threading.Thread | None = None
        self._messages: queue.Queue = queue.Queue()
        self._build_ui()
        self.workbook_var.trace_add("write", self._source_changed)

    @property
    def busy(self) -> bool:
        return self._busy

    def _build_ui(self) -> None:
        # Follow Rebuild's scrollable body so all controls remain reachable at
        # the application's minimum window size and with display scaling.
        host = ttk.Frame(self, style="Surface.TFrame")
        host.pack(fill="both", expand=True)
        host.rowconfigure(0, weight=1)
        host.columnconfigure(0, weight=1)
        canvas = tk.Canvas(host, borderwidth=0, highlightthickness=0,
                           background=self.winfo_toplevel().cget("background"))
        scrollbar = ttk.Scrollbar(host, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        panel = ttk.Frame(canvas, style="Surface.TFrame", padding=20)
        window = canvas.create_window((0, 0), window=panel, anchor="nw")
        panel.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(window, width=event.width))
        panel.columnconfigure(0, weight=1)
        ttk.Label(panel, text="Financial Forecast", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            panel, style="Muted.TLabel", wraplength=680,
            text="Create or refresh Finance Input and Cash Flow on a new workbook copy. "
                 "Save and close Excel before continuing.",
        ).grid(row=1, column=0, sticky="w", pady=(4, 12))

        source = ttk.Frame(panel, style="Card.TFrame", padding=12)
        source.grid(row=2, column=0, sticky="ew")
        source.columnconfigure(0, weight=1)
        ttk.Label(source, text="Workbook", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        self.path_entry = ttk.Entry(source, textvariable=self.workbook_var)
        self.path_entry.grid(row=1, column=0, sticky="ew", pady=6)
        self.browse_button = ttk.Button(source, text="Browse...", command=self._browse)
        self.browse_button.grid(row=1, column=1, padx=8)
        self.check_button = ttk.Button(source, text="Check Workbook", command=self._check)
        self.check_button.grid(row=1, column=2)
        ttk.Label(source, textvariable=self.status_var, style="Muted.TLabel", wraplength=680).grid(
            row=2, column=0, columnspan=3, sticky="w")

        settings = ttk.Frame(panel, style="Card.TFrame", padding=12)
        settings.grid(row=3, column=0, sticky="ew", pady=10)
        settings.columnconfigure(1, weight=1)
        self.initial_entries = []
        for row, (label, variable) in enumerate((
            ("Opening date (YYYY-MM-DD)", self.opening_var),
            ("Actuals through (YYYY-MM-DD)", self.actuals_var),
            ("Project currency", self.currency_var),
        )):
            ttk.Label(settings, text=label, style="Surface.TLabel").grid(row=row, column=0, sticky="w", pady=3)
            entry = ttk.Entry(settings, textvariable=variable, width=24, state="disabled")
            entry.grid(row=row, column=1, sticky="w", padx=12, pady=3)
            self.initial_entries.append(entry)
        ttk.Label(settings, text="Prepared rows per table (1–2000)", style="Surface.TLabel").grid(row=3, column=0, sticky="w", pady=3)
        self.capacity_entry = ttk.Spinbox(settings, from_=1, to=2000, textvariable=self.capacity_var, width=22, state="disabled")
        self.capacity_entry.grid(row=3, column=1, sticky="w", padx=12, pady=3)
        ttk.Label(
            settings, style="Muted.TLabel", wraplength=680,
            text="Dates and currency are required only for first creation. On refresh, saved settings "
                 "and inputs are retained; prepared rows can increase. Edit financial records in Excel.",
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))

        actions = ttk.Frame(panel, style="Surface.TFrame")
        actions.grid(row=4, column=0, sticky="w")
        self.generate_button = ttk.Button(actions, text="Create Finance Workbook", style="Accent.TButton", command=self._generate, state="disabled")
        self.generate_button.pack(side="left")
        self.open_button = ttk.Button(actions, text="Open Result", command=self._open_result, state="disabled")
        self.open_button.pack(side="left", padx=8)
        ttk.Label(panel, textvariable=self.result_var, style="Muted.TLabel", wraplength=680).grid(row=5, column=0, sticky="w", pady=8)
        ttk.Label(
            panel, style="Muted.TLabel", wraplength=680,
            text="First creation starts with zero opening cash and incomplete declarations. "
                 "Enter reconciled cash and remaining forecasts before relying on funding. "
                 "Credit term defaults to 30 days and retention to 5%. F9 / Save recalculates Excel formulas.",
        ).grid(row=6, column=0, sticky="w")

    def _source_changed(self, *_args) -> None:
        self._analysis = None
        self._output_path = None
        self.result_var.set("")
        self.status_var.set("Check the selected saved workbook before creating or refreshing finance.")
        self._update_controls()

    def _update_controls(self) -> None:
        state = "disabled" if self._busy else "normal"
        for widget in (self.path_entry, self.browse_button, self.check_button):
            widget.configure(state=state)
        ready = self._analysis is not None and not self._busy
        existing = self._analysis is not None and self._analysis.existing_finance
        for widget in self.initial_entries:
            widget.configure(state="normal" if ready and not existing else "disabled")
        self.capacity_entry.configure(state="normal" if ready else "disabled")
        self.generate_button.configure(
            state="normal" if ready else "disabled",
            text="Refresh Finance Workbook" if existing else "Create Finance Workbook",
        )
        self.open_button.configure(state="normal" if self._output_path and not self._busy else "disabled")

    def _browse(self) -> None:
        if self._busy:
            return
        selected = filedialog.askopenfilename(title="Select Progress Studio workbook", filetypes=[("Excel workbook", "*.xlsx")])
        if selected:
            self.workbook_var.set(selected)
            self._check()

    def _check(self) -> None:
        if self._busy:
            return
        raw = self.workbook_var.get().strip()
        if not raw:
            messagebox.showwarning("Financial Forecast", "Select a saved .xlsx workbook first.")
            return
        self._analysis = None
        self._output_path = None
        self.result_var.set("")
        self._start("check", Path(raw).expanduser().resolve())

    def _generate(self) -> None:
        if self._busy:
            return
        analysis = self._analysis
        if analysis is None:
            messagebox.showwarning("Financial Forecast", "Check the selected workbook first.")
            return
        try:
            capacity_rows = int(self.capacity_var.get())
            if not 1 <= capacity_rows <= 2000:
                raise ValueError("Prepared rows must be a whole number between 1 and 2000.")
            if analysis.existing_finance and capacity_rows < analysis.capacity_rows:
                raise ValueError("Prepared rows cannot be reduced. Keep the existing count or increase it.")
            options = {"capacity_rows": capacity_rows}
            if not analysis.existing_finance:
                try:
                    opening = date.fromisoformat(self.opening_var.get().strip())
                    actuals = date.fromisoformat(self.actuals_var.get().strip())
                except ValueError:
                    raise ValueError("Enter opening date and actuals through as YYYY-MM-DD.") from None
                if opening < date(1900, 3, 1) or actuals < opening:
                    raise ValueError("Opening date must be from 1900-03-01; actuals through must be on or after opening date.")
                currency = self.currency_var.get().strip()
                if not currency:
                    raise ValueError("Enter the project currency.")
                options.update(opening_date=opening, actuals_through=actuals, currency=currency)
        except (ValueError, tk.TclError) as exc:
            messagebox.showwarning("Financial Forecast", str(exc))
            return
        output = filedialog.asksaveasfilename(
            title="Save finance workbook as a new file", defaultextension=".xlsx",
            initialdir=str(analysis.workbook.parent), initialfile=f"{analysis.workbook.stem}_finance.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not output:
            return
        output = Path(output).expanduser().resolve()
        if output == analysis.workbook or output.exists() or output.suffix.lower() != ".xlsx":
            messagebox.showwarning("Financial Forecast", "Choose a new .xlsx file path. Existing files cannot be overwritten.")
            return
        self._output_path = None
        self._start("generate", analysis.workbook, output, options, analysis.existing_finance)

    def _start(self, operation, *args) -> None:
        self._busy = True
        self.result_var.set("Checking workbook..." if operation == "check" else "Building finance workbook...")
        self._update_controls()
        self._worker = threading.Thread(target=self._run, args=(operation, *args), daemon=True)
        self._worker.start()
        self.after(50, self._poll)

    def _run(self, operation, source, output=None, options=None, expected_existing=None) -> None:
        try:
            analysis = self.service.analyze(source)
            if operation == "generate":
                # Recheck the saved file in case Excel changed it after Check Workbook.
                if analysis.existing_finance != expected_existing:
                    raise ValueError("Finance sheets changed since checking. Check the workbook again.")
                if analysis.existing_finance and options["capacity_rows"] < analysis.capacity_rows:
                    raise ValueError("Prepared rows changed since checking. Check the workbook again.")
                output = self.service.generate(source, output, **options)
                analysis = self.service.analyze(output)
            self._messages.put((operation, analysis, output, None))
        except Exception as exc:
            self._messages.put((operation, None, None, exc))

    def _poll(self) -> None:
        try:
            operation, analysis, output, error = self._messages.get_nowait()
        except queue.Empty:
            self.after(50, self._poll)
            return
        self._busy = False
        if error is not None:
            self._analysis = None
            self._output_path = None
            self.status_var.set("Not ready — correct the issue and Check Workbook again.")
            self.result_var.set("Finance operation failed.")
            self._update_controls()
            messagebox.showerror("Financial Forecast", str(error))
            return
        if operation == "generate":
            self.workbook_var.set(str(output))
            self._output_path = output
        self._analysis = analysis
        self.capacity_var.set(str(analysis.capacity_rows))
        self.opening_var.set(analysis.opening_date.isoformat() if analysis.opening_date else "")
        self.actuals_var.set(analysis.actuals_through.isoformat() if analysis.actuals_through else "")
        self.currency_var.set(analysis.currency or "")
        self.status_var.set("Ready to refresh — saved Finance Input will be retained." if analysis.existing_finance else "Ready for first creation — enter the finance dates and currency.")
        self.result_var.set(f"Created {output.name}. Open Result, edit Finance Input, then F9 / Save." if output else "Workbook checked. No files changed.")
        self._update_controls()

    def _open_result(self) -> None:
        path = self._output_path
        if self._busy or path is None:
            return
        try:
            if not path.is_file():
                raise FileNotFoundError(f"Workbook was not found: {path}")
            if os.name == "nt":
                os.startfile(str(path))
            elif sys.platform == "darwin":
                os.spawnlp(os.P_NOWAIT, "open", "open", str(path))
            else:
                os.spawnlp(os.P_NOWAIT, "xdg-open", "xdg-open", str(path))
        except Exception as exc:
            messagebox.showerror("Financial Forecast", f"Could not open workbook:\n{exc}")
