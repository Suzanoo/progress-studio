from __future__ import annotations

import os
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from progress_studio.infrastructure.excel.payment_workbook import PaymentWorkbookError
from progress_studio.services.payment_service import PaymentService


class PaymentFrame(ttk.Frame):
    """MS-PAY7 one-workbook Payment workflow.

    The selected workbook is the only user payload:
      - main: persistent user-edited progress
      - Payment Input: persistent/reconciled requirements
      - progress_table: generated snapshot, rebuilt and hidden
      - Payment: generated snapshot, replaced on every rebuild
    """

    def __init__(self, master, service: PaymentService | None = None) -> None:
        super().__init__(master, style="Surface.TFrame")
        self.service = service or PaymentService()

        self.workbook_var = tk.StringVar()
        self.workbook_status_var = tk.StringVar(value="Select an exported Progress Studio workbook.")
        self.periods_var = tk.IntVar(value=1)
        self.period_hint_var = tk.StringVar(value="Default is calculated from Project Start / Finish.")
        self.rebuild_status_var = tk.StringVar(
            value="Payment Input and Payment will live inside the same workbook."
        )

        self._validated_workbook: Path | None = None
        self._output_path: Path | None = None
        self._worker: threading.Thread | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        panel = ttk.Frame(self, style="Surface.TFrame", padding=24)
        panel.pack(fill="both", expand=True)
        panel.columnconfigure(0, weight=1)

        ttk.Label(panel, text="Payment", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            panel,
            text=(
                "One-workbook workflow: edit main and Payment Input, then rebuild "
                "the generated progress_table and Payment snapshots."
            ),
            style="Muted.TLabel",
            wraplength=860,
        ).grid(row=1, column=0, sticky="w", pady=(6, 16))

        self._build_workbook_card(panel, 2)
        self._build_rebuild_card(panel, 3)

    def _build_workbook_card(self, parent: ttk.Frame, row: int) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=16)
        card.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text="1", style="Section.TLabel").grid(row=0, column=0, sticky="nw", padx=(0, 10))
        ttk.Label(card, text="Progress / Payment Workbook", style="Section.TLabel").grid(row=0, column=1, sticky="w")
        ttk.Label(
            card,
            text="Select the workbook you want to prepare or rebuild. No separate Payment file is required.",
            style="Muted.TLabel",
            wraplength=760,
        ).grid(row=1, column=1, columnspan=2, sticky="w", pady=(3, 8))

        ttk.Entry(card, textvariable=self.workbook_var).grid(row=2, column=1, sticky="ew")
        ttk.Button(card, text="Browse...", command=self._browse_workbook).grid(row=2, column=2, padx=(8, 0))
        ttk.Label(card, textvariable=self.workbook_status_var, style="Muted.TLabel", wraplength=760).grid(
            row=3, column=1, columnspan=2, sticky="w", pady=(8, 0)
        )

    def _build_rebuild_card(self, parent: ttk.Frame, row: int) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=16)
        card.grid(row=row, column=0, sticky="ew")
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text="2", style="Section.TLabel").grid(row=0, column=0, sticky="nw", padx=(0, 10))
        ttk.Label(card, text="Prepare / Rebuild Payment", style="Section.TLabel").grid(row=0, column=1, sticky="w")
        ttk.Label(
            card,
            text=(
                "Payment Input is reconciled by Activity ID. Existing percentages are preserved; "
                "new Activities get suggested values. Payment Date is no longer an input."
            ),
            style="Muted.TLabel",
            wraplength=780,
        ).grid(row=1, column=1, columnspan=2, sticky="w", pady=(3, 8))

        period_row = ttk.Frame(card, style="Surface.TFrame")
        period_row.grid(row=2, column=1, columnspan=2, sticky="w")
        ttk.Label(period_row, text="Payment periods", style="Surface.TLabel").pack(side="left")
        self.period_spinbox = ttk.Spinbox(
            period_row,
            from_=1,
            to=120,
            width=7,
            textvariable=self.periods_var,
            state="disabled",
        )
        self.period_spinbox.pack(side="left", padx=(10, 8))
        ttk.Label(period_row, textvariable=self.period_hint_var, style="Muted.TLabel").pack(side="left")

        actions = ttk.Frame(card, style="Surface.TFrame")
        actions.grid(row=3, column=1, columnspan=2, sticky="w", pady=(12, 0))
        self.rebuild_button = ttk.Button(
            actions,
            text="Prepare / Rebuild Payment Workbook",
            style="Accent.TButton",
            command=self._rebuild,
            state="disabled",
        )
        self.rebuild_button.pack(side="left")
        self.open_button = ttk.Button(
            actions,
            text="Open Result",
            command=self._open_result,
            state="disabled",
        )
        self.open_button.pack(side="left", padx=(8, 0))

        ttk.Label(card, textvariable=self.rebuild_status_var, style="Muted.TLabel", wraplength=760).grid(
            row=4, column=1, columnspan=2, sticky="w", pady=(8, 0)
        )

    def _browse_workbook(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select Progress Studio workbook",
            filetypes=[("Excel workbook", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if not selected:
            return
        self.workbook_var.set(selected)
        self._validate(Path(selected))

    def _validate(self, source: Path) -> None:
        self._validated_workbook = None
        self._output_path = None
        self.rebuild_button.configure(state="disabled")
        self.open_button.configure(state="disabled")
        self.period_spinbox.configure(state="disabled")
        try:
            result = self.service.validate_workbook(source)
        except PaymentWorkbookError as exc:
            self.workbook_status_var.set(f"Not ready — {exc}")
            return

        existing_periods = None
        existing_requirements = None
        try:
            payment = self.service.read_payment_requirements(source)
            existing_periods = len(payment.periods)
            existing_requirements = payment.populated_requirements
        except PaymentWorkbookError:
            pass

        self._validated_workbook = source
        periods = existing_periods or result.default_payment_periods
        self.periods_var.set(periods)
        if existing_periods is not None:
            self.period_hint_var.set(
                f"Embedded Payment Input found • {existing_periods} periods • "
                f"{existing_requirements or 0:,} requirements."
            )
        elif result.project_start and result.project_finish:
            self.period_hint_var.set(
                f"New Payment Input • default {result.default_payment_periods} "
                f"from {result.project_start:%d-%b-%y} to {result.project_finish:%d-%b-%y}."
            )
        else:
            self.period_hint_var.set("Review period count before preparing Payment.")

        self.workbook_status_var.set(
            f"Ready — main found • {result.activity_rows:,} activities • "
            f"{result.max_row:,} rows × {result.max_column:,} columns"
        )
        self.period_spinbox.configure(state="normal")
        self.rebuild_button.configure(state="normal")

    def _rebuild(self) -> None:
        source = self._validated_workbook
        if source is None:
            messagebox.showwarning("Payment", "Select a valid Progress Studio workbook first.")
            return
        try:
            periods = int(self.periods_var.get())
        except (TypeError, ValueError, tk.TclError):
            messagebox.showwarning("Payment", "Payment periods must be a whole number.")
            return
        if periods < 1 or periods > 120:
            messagebox.showwarning("Payment", "Payment periods must be between 1 and 120.")
            return

        suffix = source.suffix.lower() if source.suffix.lower() in {".xlsx", ".xlsm"} else ".xlsx"
        output = filedialog.asksaveasfilename(
            title="Save rebuilt Payment workbook",
            defaultextension=suffix,
            initialdir=str(source.parent),
            initialfile=f"{source.stem}_payment_rebuilt{suffix}",
            filetypes=[("Excel workbook", f"*{suffix}"), ("All files", "*.*")],
        )
        if not output:
            return

        self.rebuild_button.configure(state="disabled")
        self.open_button.configure(state="disabled")
        self.rebuild_status_var.set(
            "Rebuilding progress_table snapshot, reconciling Payment Input, and replacing Payment..."
        )
        self._worker = threading.Thread(
            target=self._rebuild_worker,
            args=(source, Path(output), periods),
            daemon=True,
        )
        self._worker.start()

    def _rebuild_worker(self, source: Path, output: Path, periods: int) -> None:
        try:
            result = self.service.rebuild_embedded_workbook(source, output, periods)
        except Exception as exc:
            self.after(0, lambda: self._rebuild_failed(exc))
            return
        rendered_periods = result.rendered_periods if result is not None else 0
        rendered_points = result.rendered_points if result is not None else 0
        self.after(
            0,
            lambda: self._rebuild_done(output, rendered_periods, rendered_points),
        )

    def _rebuild_done(self, output: Path, periods: int, points: int) -> None:
        self._output_path = output
        self.rebuild_status_var.set(
            f"Created {output.name} • {periods} Payment backbones • {points:,} requirement points"
        )
        self.rebuild_button.configure(state="normal")
        self.open_button.configure(state="normal")

    def _rebuild_failed(self, error: Exception) -> None:
        self.rebuild_button.configure(state="normal")
        self.rebuild_status_var.set("Payment rebuild failed.")
        messagebox.showerror("Payment", str(error))

    def _open_result(self) -> None:
        if self._output_path is not None:
            self._open_file(self._output_path)

    @staticmethod
    def _open_file(path: Path) -> None:
        try:
            if os.name == "nt":
                os.startfile(str(path))
            elif sys.platform == "darwin":
                os.spawnlp(os.P_NOWAIT, "open", "open", str(path))
            else:
                os.spawnlp(os.P_NOWAIT, "xdg-open", "xdg-open", str(path))
        except Exception as exc:
            messagebox.showerror("Payment", f"Could not open workbook:\n{exc}")
