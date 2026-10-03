from __future__ import annotations

from progress_studio.presentation.gui.strings import tr

import os
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from progress_studio.infrastructure.excel.payment_workbook import (
    PaymentWorkbookError,
)
from progress_studio.services.payment_service import PaymentService


class PaymentFrame(ttk.Frame):
    """Payment workspace.

    Payment Input remains persistent/user-editable input.
    Payment-Breakdown is a derived snapshot from current `main`.
    Payment line rebuild ownership remains in the standalone Rebuild workspace.
    """

    def __init__(
        self,
        master,
        service: PaymentService | None = None,
    ) -> None:
        super().__init__(master, style="Surface.TFrame")
        self.service = service or PaymentService()

        self.workbook_var = tk.StringVar()
        self.workbook_status_var = tk.StringVar(
            value=tr('Select an exported Progress Studio workbook.')
        )

        self.periods_var = tk.IntVar(value=1)
        self.period_hint_var = tk.StringVar(
            value=tr('Default is calculated from Project Start / Finish.')
        )
        self.result_var = tk.StringVar(
            value=tr('Prepare or reconcile the persistent Payment Input sheet.')
        )
        self.breakdown_result_var = tk.StringVar(
            value=tr('Build Payment-Breakdown from repeated exact Activity Names in main.')
        )

        self._validated_workbook: Path | None = None
        self._output_path: Path | None = None
        self._breakdown_output_path: Path | None = None
        self._worker: threading.Thread | None = None

        self._build_ui()

    def _build_ui(self) -> None:
        panel = ttk.Frame(
            self,
            style="Surface.TFrame",
            padding=24,
        )
        panel.pack(fill="both", expand=True)
        panel.columnconfigure(0, weight=1)

        ttk.Label(
            panel,
            text=tr('Payment'),
            style="Title.TLabel",
        ).grid(
            row=0,
            column=0,
            sticky="w",
        )
        ttk.Label(
            panel,
            text=(
                tr('Prepare Payment Input or derive Payment-Breakdown here. Use Rebuild when you want to regenerate Payment lines.')
            ),
            style="Muted.TLabel",
            wraplength=850,
        ).grid(
            row=1,
            column=0,
            sticky="w",
            pady=(6, 16),
        )

        self._build_workbook_card(panel, 2)
        self._build_prepare_card(panel, 3)
        self._build_breakdown_card(panel, 4)

    def _build_workbook_card(
        self,
        parent: ttk.Frame,
        row: int,
    ) -> None:
        card = ttk.Frame(
            parent,
            style="Card.TFrame",
            padding=16,
        )
        card.grid(
            row=row,
            column=0,
            sticky="ew",
            pady=(0, 10),
        )
        card.columnconfigure(1, weight=1)

        ttk.Label(
            card,
            text="1",
            style="Section.TLabel",
        ).grid(
            row=0,
            column=0,
            sticky="nw",
            padx=(0, 10),
        )
        ttk.Label(
            card,
            text=tr('Workbook'),
            style="Section.TLabel",
        ).grid(
            row=0,
            column=1,
            sticky="w",
        )
        ttk.Label(
            card,
            text=(
                tr('Select the workbook that will receive Payment Input or a derived Payment-Breakdown snapshot.')
            ),
            style="Muted.TLabel",
            wraplength=760,
        ).grid(
            row=1,
            column=1,
            columnspan=2,
            sticky="w",
            pady=(3, 8),
        )
        ttk.Entry(
            card,
            textvariable=self.workbook_var,
        ).grid(
            row=2,
            column=1,
            sticky="ew",
        )
        ttk.Button(
            card,
            text=tr('Browse...'),
            command=self._browse_workbook,
        ).grid(
            row=2,
            column=2,
            padx=(8, 0),
        )
        ttk.Label(
            card,
            textvariable=self.workbook_status_var,
            style="Muted.TLabel",
            wraplength=760,
        ).grid(
            row=3,
            column=1,
            columnspan=2,
            sticky="w",
            pady=(8, 0),
        )

    def _build_prepare_card(
        self,
        parent: ttk.Frame,
        row: int,
    ) -> None:
        card = ttk.Frame(
            parent,
            style="Card.TFrame",
            padding=16,
        )
        card.grid(
            row=row,
            column=0,
            sticky="ew",
            pady=(0, 10),
        )
        card.columnconfigure(1, weight=1)

        ttk.Label(
            card,
            text="2",
            style="Section.TLabel",
        ).grid(
            row=0,
            column=0,
            sticky="nw",
            padx=(0, 10),
        )
        ttk.Label(
            card,
            text=tr('Prepare Payment Input'),
            style="Section.TLabel",
        ).grid(
            row=0,
            column=1,
            sticky="w",
        )
        ttk.Label(
            card,
            text=(
                tr('Existing user percentages are preserved by Activity ID. New Activities receive suggested fake requirements. Payment Date is not an input.')
            ),
            style="Muted.TLabel",
            wraplength=780,
        ).grid(
            row=1,
            column=1,
            columnspan=2,
            sticky="w",
            pady=(3, 8),
        )

        period_row = ttk.Frame(
            card,
            style="Surface.TFrame",
        )
        period_row.grid(
            row=2,
            column=1,
            columnspan=2,
            sticky="w",
        )
        ttk.Label(
            period_row,
            text=tr('Payment periods'),
            style="Surface.TLabel",
        ).pack(side="left")

        self.period_spinbox = ttk.Spinbox(
            period_row,
            from_=1,
            to=120,
            width=7,
            textvariable=self.periods_var,
            state="disabled",
        )
        self.period_spinbox.pack(
            side="left",
            padx=(10, 8),
        )
        ttk.Label(
            period_row,
            textvariable=self.period_hint_var,
            style="Muted.TLabel",
        ).pack(side="left")

        actions = ttk.Frame(
            card,
            style="Surface.TFrame",
        )
        actions.grid(
            row=3,
            column=1,
            columnspan=2,
            sticky="w",
            pady=(12, 0),
        )

        self.prepare_button = ttk.Button(
            actions,
            text=tr('Prepare Payment Input'),
            style="Accent.TButton",
            command=self._prepare,
            state="disabled",
        )
        self.prepare_button.pack(side="left")

        self.open_button = ttk.Button(
            actions,
            text=tr('Open Result'),
            command=self._open_result,
            state="disabled",
        )
        self.open_button.pack(
            side="left",
            padx=(8, 0),
        )

        ttk.Label(
            card,
            textvariable=self.result_var,
            style="Muted.TLabel",
            wraplength=760,
        ).grid(
            row=4,
            column=1,
            columnspan=2,
            sticky="w",
            pady=(8, 0),
        )

    def _build_breakdown_card(
        self,
        parent: ttk.Frame,
        row: int,
    ) -> None:
        card = ttk.Frame(
            parent,
            style="Card.TFrame",
            padding=16,
        )
        card.grid(
            row=row,
            column=0,
            sticky="ew",
        )
        card.columnconfigure(1, weight=1)

        ttk.Label(
            card,
            text="3",
            style="Section.TLabel",
        ).grid(
            row=0,
            column=0,
            sticky="nw",
            padx=(0, 10),
        )
        ttk.Label(
            card,
            text=tr('Payment Breakdown'),
            style="Section.TLabel",
        ).grid(
            row=0,
            column=1,
            sticky="w",
        )
        ttk.Label(
            card,
            text=(
                tr('Derive repeated exact Activity Names from current main. Each source Activity keeps its own progress first; the combined row is Amount-weighted.')
            ),
            style="Muted.TLabel",
            wraplength=780,
        ).grid(
            row=1,
            column=1,
            columnspan=2,
            sticky="w",
            pady=(3, 8),
        )

        actions = ttk.Frame(
            card,
            style="Surface.TFrame",
        )
        actions.grid(
            row=2,
            column=1,
            columnspan=2,
            sticky="w",
            pady=(8, 0),
        )

        self.breakdown_button = ttk.Button(
            actions,
            text=tr('Build Payment Breakdown'),
            style="Accent.TButton",
            command=self._prepare_breakdown,
            state="disabled",
        )
        self.breakdown_button.pack(side="left")

        self.breakdown_open_button = ttk.Button(
            actions,
            text=tr('Open Result'),
            command=self._open_breakdown_result,
            state="disabled",
        )
        self.breakdown_open_button.pack(
            side="left",
            padx=(8, 0),
        )

        ttk.Label(
            card,
            textvariable=self.breakdown_result_var,
            style="Muted.TLabel",
            wraplength=760,
        ).grid(
            row=3,
            column=1,
            columnspan=2,
            sticky="w",
            pady=(8, 0),
        )

    def _browse_workbook(self) -> None:
        selected = filedialog.askopenfilename(
            title=tr('Select Progress Studio workbook'),
            filetypes=[
                ("Excel workbook", "*.xlsx *.xlsm"),
                ("All files", "*.*"),
            ],
        )
        if not selected:
            return

        self.workbook_var.set(selected)
        self._validate(Path(selected))

    def _validate(self, source: Path) -> None:
        self._validated_workbook = None
        self._output_path = None
        self._breakdown_output_path = None

        self.prepare_button.configure(state="disabled")
        self.open_button.configure(state="disabled")
        self.breakdown_button.configure(state="disabled")
        self.breakdown_open_button.configure(state="disabled")
        self.period_spinbox.configure(state="disabled")

        try:
            result = self.service.validate_workbook(source)
        except PaymentWorkbookError as exc:
            self.workbook_status_var.set(tr('Not ready — {value1}', value1=exc))
            return

        existing_periods = None
        existing_requirements = 0
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
                tr('Existing Payment Input • {value1} periods • {value2:,} requirements.', value1=existing_periods, value2=existing_requirements)
            )
        elif result.project_start and result.project_finish:
            self.period_hint_var.set(
                tr('Default {value1} periods from {value2:%d-%b-%y} to {value3:%d-%b-%y}.', value1=result.default_payment_periods, value2=result.project_start, value3=result.project_finish)
            )
        else:
            self.period_hint_var.set(
                tr('Review the period count before preparing.')
            )

        self.workbook_status_var.set(
            tr('Ready • main found • {value1:,} activities', value1=result.activity_rows)
        )
        self.period_spinbox.configure(state="normal")
        self.prepare_button.configure(state="normal")
        self.breakdown_button.configure(state="normal")

    def _prepare(self) -> None:
        source = self._validated_workbook
        if source is None:
            messagebox.showwarning(
                tr('Payment'),
                tr('Select a valid workbook first.'),
            )
            return

        try:
            periods = int(self.periods_var.get())
        except (TypeError, ValueError, tk.TclError):
            messagebox.showwarning(
                tr('Payment'),
                tr('Payment periods must be a whole number.'),
            )
            return

        if periods < 1 or periods > 120:
            messagebox.showwarning(
                tr('Payment'),
                tr('Payment periods must be between 1 and 120.'),
            )
            return

        suffix = source.suffix.lower()
        output = filedialog.asksaveasfilename(
            title=tr('Save workbook with Payment Input'),
            defaultextension=suffix,
            initialdir=str(source.parent),
            initialfile=f"{source.stem}_payment_input{suffix}",
            filetypes=[
                ("Excel workbook", f"*{suffix}"),
                ("All files", "*.*"),
            ],
        )
        if not output:
            return

        self.prepare_button.configure(state="disabled")
        self.open_button.configure(state="disabled")
        self.result_var.set(tr('Preparing Payment Input...'))

        self._worker = threading.Thread(
            target=self._prepare_worker,
            args=(source, Path(output), periods),
            daemon=True,
        )
        self._worker.start()

    def _prepare_worker(
        self,
        source: Path,
        output: Path,
        periods: int,
    ) -> None:
        try:
            stats = self.service.prepare_embedded_payment_input(
                source,
                output,
                periods,
            )
        except Exception as exc:
            self.after(
                0,
                lambda: self._failed(exc),
            )
            return

        self.after(
            0,
            lambda: self._done(output, stats),
        )

    def _done(
        self,
        output: Path,
        stats: dict[str, int],
    ) -> None:
        self._output_path = output
        self.result_var.set(
            tr('Created {value1} • {value2} periods • {value3:,} activities • {value4:,} preserved', value1=output.name, value2=stats['periods'], value3=stats['activities'], value4=stats['preserved'])
        )
        self.prepare_button.configure(state="normal")
        self.open_button.configure(state="normal")

    def _failed(self, error: Exception) -> None:
        self.prepare_button.configure(state="normal")
        self.result_var.set(
            tr('Payment Input preparation failed.')
        )
        messagebox.showerror(
            tr('Payment'),
            tr("error.details", error=str(error)),
        )

    def _prepare_breakdown(self) -> None:
        source = self._validated_workbook
        if source is None:
            messagebox.showwarning(
                tr('Payment'),
                tr('Select a valid workbook first.'),
            )
            return

        suffix = source.suffix.lower()
        output = filedialog.asksaveasfilename(
            title=tr('Save workbook with Payment Breakdown'),
            defaultextension=suffix,
            initialdir=str(source.parent),
            initialfile=f"{source.stem}_payment_breakdown{suffix}",
            filetypes=[
                ("Excel workbook", f"*{suffix}"),
                ("All files", "*.*"),
            ],
        )
        if not output:
            return

        self.breakdown_button.configure(state="disabled")
        self.breakdown_open_button.configure(state="disabled")
        self.breakdown_result_var.set(
            tr('Building Payment Breakdown...')
        )

        self._worker = threading.Thread(
            target=self._prepare_breakdown_worker,
            args=(source, Path(output)),
            daemon=True,
        )
        self._worker.start()

    def _prepare_breakdown_worker(
        self,
        source: Path,
        output: Path,
    ) -> None:
        try:
            snapshot = self.service.prepare_payment_breakdown(
                source,
                output,
            )
        except Exception as exc:
            self.after(
                0,
                lambda: self._breakdown_failed(exc),
            )
            return

        self.after(
            0,
            lambda: self._breakdown_done(
                output,
                snapshot,
            ),
        )

    def _breakdown_done(
        self,
        output: Path,
        snapshot,
    ) -> None:
        self._breakdown_output_path = output
        self.breakdown_result_var.set(
            tr('Created {value1} • {value2:,} derived activities • {value3:,} eligible source activities • {value4:,} skipped', value1=output.name, value2=len(snapshot.activities), value3=snapshot.eligible_source_count, value4=len(snapshot.skipped_activity_ids))
        )
        self.breakdown_button.configure(state="normal")
        self.breakdown_open_button.configure(state="normal")

    def _breakdown_failed(
        self,
        error: Exception,
    ) -> None:
        self.breakdown_button.configure(state="normal")
        self.breakdown_result_var.set(
            tr('Payment Breakdown build failed.')
        )
        messagebox.showerror(
            tr('Payment'),
            tr("error.details", error=str(error)),
        )

    def _open_result(self) -> None:
        self._open_path(self._output_path)

    def _open_breakdown_result(self) -> None:
        self._open_path(self._breakdown_output_path)

    def _open_path(
        self,
        path: Path | None,
    ) -> None:
        if path is None:
            return

        try:
            if os.name == "nt":
                os.startfile(str(path))
            elif sys.platform == "darwin":
                os.spawnlp(
                    os.P_NOWAIT,
                    "open",
                    "open",
                    str(path),
                )
            else:
                os.spawnlp(
                    os.P_NOWAIT,
                    "xdg-open",
                    "xdg-open",
                    str(path),
                )
        except Exception as exc:
            messagebox.showerror(
                tr('Payment'),
                tr('Could not open workbook:\n{value1}', value1=exc),
            )
