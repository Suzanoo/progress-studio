from __future__ import annotations

from progress_studio.presentation.gui.strings import tr
from progress_studio.presentation.gui.interaction import scroll_content

import os
import queue
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

        self._events = queue.Queue()
        self._busy = False
        self._build_ui()
        self.workbook_var.trace_add("write", self._path_changed)
        self.after(60, self._poll_results)

    def _build_ui(self) -> None:
        # Scroll the complete Payment workspace, including the BOQ Payment card.
        # Bind wheel events only while the pointer is over this workspace.
        container = ttk.Frame(self, style="Surface.TFrame")
        container.pack(fill="both", expand=True)
        canvas = tk.Canvas(container, highlightthickness=0, borderwidth=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        panel = ttk.Frame(canvas, style="Surface.TFrame", padding=(8, 6))
        window_id = canvas.create_window((0, 0), window=panel, anchor="nw")
        panel.columnconfigure(0, weight=1)

        def update_region(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def update_width(event):
            canvas.itemconfigure(window_id, width=event.width)

        panel.bind("<Configure>", update_region)
        canvas.bind("<Configure>", update_width)

        scroll_content(canvas, panel)

        def section(row, title):
            card = ttk.Frame(panel, style="Card.TFrame", padding=(12, 7))
            card.grid(row=row, column=0, sticky="ew", pady=(0, 4))
            ttk.Label(card, text=title, style="Section.TLabel").pack(anchor="w")
            return card

        source = section(0, tr('pu.source'))
        path_bar = ttk.Frame(source)
        path_bar.pack(fill="x")
        self.source_entry = ttk.Entry(path_bar, textvariable=self.workbook_var)
        self.source_entry.pack(side="left", fill="x", expand=True)
        self.source_entry.bind("<Return>", lambda _e: self._check_source())
        self.check_button = ttk.Button(path_bar, text=tr("ux.check_source"), command=self._check_source)
        self.check_button.pack(side="left")
        self.browse_button = ttk.Button(path_bar, text=tr('Browse...'), command=self._browse_workbook)
        self.browse_button.pack(side="left", padx=4)
        self._status_line(source, self.workbook_status_var)

        inputs = section(1, tr('pu.input'))
        bar = ttk.Frame(inputs)
        bar.pack(anchor="w")
        ttk.Label(bar, text=tr('Payment periods')).pack(side="left")
        self.period_spinbox = ttk.Spinbox(bar, from_=1, to=120, width=5,
                                          textvariable=self.periods_var, state="disabled")
        self.period_spinbox.pack(side="left", padx=5)
        bar = ttk.Frame(inputs)
        bar.pack(anchor='w', pady=4)
        self.prepare_button = self._work_button(bar, tr('Prepare Payment Input'), self._prepare)
        self.open_button = self._work_button(bar, tr('Open Result'), self._open_result)
        self._status_line(inputs, self.period_hint_var)
        self._status_line(inputs, self.result_var)

        breakdown = section(2, tr('Payment Breakdown'))
        bar = ttk.Frame(breakdown)
        bar.pack(anchor="w")
        self.breakdown_button = self._work_button(bar, tr('Build Payment Breakdown'), self._prepare_breakdown)
        self.breakdown_open_button = self._work_button(bar, tr('Open Result'), self._open_breakdown_result)
        self._status_line(breakdown, self.breakdown_result_var)

        boq = section(3, tr('pu.boq'))
        bar = ttk.Frame(boq)
        bar.pack(anchor="w")
        ttk.Button(bar, text=tr('bp4.prepare'), command=self._prepare_boq_requirements).pack(anchor='w', pady=2)
        ttk.Label(bar, text=tr('pu.save_excel')).pack(anchor='w', padx=8)
        ttk.Button(bar, text=tr('bp2.0'), command=self._open_boq_payment).pack(anchor='w', pady=2)

        ttk.Button(bar, text=tr('bpr.word'), command=self._open_boq_word).pack(anchor='w', pady=2)

        report = section(4, tr('pr2.title'))
        ttk.Button(report, text=tr('pu.open_report'), command=self._open_plan_report,
                   style="Accent.TButton").pack(anchor="w")

    @staticmethod
    def _work_button(parent, label, command):
        button = ttk.Button(parent, text=label, command=command, state="disabled")
        button.pack(side="left", padx=(0, 4))
        return button

    @staticmethod
    def _status_line(parent, variable):
        label = ttk.Label(parent, textvariable=variable, style="Muted.TLabel")
        label.pack(fill="x")
        label.bind("<Configure>", lambda event: label.configure(wraplength=max(120, event.width - 4)))

    def _open_plan_report(self):
        from progress_studio.presentation.gui.payment_plan_dialog import PaymentPlanDialog
        dialog = getattr(self, '_payment_plan_dialog', None)
        if dialog is not None and dialog.winfo_exists():
            dialog.deiconify()
            dialog.lift()
            return
        self._payment_plan_dialog = PaymentPlanDialog(self)

    def _open_boq_word(self):
        from progress_studio.presentation.gui.boq_word_dialog import BOQWordDialog
        dialog = getattr(self, '_boq_word_dialog', None)
        if dialog is not None and dialog.winfo_exists():
            dialog.lift()
            return
        self._boq_word_dialog = BOQWordDialog(self)

    def _prepare_boq_requirements(self):
        from progress_studio.presentation.gui.boq_payment_dialog import BOQPaymentDialog
        BOQPaymentDialog(self, prepare=True)

    def _open_boq_payment(self):
        from progress_studio.presentation.gui.boq_payment_dialog import BOQPaymentDialog
        dialog = getattr(self, '_boq_payment_dialog', None)
        if dialog is not None and dialog.winfo_exists():
            dialog.lift()
            return
        self._boq_payment_dialog = BOQPaymentDialog(self)

    def _path_changed(self, *_):
        self._validated_workbook = None
        for button in (self.prepare_button, self.breakdown_button):
            button.configure(state="disabled")
        self.period_spinbox.configure(state="disabled")
        self.workbook_status_var.set(tr('ux.source_changed'))

    def _check_source(self):
        if not self._busy and self.workbook_var.get().strip():
            self._validate(Path(self.workbook_var.get().strip()))

    def _execution_source(self):
        if self._busy:
            return None
        displayed = self.workbook_var.get().strip()
        if not displayed or self._validated_workbook != Path(displayed):
            self._path_changed()
            return None
        return self._validated_workbook

    def _set_busy(self, busy):
        self._busy = busy
        for widget in (self.source_entry, self.browse_button, self.check_button):
            widget.configure(state="disabled" if busy else "normal")
        ready = self._validated_workbook is not None and not busy
        for widget in (self.prepare_button, self.breakdown_button, self.period_spinbox):
            widget.configure(state="normal" if ready else "disabled")

    def _poll_results(self):
        try:
            while True:
                callback, args = self._events.get_nowait()
                self._set_busy(False)
                callback(*args)
        except queue.Empty:
            pass
        if self.winfo_exists():
            self.after(60, self._poll_results)

    def _browse_workbook(self) -> None:
        if self._busy:
            return
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
        if self._busy:
            return
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
        except (PaymentWorkbookError, OSError, ValueError) as exc:
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
        source = self._execution_source()
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

        self._set_busy(True)
        self._worker = threading.Thread(
            target=self._prepare_worker,
            args=(source, Path(output), periods),
            daemon=True,
        )
        try:
            self._worker.start()
        except Exception as exc:
            self._set_busy(False)
            self._failed(exc)

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
            self._events.put((self._failed, (exc,)))
            return

        self._events.put((self._done, (output, stats)))

    def _done(
        self,
        output: Path,
        stats: dict[str, int],
    ) -> None:
        self._output_path = output
        self.result_var.set(
            tr('Created {value1} • {value2} periods • {value3:,} activities • {value4:,} preserved', value1=output.name, value2=stats['periods'], value3=stats['activities'], value4=stats['preserved'])
        )
        self._set_busy(False)
        self.open_button.configure(state="normal")

    def _failed(self, error: Exception) -> None:
        self._set_busy(False)
        self.result_var.set(
            tr('Payment Input preparation failed.')
        )
        messagebox.showerror(
            tr('Payment'),
            tr("error.details", error=str(error)),
        )

    def _prepare_breakdown(self) -> None:
        source = self._execution_source()
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

        self._set_busy(True)
        self._worker = threading.Thread(
            target=self._prepare_breakdown_worker,
            args=(source, Path(output)),
            daemon=True,
        )
        try:
            self._worker.start()
        except Exception as exc:
            self._set_busy(False)
            self._breakdown_failed(exc)

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
            self._events.put((self._breakdown_failed, (exc,)))
            return

        self._events.put((self._breakdown_done, (output, snapshot)))

    def _breakdown_done(
        self,
        output: Path,
        snapshot,
    ) -> None:
        self._breakdown_output_path = output
        self.breakdown_result_var.set(
            tr('Created {value1} • {value2:,} derived activities • {value3:,} eligible source activities • {value4:,} skipped', value1=output.name, value2=len(snapshot.activities), value3=snapshot.eligible_source_count, value4=len(snapshot.skipped_activity_ids))
        )
        self._set_busy(False)
        self.breakdown_open_button.configure(state="normal")

    def _breakdown_failed(
        self,
        error: Exception,
    ) -> None:
        self._set_busy(False)
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
