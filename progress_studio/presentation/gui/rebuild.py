from __future__ import annotations

from progress_studio.presentation.gui.strings import tr

import os
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from progress_studio.domain.rebuild_models import RebuildMode
from progress_studio.services.earned_value_rebuild_service import (
    EarnedValueRebuildError,
    EarnedValueRebuildService,
)
from progress_studio.services.rebuild_service import (
    RebuildContractError,
    WorkbookRebuildEngine,
)


class RebuildFrame(ttk.Frame):
    """Standalone workbook rebuild workspace.

    Progress/Payment keep their existing rebuild ownership. Earned Value is an
    additive workspace inside Rebuild and reads only the selected workbook.
    """

    def __init__(
        self,
        master,
        engine: WorkbookRebuildEngine | None = None,
        ev_service: EarnedValueRebuildService | None = None,
    ) -> None:
        super().__init__(master, style="Surface.TFrame")
        self.engine = engine or WorkbookRebuildEngine()
        self.ev_service = ev_service or EarnedValueRebuildService()

        self.workbook_var = tk.StringVar()
        self.output_mode_var = tk.StringVar(value="snapshot")
        self.mode_var = tk.StringVar(value=RebuildMode.PROGRESS.value)
        self.target_var = tk.StringVar(value="progress")

        self.analysis_var = tk.StringVar(value=tr('Select a workbook to analyze.'))
        self.detail_var = tk.StringVar(
            value=tr('main is preserved and used as the source of truth.')
        )
        self.result_var = tk.StringVar(value="")

        self.ev_source_var = tk.StringVar(value="BOQ Mapping")
        self.ev_status_var = tk.StringVar(
            value=tr('Select a workbook to check Earned Value readiness.')
        )
        self.ev_detail_var = tk.StringVar(
            value=tr('Select EV monetary source. Activity Amount uses Allocated Contract Value.')
        )

        self._validated_path: Path | None = None
        self._ev_validated_path: Path | None = None
        self._output_path: Path | None = None
        self._worker: threading.Thread | None = None
        self._busy = False
        self._build_ui()
        self.workbook_var.trace_add("write", self._path_changed)
        self._target_changed()

    def _build_ui(self) -> None:
        # Rebuild now contains several independent workspaces. Keep the page
        # usable on smaller/low-resolution windows by making only the content
        # body vertically scrollable; the surrounding application chrome stays
        # fixed.
        scroll_host = ttk.Frame(self, style="Surface.TFrame")
        scroll_host.pack(fill="both", expand=True)
        scroll_host.rowconfigure(0, weight=1)
        scroll_host.columnconfigure(0, weight=1)

        canvas = tk.Canvas(
            scroll_host,
            borderwidth=0,
            highlightthickness=0,
            background=self.winfo_toplevel().cget("background"),
        )
        scrollbar = ttk.Scrollbar(
            scroll_host,
            orient="vertical",
            command=canvas.yview,
        )
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        panel = ttk.Frame(canvas, style="Surface.TFrame", padding=24)
        panel.columnconfigure(0, weight=1)
        window_id = canvas.create_window((0, 0), window=panel, anchor="nw")

        def _sync_scroll_region(_event=None) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _sync_panel_width(event) -> None:
            canvas.itemconfigure(window_id, width=event.width)

        panel.bind("<Configure>", _sync_scroll_region)
        canvas.bind("<Configure>", _sync_panel_width)

        def _on_mousewheel(event) -> str:
            # Windows/macOS report MouseWheel; Linux/X11 uses Button-4/5.
            if getattr(event, "num", None) == 4:
                units = -3
            elif getattr(event, "num", None) == 5:
                units = 3
            else:
                delta = getattr(event, "delta", 0)
                if delta == 0:
                    return "break"
                units = -1 if delta > 0 else 1
            canvas.yview_scroll(units, "units")
            return "break"

        def _enable_mousewheel(_event=None) -> None:
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", _on_mousewheel)
            canvas.bind_all("<Button-5>", _on_mousewheel)

        def _disable_mousewheel(_event=None) -> None:
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        canvas.bind("<Enter>", _enable_mousewheel)
        canvas.bind("<Leave>", _disable_mousewheel)
        panel.bind("<Enter>", _enable_mousewheel)
        panel.bind("<Leave>", _disable_mousewheel)

        ttk.Label(panel, text=tr('Rebuild Workbook'), style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            panel,
            text=(
                tr('Rebuild generated views from the workbook itself. No Progress Studio session, BOQ file, or XML source is required.')
            ),
            style="Muted.TLabel",
            wraplength=850,
        ).grid(row=1, column=0, sticky="w", pady=(6, 16))

        self._build_file_card(panel, 2)
        self._build_scope_card(panel, 3)
        self._build_output_mode_card(panel, 4)
        self._build_ev_card(panel, 5)
        self._build_action_card(panel, 6)

    def _build_file_card(self, parent: ttk.Frame, row: int) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=16)
        card.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text="1", style="Section.TLabel").grid(
            row=0, column=0, sticky="nw", padx=(0, 10)
        )
        ttk.Label(card, text=tr('Workbook'), style="Section.TLabel").grid(
            row=0, column=1, sticky="w"
        )
        ttk.Label(
            card,
            text=tr('Select the saved Excel workbook containing your latest progress edits.'),
            style="Muted.TLabel",
            wraplength=760,
        ).grid(row=1, column=1, columnspan=2, sticky="w", pady=(3, 8))
        ttk.Entry(card, textvariable=self.workbook_var, state="readonly").grid(
            row=2, column=1, sticky="ew"
        )
        ttk.Button(card, text=tr('Browse...'), command=self._browse).grid(
            row=2, column=2, padx=(8, 0)
        )
        ttk.Label(
            card,
            textvariable=self.analysis_var,
            style="Muted.TLabel",
            wraplength=780,
        ).grid(row=3, column=1, columnspan=2, sticky="w", pady=(8, 0))

    def _build_output_mode_card(self, parent: ttk.Frame, row: int) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=16)
        self.output_options = card
        card.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text="3", style="Section.TLabel").grid(
            row=0, column=0, sticky="nw", padx=(0, 10)
        )
        ttk.Label(card, text=tr('Output Mode'), style="Section.TLabel").grid(
            row=0, column=1, sticky="w"
        )
        ttk.Radiobutton(
            card,
            text=tr('Snapshot Workbook'),
            variable=self.output_mode_var,
            value="snapshot",
            command=self._output_mode_changed,
        ).grid(row=1, column=1, sticky="w", pady=(10, 0))
        ttk.Label(
            card,
            text=tr('Refresh generated views from saved workbook edits.'),
            style="Muted.TLabel",
        ).grid(row=2, column=1, sticky="w", padx=(24, 0), pady=(2, 8))
        ttk.Radiobutton(
            card,
            text=tr('Live Workbook'),
            variable=self.output_mode_var,
            value="live",
            command=self._output_mode_changed,
        ).grid(row=3, column=1, sticky="w")
        ttk.Label(
            card,
            text=tr('Keep supported views linked to Excel inputs. Use F9 / Save to recalculate.'),
            wraplength=650,
            style="Muted.TLabel",
        ).grid(row=4, column=1, sticky="w", padx=(24, 0), pady=(2, 0))

    def _build_scope_card(self, parent: ttk.Frame, row: int) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=16)
        card.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text="2", style="Section.TLabel").grid(
            row=0, column=0, sticky="nw", padx=(0, 10)
        )
        ttk.Label(card, text=tr('Target'), style="Section.TLabel").grid(
            row=0, column=1, sticky="w"
        )

        progress = ttk.Radiobutton(
            card,
            text=tr('Progress'),
            variable=self.target_var,
            value=RebuildMode.PROGRESS.value,
            command=self._target_changed,
        )
        progress.grid(row=1, column=1, sticky="w", pady=(10, 0))
        ttk.Label(
            card,
            text=tr('Refresh progress, monthly reporting and Dashboard views.'),
            style="Muted.TLabel",
        ).grid(row=2, column=1, sticky="w", padx=(24, 0), pady=(2, 8))

        payment = ttk.Radiobutton(
            card,
            text=tr('Payment'),
            variable=self.target_var,
            value=RebuildMode.PAYMENT.value,
            command=self._target_changed,
        )
        payment.grid(row=3, column=1, sticky="w")
        ttk.Label(
            card,
            text=tr('Refresh Payment lines from main and Payment Input.'),
            style="Muted.TLabel",
        ).grid(row=4, column=1, sticky="w", padx=(24, 0), pady=(2, 0))

        ttk.Radiobutton(card, text=tr('Earned Value'), variable=self.target_var,
                        value="ev", command=self._target_changed).grid(row=5, column=1, sticky="w", pady=(8, 0))
        ttk.Label(
            card,
            textvariable=self.detail_var,
            style="Muted.TLabel",
            wraplength=780,
        ).grid(row=6, column=1, sticky="w", pady=(12, 0))

    def _build_ev_card(self, parent: ttk.Frame, row: int) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=16)
        self.ev_options = card
        card.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text="3", style="Section.TLabel").grid(
            row=0, column=0, sticky="nw", padx=(0, 10)
        )
        ttk.Label(card, text=tr('Earned Value — Live'), style="Section.TLabel").grid(
            row=0, column=1, sticky="w"
        )
        ttk.Label(
            card,
            text=(
                tr('Generate / Refresh the Earned Value view from current main progress and the selected monetary source. BAC changes require EV refresh.')
            ),
            style="Muted.TLabel",
            wraplength=760,
        ).grid(row=1, column=1, columnspan=2, sticky="w", pady=(3, 8))

        actions = ttk.Frame(card, style="Surface.TFrame")
        actions.grid(row=2, column=1, sticky="w")

        for source in ("Activity Amount", "BOQ Mapping"):
            ttk.Radiobutton(actions, text=tr(source), variable=self.ev_source_var,
                            value=source, command=self._target_changed).pack(anchor="w", pady=3)
        self.prepare_button = ttk.Button(actions, text=tr('Prepare monetary inputs'), command=self._prepare_ev_inputs)
        self.prepare_button.pack(anchor="w", pady=8)

        ttk.Label(
            card,
            textvariable=self.ev_status_var,
            style="Muted.TLabel",
            wraplength=780,
        ).grid(row=3, column=1, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Label(
            card,
            textvariable=self.ev_detail_var,
            style="Muted.TLabel",
            wraplength=780,
        ).grid(row=4, column=1, columnspan=2, sticky="w", pady=(3, 0))

    def _build_action_card(self, parent: ttk.Frame, row: int) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=16)
        card.grid(row=row, column=0, sticky="ew")
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text="4", style="Section.TLabel").grid(
            row=0, column=0, sticky="nw", padx=(0, 10)
        )
        ttk.Label(card, text=tr('Build'), style="Section.TLabel").grid(
            row=0, column=1, sticky="w"
        )
        actions = ttk.Frame(card, style="Surface.TFrame")
        actions.grid(row=1, column=1, sticky="w", pady=(10, 0))

        self.rebuild_button = ttk.Button(
            actions,
            text=tr('Rebuild'),
            style="Accent.TButton",
            state="disabled",
            command=self._run_target,
        )
        self.rebuild_button.pack(side="left")
        self.ev_generate_button = self.rebuild_button

        self.open_button = ttk.Button(
            actions,
            text=tr('Open Result'),
            state="disabled",
            command=self._open_result,
        )
        self.open_button.pack(side="left", padx=(8, 0))

        ttk.Label(
            card,
            textvariable=self.result_var,
            style="Muted.TLabel",
            wraplength=780,
        ).grid(row=2, column=1, sticky="w", pady=(8, 0))

    def _browse(self) -> None:
        selected = filedialog.askopenfilename(
            title=tr('Select workbook to rebuild'),
            filetypes=[
                ("Excel workbook", "*.xlsx *.xlsm"),
                ("All files", "*.*"),
            ],
        )
        if not selected:
            return
        path = Path(selected)
        self.workbook_var.set(selected)
        self._target_changed()

    def _path_changed(self, *_args) -> None:
        self._validated_path = self._ev_validated_path = None
        self._output_path = None
        self.rebuild_button.configure(state="disabled")
        self.open_button.configure(state="disabled")

    def _target_changed(self) -> None:
        if self._busy:
            return
        ev = self.target_var.get() == "ev"
        self.output_options.grid_remove() if ev else self.output_options.grid()
        self.ev_options.grid() if ev else self.ev_options.grid_remove()
        self.prepare_button.pack_forget()
        if ev and self.ev_source_var.get() == "Activity Amount":
            self.prepare_button.pack(anchor="w", pady=8)
        self._path_changed()
        self.result_var.set("")
        self.detail_var.set("")
        self.analysis_var.set(tr("Select a workbook to analyze."))
        self.rebuild_button.configure(text=tr("Generate / Refresh EV" if ev else "Rebuild"))
        if not ev:
            self.mode_var.set(self.target_var.get())
        current = self.workbook_var.get().strip()
        if current:
            if ev:
                self._analyze_ev(Path(current))
                self.analysis_var.set(self.ev_status_var.get())
            else:
                self._analyze(Path(current))

    def _run_target(self) -> None:
        if self._busy:
            return
        if self.target_var.get() == "ev":
            self._generate_ev()
        else:
            self._rebuild()

    def _output_mode_changed(self) -> None:
        if self._busy:
            return
        current = self.workbook_var.get().strip()
        if current:
            self._analyze(Path(current))
        else:
            self._apply_output_mode_state()

    def _mode_changed(self) -> None:
        current = self.workbook_var.get().strip()
        if current:
            self._analyze(Path(current))

    def _apply_output_mode_state(self) -> None:
        mode = RebuildMode(self.mode_var.get())
        if self.output_mode_var.get() == "live" and mode is RebuildMode.PAYMENT:
            self.result_var.set(
                tr('Live Payment uses the saved Payment Input and progress data.')
            )
        if self._validated_path is not None:
            self.rebuild_button.configure(state="normal")

    def _analyze(self, path: Path) -> None:
        self._validated_path = None
        self._output_path = None
        self.rebuild_button.configure(state="disabled")
        self.open_button.configure(state="disabled")
        self.result_var.set("")

        mode = RebuildMode(self.mode_var.get())
        try:
            analysis = self.engine.analyze(path, mode)
        except RebuildContractError as exc:
            self.analysis_var.set(tr('Not ready — {value1}', value1=exc))
            self.detail_var.set(self._mode_description(mode))
            return

        self._validated_path = path
        payment_text = (
            "Payment Input found"
            if analysis.payment_input_present
            else "Payment Input not found"
        )
        self.analysis_var.set(
            tr('Ready • main found • {value1:,} activities • {value2}', value1=analysis.activity_count, value2=payment_text)
        )

        if mode is RebuildMode.PROGRESS:
            found = len(analysis.existing_generated_sheets)
            self.detail_var.set(
                tr('Progress rebuild will replace {value1} generated sheets ({value2} currently present). main is preserved.', value1=len(analysis.contract.generated_progress), value2=found)
            )
        else:
            self.detail_var.set(
                tr('Payment rebuild replaces Payment only. Progress, monthly and dashboard sheets are preserved.')
            )
        self._apply_output_mode_state()

    def _analyze_ev(self, path: Path) -> None:
        self._ev_validated_path = None
        self.ev_generate_button.configure(state="disabled")
        try:
            analysis = self.ev_service.analyze(path, monetary_source=self.ev_source_var.get())
        except EarnedValueRebuildError as exc:
            self.ev_status_var.set(tr('Not ready — {value1}', value1=exc))
            self.ev_detail_var.set(
                tr('Check the selected monetary source and resolve the reported input issues before generating EV.')
            )
            return

        self._ev_validated_path = path
        action = "Refresh" if analysis.existing_earned_value_sheet else "Generate"
        self.ev_status_var.set(
            tr('Ready • {value1:,} activities • {value2}', value1=analysis.activity_count, value2=self.ev_source_var.get())
        )
        self.ev_detail_var.set(
            f"{action} Earned Value • BAC {analysis.project_bac:,.2f} • "
            f"Allocated Contract Value. {analysis.allocation_count:,} BOQ allocations."
            if self.ev_source_var.get() == "BOQ Mapping" else
            f"{action} Earned Value • Allocated Contract Value {analysis.project_bac:,.2f}."
        )
        self.ev_generate_button.configure(state="normal")

    @staticmethod
    def _mode_description(mode: RebuildMode) -> str:
        if mode is RebuildMode.PAYMENT:
            return "Payment requires main + embedded Payment Input."
        return "Progress requires a valid main sheet."

    def _rebuild(self) -> None:
        source = self._validated_path
        if source is None:
            messagebox.showwarning(tr('Rebuild'), tr('Select a valid workbook first.'))
            return

        suffix = source.suffix.lower()
        mode = RebuildMode(self.mode_var.get())
        live = self.output_mode_var.get() == "live"
        mode_suffix = (
            ("progress_live" if live else "progress_rebuilt")
            if mode is RebuildMode.PROGRESS
            else "payment_rebuilt"
        )

        output = filedialog.asksaveasfilename(
            title=tr('Save rebuilt workbook'),
            defaultextension=suffix,
            initialdir=str(source.parent),
            initialfile=f"{source.stem}_{mode_suffix}{suffix}",
            filetypes=[("Excel workbook", f"*{suffix}"), ("All files", "*.*")],
        )
        if not output:
            return

        self.rebuild_button.configure(state="disabled")
        self.open_button.configure(state="disabled")
        self.result_var.set(
            "Rebuilding Progress-generated sheets..."
            if mode is RebuildMode.PROGRESS
            else "Rebuilding Payment only..."
        )

        self._worker = threading.Thread(
            target=self._worker_rebuild,
            args=(source, Path(output), mode, self.output_mode_var.get()),
            daemon=True,
        )
        self._set_busy(True)
        self._worker.start()

    def _prepare_ev_inputs(self) -> None:
        from progress_studio.infrastructure.excel.ev_monetary_inputs import prepare_inputs
        source = Path(self.workbook_var.get())
        if not source.is_file():
            messagebox.showwarning(tr('Earned Value'), tr('Select a workbook first.'))
            return
        output = filedialog.asksaveasfilename(title=tr('Prepare EV monetary inputs'), defaultextension=source.suffix,
                                             initialfile=f"{source.stem}_ev_inputs{source.suffix}")
        if not output:
            return
        try:
            prepare_inputs(source, Path(output))
        except Exception as exc:
            messagebox.showerror(tr('Earned Value'), tr("error.details", error=str(exc)))
            return
        messagebox.showinfo(tr('Earned Value'), tr('Open the new workbook and fill EV Monetary Inputs with allocated Contract Values. Mark milestones and enter their dates. Save, select this workbook, and refresh EV.'))

    def _generate_ev(self) -> None:
        source = self._ev_validated_path
        if source is None:
            messagebox.showwarning(
                tr('Earned Value'),
                tr('Select an EV-ready workbook first.'),
            )
            return

        suffix = source.suffix.lower()
        output = filedialog.asksaveasfilename(
            title=tr('Save Earned Value workbook'),
            defaultextension=suffix,
            initialdir=str(source.parent),
            initialfile=f"{source.stem}_earned_value{suffix}",
            filetypes=[("Excel workbook", f"*{suffix}"), ("All files", "*.*")],
        )
        if not output:
            return

        self.ev_generate_button.configure(state="disabled")
        self.open_button.configure(state="disabled")
        self.result_var.set(tr('Generating Earned Value view...'))

        self._worker = threading.Thread(
            target=self._worker_generate_ev,
            args=(source, Path(output), self.ev_source_var.get()),
            daemon=True,
        )
        self._set_busy(True)
        self._worker.start()

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        if busy:
            self._control_states = []
            def disable(parent):
                for child in parent.winfo_children():
                    if isinstance(child, (ttk.Button, ttk.Radiobutton, ttk.Entry, ttk.Combobox)):
                        self._control_states.append((child, child.cget("state")))
                        child.configure(state="disabled")
                    disable(child)
            disable(self)
        else:
            for control, state in getattr(self, "_control_states", []):
                control.configure(state=state)

    def _worker_rebuild(
        self,
        source: Path,
        output: Path,
        mode: RebuildMode,
        output_mode: str,
    ) -> None:
        try:
            if mode is RebuildMode.PROGRESS:
                if output_mode == "live":
                    result = self.engine.rebuild_live_progress(
                        source,
                        output,
                        project_name=source.stem,
                    )
                else:
                    result = self.engine.rebuild_progress(
                        source,
                        output,
                        project_name=source.stem,
                    )
                summary = (
                    f"Created {output.name} • {result.activity_count:,} activities • "
                    f"{result.week_count} weeks • {result.monthly_periods} months"
                )
            else:
                if output_mode == "live":
                    result = self.engine.rebuild_live_payment(source, output)
                else:
                    result = self.engine.rebuild_payment(source, output)
                summary = (
                    f"Created {output.name} • {result.rendered_periods} payments • "
                    f"{result.rendered_points:,} requirement points"
                )
        except Exception as exc:
            self.after(0, lambda error=exc: self._failed(error))
            return
        self.after(0, lambda: self._done(output, summary))

    def _worker_generate_ev(self, source: Path, output: Path, monetary_source: str = "BOQ Mapping") -> None:
        try:
            result = self.ev_service.generate(source, output, monetary_source=monetary_source)
            summary = (
                f"Created {output.name} • Earned Value refreshed • "
                f"{monetary_source} • BAC {result.project_bac:,.2f}"
            )
        except Exception as exc:
            self.after(0, lambda error=exc: self._ev_failed(error))
            return
        self.after(0, lambda: self._ev_done(output, summary))

    def _done(self, output: Path, summary: str) -> None:
        self._set_busy(False)
        self._output_path = output
        self.result_var.set(summary)
        self.rebuild_button.configure(state="normal")
        self.open_button.configure(state="normal")

    def _failed(self, error: Exception) -> None:
        self._set_busy(False)
        self.result_var.set(tr('Rebuild failed.'))
        self.rebuild_button.configure(state="normal")
        messagebox.showerror(tr('Rebuild'), tr("error.details", error=str(error)))

    def _ev_done(self, output: Path, summary: str) -> None:
        self._set_busy(False)
        self._output_path = output
        self.result_var.set(summary)
        self.ev_generate_button.configure(state="normal")
        self.open_button.configure(state="normal")

    def _ev_failed(self, error: Exception) -> None:
        self._set_busy(False)
        self.result_var.set(tr('Earned Value generation failed.'))
        if self._ev_validated_path is not None:
            self.ev_generate_button.configure(state="normal")
        messagebox.showerror(tr('Earned Value'), tr("error.details", error=str(error)))

    def _open_result(self) -> None:
        if self._output_path is None:
            return
        try:
            if os.name == "nt":
                os.startfile(str(self._output_path))
            elif sys.platform == "darwin":
                os.spawnlp(os.P_NOWAIT, "open", "open", str(self._output_path))
            else:
                os.spawnlp(
                    os.P_NOWAIT,
                    "xdg-open",
                    "xdg-open",
                    str(self._output_path),
                )
        except Exception as exc:
            messagebox.showerror(tr('Rebuild'), tr('Could not open workbook:\n{value1}', value1=exc))
