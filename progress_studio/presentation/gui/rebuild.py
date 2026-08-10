from __future__ import annotations

import os
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from progress_studio.domain.rebuild_models import RebuildMode
from progress_studio.services.rebuild_service import (
    RebuildContractError,
    WorkbookRebuildEngine,
)


class RebuildFrame(ttk.Frame):
    """Standalone workbook rebuild workspace.

    Input is always one workbook. Progress mode trusts ``main``; Payment mode
    trusts ``main`` + ``Payment Input``. No project session, BOQ, XML, or mapping
    tree is requested by this UI.
    """

    def __init__(
        self,
        master,
        engine: WorkbookRebuildEngine | None = None,
    ) -> None:
        super().__init__(master, style="Surface.TFrame")
        self.engine = engine or WorkbookRebuildEngine()

        self.workbook_var = tk.StringVar()
        self.output_mode_var = tk.StringVar(value="snapshot")
        self.mode_var = tk.StringVar(value=RebuildMode.PROGRESS.value)
        self.analysis_var = tk.StringVar(value="Select a workbook to analyze.")
        self.detail_var = tk.StringVar(
            value="main is preserved and used as the source of truth."
        )
        self.result_var = tk.StringVar(value="")
        self._validated_path: Path | None = None
        self._output_path: Path | None = None
        self._worker: threading.Thread | None = None

        self._build_ui()

    def _build_ui(self) -> None:
        panel = ttk.Frame(self, style="Surface.TFrame", padding=24)
        panel.pack(fill="both", expand=True)
        panel.columnconfigure(0, weight=1)

        ttk.Label(panel, text="Rebuild Workbook", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            panel,
            text=(
                "Rebuild generated views from the workbook itself. "
                "No Progress Studio session, BOQ file, or XML source is required."
            ),
            style="Muted.TLabel",
            wraplength=850,
        ).grid(row=1, column=0, sticky="w", pady=(6, 16))

        self._build_file_card(panel, 2)
        self._build_output_mode_card(panel, 3)
        self._build_scope_card(panel, 4)
        self._build_action_card(panel, 5)

    def _build_file_card(self, parent: ttk.Frame, row: int) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=16)
        card.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text="1", style="Section.TLabel").grid(
            row=0, column=0, sticky="nw", padx=(0, 10)
        )
        ttk.Label(card, text="Workbook", style="Section.TLabel").grid(
            row=0, column=1, sticky="w"
        )
        ttk.Label(
            card,
            text="Choose the workbook the user has actually edited. main is authoritative.",
            style="Muted.TLabel",
            wraplength=760,
        ).grid(row=1, column=1, columnspan=2, sticky="w", pady=(3, 8))

        ttk.Entry(card, textvariable=self.workbook_var).grid(
            row=2, column=1, sticky="ew"
        )
        ttk.Button(card, text="Browse...", command=self._browse).grid(
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
        card.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text="2", style="Section.TLabel").grid(
            row=0, column=0, sticky="nw", padx=(0, 10)
        )
        ttk.Label(card, text="Output Mode", style="Section.TLabel").grid(
            row=0, column=1, sticky="w"
        )

        ttk.Radiobutton(
            card,
            text="Snapshot Workbook",
            variable=self.output_mode_var,
            value="snapshot",
            command=self._output_mode_changed,
        ).grid(row=1, column=1, sticky="w", pady=(10, 0))
        ttk.Label(
            card,
            text="Current production engine • fast, lightweight generated views.",
            style="Muted.TLabel",
        ).grid(row=2, column=1, sticky="w", padx=(24, 0), pady=(2, 8))

        ttk.Radiobutton(
            card,
            text="Live Workbook",
            variable=self.output_mode_var,
            value="live",
            command=self._output_mode_changed,
        ).grid(row=3, column=1, sticky="w")
        ttk.Label(
            card,
            text="LW-9 active • Live Progress + Live Payment • one-pass workbook writers.",
            style="Muted.TLabel",
        ).grid(row=4, column=1, sticky="w", padx=(24, 0), pady=(2, 0))

    def _build_scope_card(self, parent: ttk.Frame, row: int) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=16)
        card.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text="3", style="Section.TLabel").grid(
            row=0, column=0, sticky="nw", padx=(0, 10)
        )
        ttk.Label(card, text="Rebuild Scope", style="Section.TLabel").grid(
            row=0, column=1, sticky="w"
        )

        progress = ttk.Radiobutton(
            card,
            text="Progress",
            variable=self.mode_var,
            value=RebuildMode.PROGRESS.value,
            command=self._mode_changed,
        )
        progress.grid(row=1, column=1, sticky="w", pady=(10, 0))
        ttk.Label(
            card,
            text="Snapshot: monthly/progress/table/dashboard • Live: monthly cache + dashboard (no progress_table)",
            style="Muted.TLabel",
        ).grid(row=2, column=1, sticky="w", padx=(24, 0), pady=(2, 8))

        payment = ttk.Radiobutton(
            card,
            text="Payment",
            variable=self.mode_var,
            value=RebuildMode.PAYMENT.value,
            command=self._mode_changed,
        )
        payment.grid(row=3, column=1, sticky="w")
        ttk.Label(
            card,
            text="Payment only • source: main + Payment Input",
            style="Muted.TLabel",
        ).grid(row=4, column=1, sticky="w", padx=(24, 0), pady=(2, 0))

        ttk.Label(
            card,
            textvariable=self.detail_var,
            style="Muted.TLabel",
            wraplength=780,
        ).grid(row=5, column=1, sticky="w", pady=(12, 0))

    def _build_action_card(self, parent: ttk.Frame, row: int) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=16)
        card.grid(row=row, column=0, sticky="ew")
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text="4", style="Section.TLabel").grid(
            row=0, column=0, sticky="nw", padx=(0, 10)
        )
        ttk.Label(card, text="Build", style="Section.TLabel").grid(
            row=0, column=1, sticky="w"
        )

        actions = ttk.Frame(card, style="Surface.TFrame")
        actions.grid(row=1, column=1, sticky="w", pady=(10, 0))
        self.rebuild_button = ttk.Button(
            actions,
            text="Rebuild",
            style="Accent.TButton",
            state="disabled",
            command=self._rebuild,
        )
        self.rebuild_button.pack(side="left")
        self.open_button = ttk.Button(
            actions,
            text="Open Result",
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
            title="Select workbook to rebuild",
            filetypes=[
                ("Excel workbook", "*.xlsx *.xlsm"),
                ("All files", "*.*"),
            ],
        )
        if not selected:
            return
        self.workbook_var.set(selected)
        self._analyze(Path(selected))

    def _output_mode_changed(self) -> None:
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
                "Live Payment active in LW-9 • sparse Payment Input + MainDataset • one-pass writer."
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
            self.analysis_var.set(f"Not ready — {exc}")
            self.detail_var.set(self._mode_description(mode))
            return

        self._validated_path = path
        payment_text = (
            "Payment Input found"
            if analysis.payment_input_present
            else "Payment Input not found"
        )
        self.analysis_var.set(
            f"Ready • main found • {analysis.activity_count:,} activities • {payment_text}"
        )
        if mode is RebuildMode.PROGRESS:
            found = len(analysis.existing_generated_sheets)
            self.detail_var.set(
                f"Progress rebuild will replace {len(analysis.contract.generated_progress)} "
                f"generated sheets ({found} currently present). main is preserved."
            )
        else:
            self.detail_var.set(
                "Payment rebuild replaces Payment only. "
                "Progress, monthly and dashboard sheets are preserved."
            )
        self._apply_output_mode_state()

    @staticmethod
    def _mode_description(mode: RebuildMode) -> str:
        if mode is RebuildMode.PAYMENT:
            return "Payment requires main + embedded Payment Input."
        return "Progress requires a valid main sheet."

    def _rebuild(self) -> None:
        source = self._validated_path
        if source is None:
            messagebox.showwarning("Rebuild", "Select a valid workbook first.")
            return

        suffix = source.suffix.lower()
        mode = RebuildMode(self.mode_var.get())
        live = self.output_mode_var.get() == "live"
        mode_suffix = ("progress_live" if live else "progress_rebuilt") if mode is RebuildMode.PROGRESS else "payment_rebuilt"
        output = filedialog.asksaveasfilename(
            title="Save rebuilt workbook",
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
        self._worker.start()

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
            self.after(0, lambda: self._failed(exc))
            return
        self.after(0, lambda: self._done(output, summary))

    def _done(self, output: Path, summary: str) -> None:
        self._output_path = output
        self.result_var.set(summary)
        self.rebuild_button.configure(state="normal")
        self.open_button.configure(state="normal")

    def _failed(self, error: Exception) -> None:
        self.result_var.set("Rebuild failed.")
        self.rebuild_button.configure(state="normal")
        messagebox.showerror("Rebuild", str(error))

    def _open_result(self) -> None:
        if self._output_path is None:
            return
        try:
            if os.name == "nt":
                os.startfile(str(self._output_path))
            elif sys.platform == "darwin":
                os.spawnlp(os.P_NOWAIT, "open", "open", str(self._output_path))
            else:
                os.spawnlp(os.P_NOWAIT, "xdg-open", "xdg-open", str(self._output_path))
        except Exception as exc:
            messagebox.showerror("Rebuild", f"Could not open workbook:\n{exc}")
