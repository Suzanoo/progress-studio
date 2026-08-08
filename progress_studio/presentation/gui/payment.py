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
    """Payment workflow UI before line rendering.

    Step 1: select Progress workbook and optionally create the Payment snapshot.
    Step 2: generate a lightweight editable fake Payment Requirement workbook.
    Step 3: upload/validate the edited Payment Requirement file for the renderer.
    """

    def __init__(self, master, service: PaymentService | None = None) -> None:
        super().__init__(master, style="Surface.TFrame")
        self.service = service or PaymentService()

        self.progress_var = tk.StringVar()
        self.progress_status_var = tk.StringVar(value="Select an exported Progress Studio workbook.")
        self.periods_var = tk.IntVar(value=1)
        self.period_hint_var = tk.StringVar(value="Default will be calculated from Project Start / Finish.")
        self.fake_status_var = tk.StringVar(value="No Payment Requirement workbook generated yet.")
        self.payment_input_var = tk.StringVar()
        self.payment_status_var = tk.StringVar(value="Upload the edited Payment Requirement workbook when ready.")
        self.snapshot_status_var = tk.StringVar(value="Payment snapshot has not been created yet.")

        self._validated_progress: Path | None = None
        self._snapshot_path: Path | None = None
        self._payment_input_path: Path | None = None
        self._worker: threading.Thread | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        panel = ttk.Frame(self, style="Surface.TFrame", padding=24)
        panel.pack(fill="both", expand=True)
        panel.columnconfigure(0, weight=1)

        ttk.Label(panel, text="Payment", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            panel,
            text="Prepare the Payment sheet and a lightweight Payment Requirement input. Line rendering is intentionally deferred to the next milestone.",
            style="Muted.TLabel",
            wraplength=840,
        ).grid(row=1, column=0, sticky="w", pady=(6, 16))

        self._build_progress_card(panel, row=2)
        self._build_fake_card(panel, row=3)
        self._build_upload_card(panel, row=4)

    def _build_progress_card(self, parent: ttk.Frame, row: int) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=16)
        card.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text="1", style="Section.TLabel").grid(row=0, column=0, sticky="nw", padx=(0, 10))
        ttk.Label(card, text="Progress Workbook", style="Section.TLabel").grid(row=0, column=1, sticky="w")
        ttk.Label(
            card,
            text="Validate main and create a Payment snapshot without changing the source workbook.",
            style="Muted.TLabel",
        ).grid(row=1, column=1, columnspan=2, sticky="w", pady=(3, 8))

        ttk.Entry(card, textvariable=self.progress_var).grid(row=2, column=1, sticky="ew")
        ttk.Button(card, text="Browse...", command=self._browse_progress).grid(row=2, column=2, padx=(8, 0))
        ttk.Label(card, textvariable=self.progress_status_var, style="Muted.TLabel", wraplength=740).grid(
            row=3, column=1, columnspan=2, sticky="w", pady=(8, 0)
        )

        actions = ttk.Frame(card, style="Surface.TFrame")
        actions.grid(row=4, column=1, columnspan=2, sticky="w", pady=(10, 0))
        self.snapshot_button = ttk.Button(
            actions, text="Create Payment Snapshot", style="Accent.TButton", command=self._create_snapshot, state="disabled"
        )
        self.snapshot_button.pack(side="left")
        self.open_snapshot_button = ttk.Button(actions, text="Open Snapshot", command=self._open_snapshot, state="disabled")
        self.open_snapshot_button.pack(side="left", padx=(8, 0))
        ttk.Label(card, textvariable=self.snapshot_status_var, style="Muted.TLabel", wraplength=740).grid(
            row=5, column=1, columnspan=2, sticky="w", pady=(8, 0)
        )

    def _build_fake_card(self, parent: ttk.Frame, row: int) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=16)
        card.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text="2", style="Section.TLabel").grid(row=0, column=0, sticky="nw", padx=(0, 10))
        ttk.Label(card, text="Payment Requirement", style="Section.TLabel").grid(row=0, column=1, sticky="w")
        ttk.Label(
            card,
            text="If you do not have a Payment Requirement file yet, generate a one-sheet template with Activity IDs and percentage columns.",
            style="Muted.TLabel",
            wraplength=760,
        ).grid(row=1, column=1, columnspan=2, sticky="w", pady=(3, 8))

        period_row = ttk.Frame(card, style="Surface.TFrame")
        period_row.grid(row=2, column=1, columnspan=2, sticky="w")
        ttk.Label(period_row, text="Payment periods", style="Surface.TLabel").pack(side="left")
        self.period_spinbox = ttk.Spinbox(period_row, from_=1, to=120, width=7, textvariable=self.periods_var, state="disabled")
        self.period_spinbox.pack(side="left", padx=(10, 8))
        ttk.Label(period_row, textvariable=self.period_hint_var, style="Muted.TLabel").pack(side="left")

        self.fake_button = ttk.Button(
            card,
            text="Generate Fake Payment Workbook",
            command=self._generate_fake_payment,
            state="disabled",
        )
        self.fake_button.grid(row=3, column=1, sticky="w", pady=(10, 0))
        ttk.Label(card, textvariable=self.fake_status_var, style="Muted.TLabel", wraplength=740).grid(
            row=4, column=1, columnspan=2, sticky="w", pady=(8, 0)
        )

    def _build_upload_card(self, parent: ttk.Frame, row: int) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=16)
        card.grid(row=row, column=0, sticky="ew")
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text="3", style="Section.TLabel").grid(row=0, column=0, sticky="nw", padx=(0, 10))
        ttk.Label(card, text="Upload Payment Workbook", style="Section.TLabel").grid(row=0, column=1, sticky="w")
        ttk.Label(
            card,
            text="Upload the file after editing the required percentages. Progress Studio will validate Activity IDs and prepare it for the line renderer.",
            style="Muted.TLabel",
            wraplength=760,
        ).grid(row=1, column=1, columnspan=2, sticky="w", pady=(3, 8))

        ttk.Entry(card, textvariable=self.payment_input_var).grid(row=2, column=1, sticky="ew")
        self.payment_browse_button = ttk.Button(card, text="Browse...", command=self._browse_payment_input, state="disabled")
        self.payment_browse_button.grid(row=2, column=2, padx=(8, 0))
        ttk.Label(card, textvariable=self.payment_status_var, style="Muted.TLabel", wraplength=740).grid(
            row=3, column=1, columnspan=2, sticky="w", pady=(8, 0)
        )

        self.render_ready_button = ttk.Button(card, text="Render Payment Lines", style="Accent.TButton", state="disabled")
        self.render_ready_button.grid(row=4, column=1, sticky="w", pady=(10, 0))
        ttk.Label(
            card,
            text="Renderer is intentionally disabled in this milestone. A valid Payment Workbook will show Ready for Render above.",
            style="Muted.TLabel",
        ).grid(row=5, column=1, columnspan=2, sticky="w", pady=(6, 0))

    def _browse_progress(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select Progress Studio workbook",
            filetypes=[("Excel workbook", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if not selected:
            return
        self.progress_var.set(selected)
        self._validate_progress(Path(selected))

    def _validate_progress(self, source: Path) -> None:
        self._validated_progress = None
        self.snapshot_button.configure(state="disabled")
        self.fake_button.configure(state="disabled")
        self.period_spinbox.configure(state="disabled")
        self.payment_browse_button.configure(state="disabled")
        try:
            result = self.service.validate_workbook(source)
        except PaymentWorkbookError as exc:
            self.progress_status_var.set(f"Not ready — {exc}")
            return

        self._validated_progress = source
        self.periods_var.set(result.default_payment_periods)
        if result.project_start and result.project_finish:
            self.period_hint_var.set(
                f"Default {result.default_payment_periods} from {result.project_start:%d-%b-%y} to {result.project_finish:%d-%b-%y}."
            )
        else:
            self.period_hint_var.set("Project dates were not resolved; review the period count before generating.")
        self.progress_status_var.set(
            f"Ready — main found • {result.activity_rows:,} activities • {result.max_row:,} rows × {result.max_column:,} columns"
        )
        self.snapshot_button.configure(state="normal")
        self.fake_button.configure(state="normal")
        self.period_spinbox.configure(state="normal")
        self.payment_browse_button.configure(state="normal")

    def _create_snapshot(self) -> None:
        source = self._require_progress()
        if source is None:
            return
        suffix = source.suffix.lower() if source.suffix.lower() in {".xlsx", ".xlsm"} else ".xlsx"
        output = filedialog.asksaveasfilename(
            title="Save Payment snapshot workbook",
            defaultextension=suffix,
            initialdir=str(source.parent),
            initialfile=f"{source.stem}_payment{suffix}",
            filetypes=[("Excel workbook", f"*{suffix}"), ("All files", "*.*")],
        )
        if not output:
            return
        self.snapshot_button.configure(state="disabled")
        self.snapshot_status_var.set("Creating Payment snapshot...")
        self._worker = threading.Thread(target=self._snapshot_worker, args=(source, Path(output)), daemon=True)
        self._worker.start()

    def _snapshot_worker(self, source: Path, output: Path) -> None:
        try:
            result = self.service.create_payment_snapshot(source, output)
        except Exception as exc:
            self.after(0, lambda: self._snapshot_failed(exc))
            return
        self.after(0, lambda: self._snapshot_done(result.output_workbook, result.replaced_existing_sheet))

    def _snapshot_done(self, output: Path, replaced: bool) -> None:
        self._snapshot_path = output
        suffix = " • replaced existing Payment sheet" if replaced else ""
        self.snapshot_status_var.set(f"Created {output.name} • main → Payment snapshot{suffix}")
        self.snapshot_button.configure(state="normal")
        self.open_snapshot_button.configure(state="normal")

    def _snapshot_failed(self, error: Exception) -> None:
        self.snapshot_button.configure(state="normal")
        self.snapshot_status_var.set("Payment snapshot failed.")
        messagebox.showerror("Payment", str(error))

    def _generate_fake_payment(self) -> None:
        source = self._require_progress()
        if source is None:
            return
        try:
            periods = int(self.periods_var.get())
        except (TypeError, ValueError, tk.TclError):
            messagebox.showwarning("Payment", "Payment periods must be a whole number.")
            return
        if periods < 1 or periods > 120:
            messagebox.showwarning("Payment", "Payment periods must be between 1 and 120.")
            return

        output = filedialog.asksaveasfilename(
            title="Save Payment Requirement workbook",
            defaultextension=".xlsx",
            initialdir=str(source.parent),
            initialfile=f"{source.stem}_payment_input.xlsx",
            filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")],
        )
        if not output:
            return
        self.fake_button.configure(state="disabled")
        self.fake_status_var.set("Generating lightweight Payment Requirement workbook...")
        self._worker = threading.Thread(target=self._fake_worker, args=(source, Path(output), periods), daemon=True)
        self._worker.start()

    def _fake_worker(self, source: Path, output: Path, periods: int) -> None:
        try:
            result = self.service.create_fake_payment_input(source, output, periods)
        except Exception as exc:
            self.after(0, lambda: self._fake_failed(exc))
            return
        self.after(0, lambda: self._fake_done(result.output_workbook, result.activity_rows, result.payment_periods))

    def _fake_done(self, output: Path, activities: int, periods: int) -> None:
        self._payment_input_path = output
        self.payment_input_var.set(str(output))
        self.fake_status_var.set(f"Created {output.name} • {activities:,} activities × {periods} payments")
        self.fake_button.configure(state="normal")
        self._validate_payment_input(output)

    def _fake_failed(self, error: Exception) -> None:
        self.fake_button.configure(state="normal")
        self.fake_status_var.set("Payment Requirement generation failed.")
        messagebox.showerror("Payment", str(error))

    def _browse_payment_input(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select Payment Requirement workbook",
            filetypes=[("Excel workbook", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if not selected:
            return
        self.payment_input_var.set(selected)
        self._validate_payment_input(Path(selected))

    def _validate_payment_input(self, payment_file: Path) -> None:
        progress = self._validated_progress
        try:
            if progress is None:
                result = self.service.validate_payment_input(payment_file, None)
                preparation = None
            else:
                preparation = self.service.prepare_payment_input(progress, payment_file)
                result = preparation.validation
        except PaymentWorkbookError as exc:
            self.payment_status_var.set(f"Not ready — {exc}")
            return
        self._payment_input_path = payment_file
        if result.missing_activities:
            self.payment_status_var.set(
                f"Review — {result.payment_periods} payments • {result.matched_activities:,}/{result.activity_rows:,} Activity IDs matched • {result.missing_activities:,} missing"
            )
        elif preparation is not None:
            positions = preparation.positions
            self.payment_status_var.set(
                f"Ready for Render — {result.payment_periods} payments • {result.populated_requirements:,} requirements • "
                f"{positions.resolved_count:,} positions resolved • {len(positions.issues):,} issues"
            )
        else:
            self.payment_status_var.set(
                f"Ready — {result.payment_periods} payments • {result.populated_requirements:,} requirements"
            )

    def _require_progress(self) -> Path | None:
        if self._validated_progress is not None:
            return self._validated_progress
        raw = self.progress_var.get().strip()
        if not raw:
            messagebox.showwarning("Payment", "Select a Progress Studio workbook first.")
            return None
        self._validate_progress(Path(raw))
        return self._validated_progress

    def _open_snapshot(self) -> None:
        if self._snapshot_path is not None:
            self._open_file(self._snapshot_path)

    @staticmethod
    def _open_file(path: Path) -> None:
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}"')
        except Exception as exc:
            messagebox.showerror("Open failed", str(exc))
