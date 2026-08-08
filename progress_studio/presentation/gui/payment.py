from __future__ import annotations

import os
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from progress_studio.infrastructure.excel.payment_workbook import PaymentWorkbookError
from progress_studio.presentation.gui.theme import PALETTE
from progress_studio.services.payment_service import PaymentService


class PaymentFrame(ttk.Frame):
    """MS-PAY1 workspace: select workbook, validate main, create Payment snapshot."""

    def __init__(self, master, service: PaymentService | None = None) -> None:
        super().__init__(master, style="Surface.TFrame")
        self.service = service or PaymentService()
        self.source_var = tk.StringVar()
        self.validation_var = tk.StringVar(value="Select an exported Progress Studio workbook.")
        self.output_var = tk.StringVar(value="Payment sheet has not been generated yet.")
        self._validated_source: Path | None = None
        self._output_path: Path | None = None
        self._worker: threading.Thread | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        panel = ttk.Frame(self, style="Surface.TFrame", padding=28)
        panel.pack(fill="both", expand=True)

        ttk.Label(panel, text="Payment", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            panel,
            text="Create a lightweight Payment worksheet by snapshotting the exported main sheet. Payment lines will be added in the next milestone.",
            style="Muted.TLabel",
            wraplength=780,
        ).pack(anchor="w", pady=(8, 22))

        card = ttk.Frame(panel, style="Card.TFrame", padding=18)
        card.pack(fill="x", anchor="n")
        card.columnconfigure(0, weight=1)

        ttk.Label(card, text="Progress Workbook", style="Section.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Entry(card, textvariable=self.source_var).grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(card, text="Browse...", command=self._browse).grid(row=1, column=1, padx=(8, 0), pady=(8, 0))

        self.validation_label = ttk.Label(card, textvariable=self.validation_var, style="Muted.TLabel", wraplength=740)
        self.validation_label.grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 0))

        actions = ttk.Frame(card, style="Surface.TFrame")
        actions.grid(row=3, column=0, columnspan=2, sticky="w", pady=(18, 0))
        self.generate_button = ttk.Button(actions, text="Generate Payment Sheet", style="Accent.TButton", command=self._generate, state="disabled")
        self.generate_button.pack(side="left")
        self.open_button = ttk.Button(actions, text="Open Output", command=self._open_output, state="disabled")
        self.open_button.pack(side="left", padx=(8, 0))

        ttk.Label(card, textvariable=self.output_var, style="Muted.TLabel", wraplength=740).grid(row=4, column=0, columnspan=2, sticky="w", pady=(12, 0))

        note = ttk.Frame(panel, style="Surface.TFrame", padding=(2, 18, 2, 0))
        note.pack(fill="x")
        ttk.Label(note, text="MS-PAY1 scope", style="Section.TLabel").pack(anchor="w")
        ttk.Label(
            note,
            text="Input: one exported workbook  •  Source: main  •  Output: Payment snapshot  •  No Payment Matrix  •  No Payment Line yet",
            style="Muted.TLabel",
            wraplength=820,
        ).pack(anchor="w", pady=(6, 0))

    def _browse(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select Progress Studio workbook",
            filetypes=[("Excel workbook", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if not selected:
            return
        self.source_var.set(selected)
        self._validate(Path(selected))

    def _validate(self, source: Path) -> None:
        self._validated_source = None
        self.generate_button.configure(state="disabled")
        try:
            result = self.service.validate_workbook(source)
        except PaymentWorkbookError as exc:
            self.validation_var.set(f"Not ready — {exc}")
            return
        self._validated_source = source
        self.validation_var.set(
            f"Ready — main found • {result.activity_rows:,} activity rows • {result.max_row:,} rows × {result.max_column:,} columns"
        )
        self.generate_button.configure(state="normal")

    def _generate(self) -> None:
        source = self._validated_source
        if source is None:
            raw = self.source_var.get().strip()
            if not raw:
                messagebox.showwarning("Payment", "Select a Progress Studio workbook first.")
                return
            self._validate(Path(raw))
            source = self._validated_source
            if source is None:
                return

        suffix = source.suffix.lower() if source.suffix.lower() in {".xlsx", ".xlsm"} else ".xlsx"
        default_name = f"{source.stem}_payment{suffix}"
        output = filedialog.asksaveasfilename(
            title="Save Payment workbook",
            defaultextension=suffix,
            initialdir=str(source.parent),
            initialfile=default_name,
            filetypes=[("Excel workbook", f"*{suffix}"), ("All files", "*.*")],
        )
        if not output:
            return

        self.generate_button.configure(state="disabled")
        self.output_var.set("Generating Payment snapshot...")
        self._worker = threading.Thread(target=self._generate_worker, args=(source, Path(output)), daemon=True)
        self._worker.start()

    def _generate_worker(self, source: Path, output: Path) -> None:
        try:
            result = self.service.create_payment_snapshot(source, output)
        except Exception as exc:
            self.after(0, lambda: self._generation_failed(exc))
            return
        self.after(0, lambda: self._generation_done(result))

    def _generation_done(self, result) -> None:
        self._output_path = result.output_workbook
        replaced = " • replaced existing Payment sheet" if result.replaced_existing_sheet else ""
        self.output_var.set(f"Created: {result.output_workbook.name} • Payment snapshot from main{replaced}")
        self.generate_button.configure(state="normal")
        self.open_button.configure(state="normal")
        messagebox.showinfo("Payment", f"Payment sheet created:\n{result.output_workbook}")

    def _generation_failed(self, error: Exception) -> None:
        self.generate_button.configure(state="normal")
        self.output_var.set("Payment snapshot failed.")
        messagebox.showerror("Payment", str(error))

    def _open_output(self) -> None:
        if self._output_path is None:
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(self._output_path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{self._output_path}"')
            else:
                os.system(f'xdg-open "{self._output_path}"')
        except Exception as exc:
            messagebox.showerror("Open failed", str(exc))
