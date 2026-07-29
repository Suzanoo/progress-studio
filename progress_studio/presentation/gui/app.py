from __future__ import annotations

import os
import queue
import sys
import threading
import tkinter as tk
from contextlib import redirect_stderr, redirect_stdout
from io import StringIO
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from progress_studio.app.desktop import DesktopRunOptions, DesktopRunner
from progress_studio.app.pipeline import PipelineEvent
from progress_studio.config import SETTINGS
from progress_studio.presentation.gui.amount_mapping import AmountMappingFrame


STEP_LABELS = {
    "import-primavera-xml": "Import Primavera XML",
    "prepare-plan-actual-schedule": "Prepare plan / actual schedule",
    "build-timescale": "Build weekly timescale",
    "build-and-apply-amounts": "Build amount mapping",
    "build-progress-workbook": "Build progress workbook",
    "generate-plan-distribution": "Generate plan distribution",
    "build-okd-sheets": "Build OKD sheets",
}

DAYS = [
    ("1", "Monday"), ("2", "Tuesday"), ("3", "Wednesday"),
    ("4", "Thursday"), ("5", "Friday"), ("6", "Saturday"), ("7", "Sunday"),
]


class QueueWriter:
    def __init__(self, messages: queue.Queue[tuple[str, object]]) -> None:
        self._messages = messages

    def write(self, text: str) -> int:
        if text:
            self._messages.put(("log", text))
        return len(text)

    def flush(self) -> None:
        return None


class ProgressStudioDesktopApp(tk.Tk):
    def __init__(self, runner: DesktopRunner | None = None) -> None:
        super().__init__()
        self.runner = runner or DesktopRunner()
        self.messages: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.output_file: Path | None = None
        self.project_folder: Path | None = None

        self.title(f"{SETTINGS.title} Desktop {SETTINGS.version}")
        self.geometry("1050x720")
        self.minsize(900, 620)
        self.configure(bg="#f4f6f8")

        self.xml_var = tk.StringVar()
        self.cutoff_var = tk.StringVar(value="5 - Friday")
        self.amount_var = tk.StringVar(value=f"{SETTINGS.default_activity_amount:.0f}")
        self.distribution_var = tk.StringVar(value="auto")
        self.status_var = tk.StringVar(value="Ready")
        self.step_var = tk.StringVar(value="Select a Primavera XML file to begin.")
        self.progress_var = tk.DoubleVar(value=0)

        self._configure_style()
        self._build_ui()
        self.after(100, self._drain_messages)

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Segoe UI", 22, "bold"), background="#f4f6f8")
        style.configure("Subtitle.TLabel", font=("Segoe UI", 10), foreground="#52606d", background="#f4f6f8")
        style.configure("Card.TFrame", background="white", relief="solid", borderwidth=1)
        style.configure("Card.TLabel", background="white")
        style.configure("Section.TLabel", font=("Segoe UI", 12, "bold"), background="white")
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(18, 9))
        style.configure("TProgressbar", thickness=18)

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=22)
        root.pack(fill="both", expand=True)

        ttk.Label(root, text="Progress Studio", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            root,
            text="Desktop workflow for Primavera XML → Progress workbook → OKD",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(2, 18))

        content = ttk.Panedwindow(root, orient="horizontal")
        content.pack(fill="both", expand=True)

        left = ttk.Frame(content, style="Card.TFrame", padding=20)
        right = ttk.Frame(content, style="Card.TFrame", padding=20)
        content.add(left, weight=5)
        content.add(right, weight=4)

        ttk.Label(left, text="1. Project input", style="Section.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))
        ttk.Label(left, text="Primavera XML", style="Card.TLabel").grid(row=1, column=0, sticky="w")
        entry = ttk.Entry(left, textvariable=self.xml_var)
        entry.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(5, 5))
        ttk.Button(left, text="Browse...", command=self._browse_xml).grid(row=2, column=2, padx=(8, 0))
        ttk.Label(left, text="The original XML is read-only. Output is created on Desktop.", style="Card.TLabel", foreground="#68737d").grid(row=3, column=0, columnspan=3, sticky="w")

        ttk.Separator(left).grid(row=4, column=0, columnspan=3, sticky="ew", pady=18)
        ttk.Label(left, text="2. Processing options", style="Section.TLabel").grid(row=5, column=0, columnspan=3, sticky="w", pady=(0, 12))

        ttk.Label(left, text="Weekly cutoff day", style="Card.TLabel").grid(row=6, column=0, sticky="w", pady=5)
        cutoff = ttk.Combobox(left, textvariable=self.cutoff_var, state="readonly", values=[f"{n} - {name}" for n, name in DAYS])
        cutoff.grid(row=6, column=1, columnspan=2, sticky="ew", pady=5)

        ttk.Label(left, text="Fallback amount / activity", style="Card.TLabel").grid(row=7, column=0, sticky="w", pady=5)
        ttk.Entry(left, textvariable=self.amount_var).grid(row=7, column=1, columnspan=2, sticky="ew", pady=5)
        ttk.Label(left, text="Used only when the XML contains no activity amount data.", style="Card.TLabel", foreground="#68737d").grid(row=8, column=1, columnspan=2, sticky="w")

        ttk.Label(left, text="Plan distribution", style="Card.TLabel").grid(row=9, column=0, sticky="w", pady=(12, 5))
        dist = ttk.Combobox(left, textvariable=self.distribution_var, state="readonly", values=["auto", "flat", "front", "back", "bell"])
        dist.grid(row=9, column=1, columnspan=2, sticky="ew", pady=(12, 5))

        ttk.Separator(left).grid(row=10, column=0, columnspan=3, sticky="ew", pady=18)
        ttk.Label(left, text="3. Run", style="Section.TLabel").grid(row=11, column=0, columnspan=3, sticky="w", pady=(0, 12))
        self.run_button = ttk.Button(left, text="Create Progress Workbook", style="Accent.TButton", command=self._start)
        self.run_button.grid(row=12, column=0, columnspan=3, sticky="ew")
        self.open_file_button = ttk.Button(left, text="Open output workbook", command=self._open_output, state="disabled")
        self.open_file_button.grid(row=13, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        self.open_folder_button = ttk.Button(left, text="Open output folder", command=self._open_folder, state="disabled")
        self.open_folder_button.grid(row=14, column=0, columnspan=3, sticky="ew", pady=(8, 0))

        left.columnconfigure(0, weight=1)
        left.columnconfigure(1, weight=1)
        left.columnconfigure(2, weight=0)

        ttk.Label(right, text="Progress", style="Section.TLabel").pack(anchor="w")
        ttk.Label(right, textvariable=self.step_var, style="Card.TLabel", wraplength=380).pack(anchor="w", pady=(8, 8))
        ttk.Progressbar(right, variable=self.progress_var, maximum=100).pack(fill="x")
        ttk.Label(right, textvariable=self.status_var, style="Card.TLabel", foreground="#52606d").pack(anchor="w", pady=(5, 15))
        ttk.Separator(right).pack(fill="x", pady=(0, 12))
        details = ttk.Notebook(right)
        details.pack(fill="both", expand=True)

        log_tab = ttk.Frame(details, padding=6)
        mapping_tab = ttk.Frame(details, padding=0)
        details.add(log_tab, text="Activity log")
        details.add(mapping_tab, text="Amount Mapping")

        self.amount_mapping = AmountMappingFrame(mapping_tab)
        self.amount_mapping.pack(fill="both", expand=True)

        log_frame = ttk.Frame(log_tab)
        log_frame.pack(fill="both", expand=True)
        self.log = tk.Text(log_frame, wrap="word", font=("Consolas", 9), borderwidth=0, bg="#101820", fg="#d6e3ea", insertbackground="white")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scrollbar.set)
        self.log.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self._append_log("Progress Studio Desktop ready.\n")

    def _browse_xml(self) -> None:
        selected = filedialog.askopenfilename(title="Select Primavera XML", filetypes=[("Primavera XML", "*.xml"), ("All files", "*.*")])
        if selected:
            self.xml_var.set(selected)
            self.step_var.set("Input selected. Review options and run the pipeline.")

    def _start(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        try:
            xml = Path(self.xml_var.get().strip())
            amount = float(self.amount_var.get().replace(",", "").strip())
            cutoff = self.cutoff_var.get().split(" ", 1)[0]
            options = DesktopRunOptions(xml, cutoff, amount, self.distribution_var.get())
        except ValueError:
            messagebox.showerror("Invalid input", "Fallback amount must be a valid number.")
            return

        self.output_file = None
        self.project_folder = None
        self.progress_var.set(0)
        self.status_var.set("Running")
        self.step_var.set("Starting pipeline...")
        self.run_button.configure(state="disabled")
        self.open_file_button.configure(state="disabled")
        self.open_folder_button.configure(state="disabled")
        self.log.delete("1.0", "end")
        self._append_log(f"INPUT : {xml}\n")
        self._append_log(f"CUTOFF: {cutoff}\n")
        self._append_log(f"DISTRIBUTION: {options.distribution_method}\n\n")

        self.worker = threading.Thread(target=self._run_worker, args=(options,), daemon=True)
        self.worker.start()

    def _run_worker(self, options: DesktopRunOptions) -> None:
        writer = QueueWriter(self.messages)
        try:
            with redirect_stdout(writer), redirect_stderr(writer):
                result = self.runner.run(options, observer=lambda event: self.messages.put(("event", event)))
            self.messages.put(("done", result))
        except Exception as exc:
            self.messages.put(("error", exc))

    def _drain_messages(self) -> None:
        try:
            while True:
                kind, payload = self.messages.get_nowait()
                if kind == "log":
                    self._append_log(str(payload))
                elif kind == "event":
                    self._handle_event(payload)  # type: ignore[arg-type]
                elif kind == "done":
                    self._handle_done(payload)
                elif kind == "error":
                    self._handle_error(payload)  # type: ignore[arg-type]
        except queue.Empty:
            pass
        self.after(100, self._drain_messages)

    def _handle_event(self, event: PipelineEvent) -> None:
        label = STEP_LABELS.get(event.step_name, event.step_name)
        self.progress_var.set(event.progress_percent)
        if event.status == "started":
            self.step_var.set(f"Step {event.step_index}/{event.step_count}: {label}")
            self.status_var.set("Working...")
            self._append_log(f"\n[{event.step_index}/{event.step_count}] {label}\n")
        elif event.status == "completed":
            self.status_var.set(f"Completed step {event.step_index} of {event.step_count}")
            self._append_log("OK\n")

    def _handle_done(self, result: object) -> None:
        self.progress_var.set(100)
        self.status_var.set("Completed")
        self.step_var.set("Progress workbook is ready for Excel recalculation and OKD upload.")
        self.run_button.configure(state="normal")
        output = getattr(result, "output_workbook", None)
        project_folder = getattr(result, "project_folder", None)
        if output:
            self.output_file = Path(output)
            self.open_file_button.configure(state="normal")
            self._append_log(f"\nCOMPLETED: {self.output_file}\n")
            try:
                self.amount_mapping.set_progress_workbook(self.output_file)
            except Exception as exc:
                self._append_log(f"MAPPING WARNING: {exc}\n")
        if project_folder:
            self.project_folder = Path(project_folder)
            self.open_folder_button.configure(state="normal")
        messagebox.showinfo("Completed", f"Output created:\n{self.output_file or project_folder}")

    def _handle_error(self, error: Exception) -> None:
        self.status_var.set("Failed")
        self.step_var.set("Pipeline stopped. Review the activity log.")
        self.run_button.configure(state="normal")
        self._append_log(f"\nERROR: {error}\n")
        messagebox.showerror("Progress Studio", str(error))

    def _append_log(self, text: str) -> None:
        self.log.insert("end", text)
        self.log.see("end")

    def _open_output(self) -> None:
        if self.output_file:
            self._open_path(self.output_file)

    def _open_folder(self) -> None:
        if self.project_folder:
            self._open_path(self.project_folder)

    @staticmethod
    def _open_path(path: Path) -> None:
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            else:
                import webbrowser
                webbrowser.open(path.resolve().as_uri())
        except Exception as exc:
            messagebox.showerror("Open failed", str(exc))
