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
from progress_studio.presentation.gui.theme import PALETTE, FONT_MONO, configure_styles
from progress_studio.presentation.gui.strings import tr
from progress_studio.infrastructure.layout_preferences import (
    LayoutPreferences,
    LayoutPreferencesRepository,
)


STEP_LABELS = {
    "import-schedule-xml": "Import Schedule XML",
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
        self.configure(bg=PALETTE.canvas)

        self.xml_var = tk.StringVar()
        self.cutoff_var = tk.StringVar(value="5 - Friday")
        self.amount_var = tk.StringVar(value=f"{SETTINGS.default_activity_amount:.0f}")
        self.distribution_var = tk.StringVar(value="auto")
        self.status_var = tk.StringVar(value="Ready")
        self.step_var = tk.StringVar(value="Select an XML schedule file to begin.")
        self.progress_var = tk.DoubleVar(value=0)
        self.layout_repository = LayoutPreferencesRepository()
        self.layout_preferences = self.layout_repository.load()
        self.generator_collapsed = self.layout_preferences.generator_collapsed
        self.sidebar_collapsed = self.layout_preferences.sidebar_collapsed
        self.focus_mapping = False

        self._configure_style()
        self._build_ui()
        self.after_idle(self._maximize_window)
        self.after_idle(lambda: self._set_generator_collapsed(self.generator_collapsed, persist=False))
        self.after_idle(lambda: self._set_sidebar_collapsed(self.sidebar_collapsed, persist=False))
        self.bind("<F11>", lambda _event: self._toggle_focus_mapping())
        self.bind("<Escape>", lambda _event: self._exit_focus_mapping())
        self.protocol("WM_DELETE_WINDOW", self._close_application)
        self.after(100, self._drain_messages)

    def _configure_style(self) -> None:
        configure_styles(self)

    def _build_ui(self) -> None:
        # Production shell: menu, sidebar, workflow stages, workspace and status bar.
        self._build_menu()

        shell = ttk.Frame(self, style="App.TFrame")
        shell.pack(fill="both", expand=True)
        shell.columnconfigure(1, weight=1)
        shell.rowconfigure(0, weight=1)

        self.shell = shell
        self.sidebar = ttk.Frame(shell, style="Sidebar.TFrame", width=190)
        self.sidebar.grid(row=0, column=0, sticky="nsw")
        self.sidebar.grid_propagate(False)
        self._build_sidebar(self.sidebar)

        main = ttk.Frame(shell, style="App.TFrame", padding=(10, 8, 10, 0))
        main.grid(row=0, column=1, sticky="nsew")
        main.rowconfigure(2, weight=1)
        main.columnconfigure(0, weight=1)

        self.main_frame = main
        self.topbar = ttk.Frame(main, style="Surface.TFrame", padding=(12, 8))
        topbar = self.topbar
        topbar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(topbar, text="☰", width=3, command=self._toggle_sidebar).pack(side="left", padx=(0, 8))
        ttk.Label(topbar, text=tr("mapping_workspace"), style="Section.TLabel").pack(side="left")
        ttk.Button(topbar, text="Focus Mapping", command=self._toggle_focus_mapping).pack(side="left", padx=(12, 0))
        ttk.Label(topbar, text=tr("project_default"), style="Muted.TLabel").pack(side="right", padx=(16, 0))
        ttk.Button(topbar, text="EN", width=5).pack(side="right", padx=(8, 0))
        ttk.Button(topbar, text="?", width=3).pack(side="right", padx=(8, 0))

        self.stages = ttk.Frame(main, style="Surface.TFrame", padding=(8, 6))
        stages = self.stages
        stages.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        for index, label in enumerate(("1  Import", "2  Validate", "3  Mapping", "4  Allocation", "5  Review", "6  Export")):
            style = "StageActive.TLabel" if index == 2 else "Stage.TLabel"
            ttk.Label(stages, text=label, style=style).pack(side="left", fill="x", expand=True, padx=2)

        self.content = ttk.Panedwindow(main, orient="horizontal")
        self.content.grid(row=2, column=0, sticky="nsew")

        self.generator_panel = ttk.Frame(self.content, style="Card.TFrame", padding=16)
        self.workspace_panel = ttk.Frame(self.content, style="Card.TFrame", padding=8)
        self.content.add(self.generator_panel, weight=3)
        self.content.add(self.workspace_panel, weight=10)
        self._build_generator_panel(self.generator_panel)
        self._build_workspace_panel(self.workspace_panel)

        self.status_bar = ttk.Frame(main, style="Surface.TFrame")
        status = self.status_bar
        status.grid(row=3, column=0, sticky="ew", pady=(6, 0))
        ttk.Label(status, text="●  Ready", style="Status.TLabel", foreground=PALETTE.success).pack(side="left")
        ttk.Label(status, text=f"Progress Studio {SETTINGS.version}  |  Database: Local", style="Status.TLabel").pack(side="right")

    def _build_menu(self) -> None:
        menu = tk.Menu(self)
        file_menu = tk.Menu(menu, tearoff=False)
        file_menu.add_command(label="Exit", command=self._close_application)
        menu.add_cascade(label="File", menu=file_menu)

        edit_menu = tk.Menu(menu, tearoff=False)
        edit_menu.add_command(label="Edit commands")
        menu.add_cascade(label="Edit", menu=edit_menu)

        view_menu = tk.Menu(menu, tearoff=False)
        view_menu.add_command(label="Toggle Sidebar", command=self._toggle_sidebar)
        view_menu.add_command(label="Toggle Workbook Generator", command=self._toggle_generator)
        view_menu.add_separator()
        view_menu.add_command(label="Focus Mapping    F11", command=self._toggle_focus_mapping)
        menu.add_cascade(label="View", menu=view_menu)

        for title in ("Mapping", "AI Helper", "Tools", "Help"):
            submenu = tk.Menu(menu, tearoff=False)
            submenu.add_command(label=f"{title} commands")
            menu.add_cascade(label=title, menu=submenu)
        self.configure(menu=menu)

    def _build_sidebar(self, sidebar: ttk.Frame) -> None:
        ttk.Label(sidebar, text="  ◉  Progress Studio", style="SidebarTitle.TLabel").pack(fill="x", pady=(15, 18))
        sections = (
            ("DASHBOARD", (("⌂  " + tr("overview"), False),)),
            ("MAPPING", (("▦  " + tr("mapping_workspace"), True), ("▤  " + tr("batch_mapping"), False), ("⚙  " + tr("mapping_rules"), False), ("◇  " + tr("mapping_memory"), False))),
            ("DATA", (("▣  " + tr("boq_data"), False), ("↔  " + tr("progress_activities"), False))),
            ("SETTINGS", (("⚙  " + tr("project_settings"), False), ("✦  " + tr("ai_settings"), False), ("⚙  " + tr("preferences"), False))),
        )
        for heading, items in sections:
            ttk.Label(sidebar, text=heading, style="Sidebar.TLabel", foreground="#9fb2c3").pack(fill="x", padx=12, pady=(9, 3))
            for label, active in items:
                ttk.Button(sidebar, text=label, style="SidebarActive.TButton" if active else "Sidebar.TButton").pack(fill="x", padx=7, pady=1)
        ttk.Label(sidebar, text=f"\n  v{SETTINGS.version}", style="Sidebar.TLabel").pack(side="bottom", fill="x", pady=12)

    def _build_generator_panel(self, left: ttk.Frame) -> None:
        header = ttk.Frame(left, style="Surface.TFrame")
        header.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        ttk.Label(header, text=tr("generator"), style="Section.TLabel").pack(side="left")
        self.focus_mapping_button = ttk.Button(header, text="Focus Mapping", command=self._toggle_generator)
        self.focus_mapping_button.pack(side="right")

        ttk.Label(left, text="Schedule XML", style="Card.TLabel").grid(row=1, column=0, sticky="w")
        ttk.Entry(left, textvariable=self.xml_var).grid(row=2, column=0, columnspan=2, sticky="ew", pady=5)
        ttk.Button(left, text="Browse...", command=self._browse_xml).grid(row=2, column=2, padx=(8, 0))
        ttk.Label(left, text="Weekly cutoff day", style="Card.TLabel").grid(row=3, column=0, sticky="w", pady=(14, 5))
        ttk.Combobox(left, textvariable=self.cutoff_var, state="readonly", values=[f"{n} - {name}" for n, name in DAYS]).grid(row=3, column=1, columnspan=2, sticky="ew", pady=(14, 5))
        ttk.Label(left, text="Fallback amount / activity", style="Card.TLabel").grid(row=4, column=0, sticky="w", pady=5)
        ttk.Entry(left, textvariable=self.amount_var).grid(row=4, column=1, columnspan=2, sticky="ew", pady=5)
        ttk.Label(left, text="Plan distribution", style="Card.TLabel").grid(row=5, column=0, sticky="w", pady=5)
        ttk.Combobox(left, textvariable=self.distribution_var, state="readonly", values=["auto", "flat", "front", "back", "bell"]).grid(row=5, column=1, columnspan=2, sticky="ew", pady=5)
        self.run_button = ttk.Button(left, text="Create Progress Workbook", style="Accent.TButton", command=self._start)
        self.run_button.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(18, 0))
        self.open_file_button = ttk.Button(left, text="Open output workbook", command=self._open_output, state="disabled")
        self.open_file_button.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        self.open_folder_button = ttk.Button(left, text="Open output folder", command=self._open_folder, state="disabled")
        self.open_folder_button.grid(row=8, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        ttk.Separator(left).grid(row=9, column=0, columnspan=3, sticky="ew", pady=16)
        ttk.Label(left, textvariable=self.step_var, style="Card.TLabel", wraplength=310).grid(row=10, column=0, columnspan=3, sticky="w")
        ttk.Progressbar(left, variable=self.progress_var, maximum=100).grid(row=11, column=0, columnspan=3, sticky="ew", pady=(8, 4))
        ttk.Label(left, textvariable=self.status_var, style="Muted.TLabel").grid(row=12, column=0, columnspan=3, sticky="w")
        left.columnconfigure(0, weight=1)
        left.columnconfigure(1, weight=1)

    def _build_workspace_panel(self, right: ttk.Frame) -> None:
        details = ttk.Notebook(right)
        details.pack(fill="both", expand=True)
        mapping_tab = ttk.Frame(details, padding=0, style="Surface.TFrame")
        log_tab = ttk.Frame(details, padding=6, style="Surface.TFrame")
        details.add(mapping_tab, text=tr("mapping_workspace"))
        details.add(log_tab, text=tr("activity_log"))
        self.amount_mapping = AmountMappingFrame(mapping_tab)
        self.amount_mapping.pack(fill="both", expand=True)
        log_frame = ttk.Frame(log_tab, style="Surface.TFrame")
        log_frame.pack(fill="both", expand=True)
        self.log = tk.Text(log_frame, wrap="word", font=(FONT_MONO, 9), borderwidth=0, bg="#101820", fg="#d6e3ea", insertbackground="white")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scrollbar.set)
        self.log.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self._append_log("Progress Studio Desktop ready.\n")

    def _maximize_window(self) -> None:
        try:
            self.state("zoomed")
        except tk.TclError:
            width = self.winfo_screenwidth()
            height = self.winfo_screenheight()
            self.geometry(f"{width}x{height}+0+0")


    def _toggle_sidebar(self) -> None:
        self._set_sidebar_collapsed(not self.sidebar_collapsed)

    def _set_sidebar_collapsed(self, collapsed: bool, persist: bool = True) -> None:
        self.sidebar_collapsed = collapsed
        if collapsed:
            self.sidebar.grid_remove()
        else:
            self.sidebar.grid()
        if persist:
            self._save_layout_preferences()

    def _toggle_focus_mapping(self) -> None:
        if self.focus_mapping:
            self._exit_focus_mapping()
            return
        self.focus_mapping = True
        self._focus_restore = (self.sidebar_collapsed, self.generator_collapsed)
        self._set_sidebar_collapsed(True, persist=False)
        self._set_generator_collapsed(True, persist=False)
        self.topbar.grid_remove()
        self.stages.grid_remove()
        self.status_bar.grid_remove()

    def _exit_focus_mapping(self) -> None:
        if not self.focus_mapping:
            return
        self.focus_mapping = False
        self.topbar.grid()
        self.stages.grid()
        self.status_bar.grid()
        sidebar_collapsed, generator_collapsed = getattr(
            self, "_focus_restore", (self.sidebar_collapsed, self.generator_collapsed)
        )
        self._set_sidebar_collapsed(sidebar_collapsed, persist=False)
        self._set_generator_collapsed(generator_collapsed, persist=False)

    def _toggle_generator(self) -> None:
        self._set_generator_collapsed(not self.generator_collapsed)

    def _set_generator_collapsed(self, collapsed: bool, persist: bool = True) -> None:
        self.generator_collapsed = collapsed
        panes = set(self.content.panes())
        generator_id = str(self.generator_panel)
        if collapsed and generator_id in panes:
            self.content.forget(self.generator_panel)
        elif not collapsed and generator_id not in panes:
            self.content.insert(0, self.generator_panel, weight=5)
        self.focus_mapping_button.configure(
            text="Show Generator" if collapsed else "Focus Mapping"
        )
        if persist:
            self._save_layout_preferences()

    def _save_layout_preferences(self) -> None:
        current = self.layout_repository.load()
        preferences = LayoutPreferences(
            mapping_inputs_collapsed=current.mapping_inputs_collapsed,
            generator_collapsed=self.generator_collapsed,
            sidebar_collapsed=self.sidebar_collapsed,
            focus_mapping=False,
            mapping_sash=current.mapping_sash,
        )
        try:
            self.layout_repository.save(preferences)
        except OSError:
            pass

    def _close_application(self) -> None:
        self._save_layout_preferences()
        self.destroy()

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
