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
from progress_studio.infrastructure.layout_preferences import LayoutPreferences, LayoutPreferencesRepository
from progress_studio.presentation.gui.amount_mapping import AmountMappingFrame
from progress_studio.presentation.gui.strings import tr
from progress_studio.presentation.gui.theme import FONT_MONO, PALETTE, configure_styles

STEP_LABELS = {
    "import-schedule-xml": "Import Schedule XML",
    "prepare-plan-actual-schedule": "Prepare plan / actual schedule",
    "build-timescale": "Build weekly timescale",
    "build-and-apply-amounts": "Build amount mapping",
    "build-progress-workbook": "Build progress workbook",
    "generate-plan-distribution": "Generate plan distribution",
    "build-okd-sheets": "Build OKD sheets",
}
DAYS = [(str(index), name) for index, name in enumerate(
    ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"), 1
)]


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
    """Style-B desktop shell: simple workspace navigation and contextual commands."""

    WORKSPACES = (
        ("home", "⌂", "Home"),
        ("import", "⇩", "Import"),
        ("mapping", "▦", "Mapping"),
        ("ai", "✦", "AI Helper"),
        ("export", "⇧", "Export"),
        ("settings", "⚙", "Settings"),
    )

    def __init__(self, runner: DesktopRunner | None = None) -> None:
        super().__init__()
        self.runner = runner or DesktopRunner()
        self.messages: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.output_file: Path | None = None
        self.project_folder: Path | None = None
        self.current_workspace = "mapping"

        self.title(f"{SETTINGS.title} Desktop {SETTINGS.version}")
        self.geometry("1180x760")
        self.minsize(960, 640)
        self.configure(bg=PALETTE.canvas)

        self.xml_var = tk.StringVar()
        self.cutoff_var = tk.StringVar(value="5 - Friday")
        self.amount_var = tk.StringVar(value=f"{SETTINGS.default_activity_amount:.0f}")
        self.distribution_var = tk.StringVar(value="auto")
        self.status_var = tk.StringVar(value="Ready")
        self.step_var = tk.StringVar(value="Select an XML schedule file to begin.")
        self.progress_var = tk.DoubleVar(value=0)
        self.workspace_title_var = tk.StringVar(value="Mapping Workspace")
        self.project_title_var = tk.StringVar(value="Local Workspace")

        self.layout_repository = LayoutPreferencesRepository()
        self.layout_preferences = self.layout_repository.load()
        self.sidebar_collapsed = self.layout_preferences.sidebar_collapsed
        self.focus_mapping = False

        configure_styles(self)
        self._build_ui()
        self.after_idle(self._maximize_window)
        self.after_idle(lambda: self._set_sidebar_collapsed(self.sidebar_collapsed, persist=False))
        self._bind_shortcuts()
        self.protocol("WM_DELETE_WINDOW", self._close_application)
        self.after(100, self._drain_messages)

    def _build_ui(self) -> None:
        self._build_menu()
        shell = ttk.Frame(self, style="App.TFrame")
        shell.pack(fill="both", expand=True)
        shell.columnconfigure(1, weight=1)
        shell.rowconfigure(0, weight=1)
        self.shell = shell

        self.sidebar = ttk.Frame(shell, style="Sidebar.TFrame", width=176)
        self.sidebar.grid(row=0, column=0, sticky="nsw")
        self.sidebar.grid_propagate(False)
        self._build_sidebar(self.sidebar)

        main = ttk.Frame(shell, style="App.TFrame")
        main.grid(row=0, column=1, sticky="nsew")
        main.rowconfigure(2, weight=1)
        main.columnconfigure(0, weight=1)
        self.main_frame = main

        self._build_header(main)
        self._build_command_bar(main)

        self.workspace_host = ttk.Frame(main, style="App.TFrame", padding=(10, 8, 10, 6))
        self.workspace_host.grid(row=2, column=0, sticky="nsew")
        self.workspace_host.rowconfigure(0, weight=1)
        self.workspace_host.columnconfigure(0, weight=1)

        self.workspace_frames: dict[str, ttk.Frame] = {}
        self._build_home_workspace()
        self._build_import_workspace()
        self._build_mapping_workspace()
        self._build_ai_workspace()
        self._build_export_workspace()
        self._build_settings_workspace()

        self.status_bar = ttk.Frame(main, style="StatusBar.TFrame", padding=(6, 2))
        self.status_bar.grid(row=3, column=0, sticky="ew")
        ttk.Label(self.status_bar, text="●  Ready", style="StatusReady.TLabel").pack(side="left")
        ttk.Label(self.status_bar, text=f"Progress Studio {SETTINGS.version}  |  Local database", style="Status.TLabel").pack(side="right")

        self._show_workspace("mapping")

    def _build_menu(self) -> None:
        menu = tk.Menu(self)
        file_menu = tk.Menu(menu, tearoff=False)
        file_menu.add_command(label="Open Project...", command=self._defer_mapping("open_project"), accelerator="Ctrl+O")
        file_menu.add_command(label="Recent Projects", command=self._defer_mapping("open_recent_project"))
        file_menu.add_separator()
        file_menu.add_command(label="Save", command=self._defer_mapping("save_project"), accelerator="Ctrl+S")
        file_menu.add_command(label="Save As...", command=self._defer_mapping("save_project_as"), accelerator="Ctrl+Shift+S")
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self._close_application)
        menu.add_cascade(label="File", menu=file_menu)

        edit_menu = tk.Menu(menu, tearoff=False)
        edit_menu.add_command(label="Undo", command=self._defer_mapping("undo_mapping"), accelerator="Ctrl+Z")
        edit_menu.add_command(label="Map Selection", command=self._defer_mapping("map_selection"))
        edit_menu.add_command(label="Unmap Selection", command=self._defer_mapping("unmap_selection"), accelerator="Delete")
        menu.add_cascade(label="Edit", menu=edit_menu)

        view_menu = tk.Menu(menu, tearoff=False)
        view_menu.add_command(label="Toggle Sidebar", command=self._toggle_sidebar)
        view_menu.add_command(label="Focus Mapping", command=self._toggle_focus_mapping, accelerator="F11")
        menu.add_cascade(label="View", menu=view_menu)

        tools_menu = tk.Menu(menu, tearoff=False)
        tools_menu.add_command(label="Import Workspace", command=lambda: self._show_workspace("import"))
        tools_menu.add_command(label="Export Workspace", command=lambda: self._show_workspace("export"))
        menu.add_cascade(label="Tools", menu=tools_menu)

        help_menu = tk.Menu(menu, tearoff=False)
        help_menu.add_command(label="About Progress Studio", command=lambda: messagebox.showinfo("Progress Studio", f"Progress Studio {SETTINGS.version}"))
        menu.add_cascade(label="Help", menu=help_menu)
        self.configure(menu=menu)

    def _build_sidebar(self, sidebar: ttk.Frame) -> None:
        ttk.Label(sidebar, text="  PS  Progress Studio", style="SidebarTitle.TLabel").pack(fill="x", pady=(17, 20))
        self.sidebar_buttons: dict[str, ttk.Button] = {}
        for key, icon, label in self.WORKSPACES:
            if key == "settings":
                ttk.Frame(sidebar, style="Sidebar.TFrame").pack(fill="both", expand=True)
            button = ttk.Button(sidebar, text=f"{icon}   {label}", style="Sidebar.TButton", command=lambda name=key: self._show_workspace(name))
            button.pack(fill="x", padx=8, pady=2)
            self.sidebar_buttons[key] = button
        ttk.Label(sidebar, text=f"Local Workspace\nVersion {SETTINGS.version}", style="SidebarMuted.TLabel").pack(fill="x", padx=14, pady=14)

    def _build_header(self, main: ttk.Frame) -> None:
        self.topbar = ttk.Frame(main, style="Surface.TFrame", padding=(12, 8))
        self.topbar.grid(row=0, column=0, sticky="ew")
        ttk.Button(self.topbar, text="☰", width=3, command=self._toggle_sidebar).pack(side="left", padx=(0, 10))
        ttk.Label(self.topbar, textvariable=self.workspace_title_var, style="WorkspaceTitle.TLabel").pack(side="left")
        ttk.Label(self.topbar, textvariable=self.project_title_var, style="Muted.TLabel").pack(side="right", padx=(12, 0))
        ttk.Label(self.topbar, text="Project", style="Muted.TLabel").pack(side="right")

    def _build_command_bar(self, main: ttk.Frame) -> None:
        self.command_bar = ttk.Frame(main, style="CommandBar.TFrame", padding=(10, 6))
        self.command_bar.grid(row=1, column=0, sticky="ew", padx=10, pady=(8, 0))
        commands = (
            ("Open", "open_project"), ("Save", "save_project"), ("Save As", "save_project_as"),
            ("Undo", "undo_mapping"), ("Map", "map_selection"), ("Unmap", "unmap_selection"),
        )
        for index, (label, method) in enumerate(commands):
            if index in (3, 4):
                ttk.Separator(self.command_bar, orient="vertical").pack(side="left", fill="y", padx=8)
            style = "Accent.TButton" if label == "Map" else "TButton"
            ttk.Button(self.command_bar, text=label, style=style, command=self._defer_mapping(method)).pack(side="left", padx=(0, 4))
        ttk.Button(self.command_bar, text="Export", command=lambda: self._show_workspace("export")).pack(side="right")

    def _new_workspace(self, key: str) -> ttk.Frame:
        frame = ttk.Frame(self.workspace_host, style="Card.TFrame", padding=10)
        frame.grid(row=0, column=0, sticky="nsew")
        self.workspace_frames[key] = frame
        return frame

    def _build_home_workspace(self) -> None:
        frame = self._new_workspace("home")
        hero = ttk.Frame(frame, style="Surface.TFrame", padding=28)
        hero.pack(fill="x")
        ttk.Label(hero, text="Progress Studio", style="Title.TLabel").pack(anchor="w")
        ttk.Label(hero, text="Open a recent project or start by loading your Progress and BOQ workbooks.", style="Muted.TLabel").pack(anchor="w", pady=(8, 18))
        actions = ttk.Frame(hero, style="Surface.TFrame")
        actions.pack(anchor="w")
        ttk.Button(actions, text="Open Project", style="Accent.TButton", command=self._defer_mapping("open_project")).pack(side="left")
        ttk.Button(actions, text="Recent Projects", command=self._defer_mapping("open_recent_project")).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Go to Mapping", command=lambda: self._show_workspace("mapping")).pack(side="left", padx=(8, 0))

    def _build_import_workspace(self) -> None:
        frame = self._new_workspace("import")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        heading = ttk.Frame(frame, style="Surface.TFrame")
        heading.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(heading, text="Create Progress Workbook", style="WorkspaceTitle.TLabel").pack(side="left")
        ttk.Button(heading, text="Go to Mapping", command=lambda: self._show_workspace("mapping")).pack(side="right")

        body = ttk.Panedwindow(frame, orient="horizontal")
        body.grid(row=1, column=0, sticky="nsew")
        generator = ttk.Frame(body, style="Surface.TFrame", padding=14)
        log_panel = ttk.Frame(body, style="Surface.TFrame", padding=8)
        body.add(generator, weight=4)
        body.add(log_panel, weight=7)
        self._build_generator_panel(generator)
        ttk.Label(log_panel, text="Activity Log", style="Section.TLabel").pack(anchor="w", pady=(0, 6))
        self.log = tk.Text(log_panel, wrap="word", font=(FONT_MONO, 9), borderwidth=0, bg=PALETTE.console_bg, fg=PALETTE.console_fg, insertbackground="white")
        scrollbar = ttk.Scrollbar(log_panel, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scrollbar.set)
        self.log.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self._append_log("Progress Studio Desktop ready.\n")

    def _build_mapping_workspace(self) -> None:
        frame = self._new_workspace("mapping")
        self.amount_mapping = AmountMappingFrame(frame)
        self.amount_mapping.pack(fill="both", expand=True)

    def _build_ai_workspace(self) -> None:
        frame = self._new_workspace("ai")
        self._placeholder(frame, "AI Helper", "Semantic suggestions will appear here after the mapping workspace is stable.")

    def _build_export_workspace(self) -> None:
        frame = self._new_workspace("export")
        panel = ttk.Frame(frame, style="Surface.TFrame", padding=28)
        panel.pack(fill="x")
        ttk.Label(panel, text="Export", style="Title.TLabel").pack(anchor="w")
        ttk.Label(panel, text="Create a mapped Progress workbook without changing the original input file.", style="Muted.TLabel").pack(anchor="w", pady=(8, 18))
        ttk.Button(panel, text="Export Mapped Workbook", style="Accent.TButton", command=self._defer_mapping("export_workbook")).pack(anchor="w")

    def _build_settings_workspace(self) -> None:
        frame = self._new_workspace("settings")
        self._placeholder(frame, "Settings", "Theme colors are loaded from progress_studio/config/theme.json.")

    @staticmethod
    def _placeholder(frame: ttk.Frame, title: str, description: str) -> None:
        panel = ttk.Frame(frame, style="Surface.TFrame", padding=28)
        panel.pack(fill="x")
        ttk.Label(panel, text=title, style="Title.TLabel").pack(anchor="w")
        ttk.Label(panel, text=description, style="Muted.TLabel").pack(anchor="w", pady=(8, 0))

    # Kept as a named method for the production-shell contract and future workspace plugins.
    def _build_workspace_panel(self, right: ttk.Frame) -> None:
        self.amount_mapping = AmountMappingFrame(right)
        self.amount_mapping.pack(fill="both", expand=True)

    def _show_workspace(self, key: str) -> None:
        frame = self.workspace_frames[key]
        frame.tkraise()
        self.current_workspace = key
        title = next(label for name, _icon, label in self.WORKSPACES if name == key)
        self.workspace_title_var.set(f"{title} Workspace" if key not in {"home", "settings"} else title)
        for name, button in self.sidebar_buttons.items():
            button.configure(style="SidebarActive.TButton" if name == key else "Sidebar.TButton")
        self.command_bar.grid_remove() if key in {"home", "settings"} else self.command_bar.grid()

    def _defer_mapping(self, method_name: str):
        def command() -> None:
            method = getattr(getattr(self, "amount_mapping", None), method_name, None)
            if method is not None:
                method()
        return command

    def _build_generator_panel(self, left: ttk.Frame) -> None:
        left.columnconfigure(1, weight=1)
        ttk.Label(left, text="Schedule XML", style="Section.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Entry(left, textvariable=self.xml_var).grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(left, text="Browse...", command=self._browse_xml).grid(row=1, column=1, padx=(8, 0), pady=(8, 0))
        ttk.Label(left, text="Weekly cutoff day").grid(row=2, column=0, sticky="w", pady=(12, 0))
        ttk.Combobox(left, textvariable=self.cutoff_var, state="readonly", values=tuple(f"{number} - {name}" for number, name in DAYS)).grid(row=2, column=1, sticky="ew", padx=(8, 0), pady=(12, 0))
        ttk.Label(left, text="Fallback amount / activity").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(left, textvariable=self.amount_var).grid(row=3, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))
        ttk.Label(left, text="Plan distribution").grid(row=4, column=0, sticky="w", pady=(8, 0))
        ttk.Combobox(left, textvariable=self.distribution_var, state="readonly", values=("auto", "flat", "front", "back", "bell")).grid(row=4, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))
        self.run_button = ttk.Button(left, text="Create Progress Workbook", style="Accent.TButton", command=self._start)
        self.run_button.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(18, 0))
        self.open_file_button = ttk.Button(left, text="Open output workbook", command=self._open_output, state="disabled")
        self.open_file_button.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.open_folder_button = ttk.Button(left, text="Open output folder", command=self._open_folder, state="disabled")
        self.open_folder_button.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Separator(left).grid(row=8, column=0, columnspan=2, sticky="ew", pady=18)
        ttk.Label(left, textvariable=self.step_var, style="Muted.TLabel", wraplength=350).grid(row=9, column=0, columnspan=2, sticky="w")
        ttk.Progressbar(left, variable=self.progress_var, maximum=100).grid(row=10, column=0, columnspan=2, sticky="ew", pady=(8, 4))
        ttk.Label(left, textvariable=self.status_var, style="Muted.TLabel").grid(row=11, column=0, columnspan=2, sticky="w")

    def _bind_shortcuts(self) -> None:
        for sequence in ("<Control-o>", "<Command-o>"):
            self.bind(sequence, lambda _event: self._defer_mapping("open_project")(), add="+")
        for sequence in ("<Control-s>", "<Command-s>"):
            self.bind(sequence, lambda _event: self._defer_mapping("save_project")(), add="+")
        for sequence in ("<Control-Shift-S>", "<Command-Shift-S>"):
            self.bind(sequence, lambda _event: self._defer_mapping("save_project_as")(), add="+")
        for sequence in ("<Control-z>", "<Command-z>"):
            self.bind(sequence, lambda _event: self._defer_mapping("undo_mapping")(), add="+")
        self.bind("<F11>", lambda _event: self._toggle_focus_mapping())
        self.bind("<Escape>", lambda _event: self._exit_focus_mapping())

    def _toggle_sidebar(self) -> None:
        self._set_sidebar_collapsed(not self.sidebar_collapsed)

    def _set_sidebar_collapsed(self, collapsed: bool, persist: bool = True) -> None:
        self.sidebar_collapsed = collapsed
        self.sidebar.grid_remove() if collapsed else self.sidebar.grid()
        if persist:
            self._save_layout_preferences()

    def _toggle_focus_mapping(self) -> None:
        if self.focus_mapping:
            self._exit_focus_mapping()
            return
        self.focus_mapping = True
        self._focus_restore = self.sidebar_collapsed
        self._show_workspace("mapping")
        self._set_sidebar_collapsed(True, persist=False)
        self.topbar.grid_remove()
        self.command_bar.grid_remove()
        self.status_bar.grid_remove()

    def _exit_focus_mapping(self) -> None:
        if not self.focus_mapping:
            return
        self.focus_mapping = False
        self.topbar.grid()
        self.command_bar.grid()
        self.status_bar.grid()
        self._set_sidebar_collapsed(getattr(self, "_focus_restore", False), persist=False)

    # Legacy generator toggle is intentionally retained as a workspace switch.
    def _toggle_generator(self) -> None:
        self._show_workspace("import" if self.current_workspace != "import" else "mapping")

    def _set_generator_collapsed(self, collapsed: bool, persist: bool = True) -> None:
        if not collapsed:
            self._show_workspace("import")

    def _save_layout_preferences(self) -> None:
        current = self.layout_repository.load()
        preferences = LayoutPreferences(
            mapping_inputs_collapsed=current.mapping_inputs_collapsed,
            generator_collapsed=True,
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

    def _maximize_window(self) -> None:
        try:
            self.state("zoomed")
        except tk.TclError:
            self.geometry(f"{self.winfo_screenwidth()}x{self.winfo_screenheight()}+0+0")

    def _browse_xml(self) -> None:
        selected = filedialog.askopenfilename(title="Select Primavera XML", filetypes=[("Primavera XML", "*.xml"), ("All files", "*.*")])
        if selected:
            self.xml_var.set(selected)
            self.step_var.set("Input selected. Review options and create the workbook.")

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
                if kind == "log": self._append_log(str(payload))
                elif kind == "event": self._handle_event(payload)  # type: ignore[arg-type]
                elif kind == "done": self._handle_done(payload)
                elif kind == "error": self._handle_error(payload)  # type: ignore[arg-type]
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
        self.step_var.set("Progress workbook is ready.")
        self.run_button.configure(state="normal")
        output = getattr(result, "output_workbook", None)
        project_folder = getattr(result, "project_folder", None)
        if output:
            self.output_file = Path(output)
            self.open_file_button.configure(state="normal")
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
        if hasattr(self, "log"):
            self.log.insert("end", text)
            self.log.see("end")

    def _open_output(self) -> None:
        if self.output_file: self._open_path(self.output_file)

    def _open_folder(self) -> None:
        if self.project_folder: self._open_path(self.project_folder)

    @staticmethod
    def _open_path(path: Path) -> None:
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}"')
        except Exception as exc:
            messagebox.showerror("Open failed", str(exc))
