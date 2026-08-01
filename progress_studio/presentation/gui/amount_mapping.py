from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from progress_studio.domain.mapping_session import WorkbookFingerprint
from progress_studio.services.boq_mapping_service import BOQMappingService
from progress_studio.services.mapping_store import MappingStore
from progress_studio.services.workbook_export_service import WorkbookExportService
from progress_studio.infrastructure.layout_preferences import (
    LayoutPreferences,
    LayoutPreferencesRepository,
)
from progress_studio.infrastructure.platform_paths import user_data_dir
from progress_studio.presentation.gui.theme import PALETTE
from progress_studio.presentation.gui.generation_progress import GenerationProgressDialog
from progress_studio.infrastructure.session import (
    MappingSessionRepository,
    RecentSessionRepository,
    SessionValidationError,
)


class AmountMappingFrame(ttk.Frame):
    ACTIVITY_PAGE_SIZE = 200
    BOQ_PAGE_SIZE = 300

    def __init__(self, master, service: BOQMappingService | None = None) -> None:
        super().__init__(master, padding=8)
        self.service = service or BOQMappingService()
        self.store = MappingStore(self.ACTIVITY_PAGE_SIZE, self.BOQ_PAGE_SIZE)
        self.session_repository = MappingSessionRepository()
        self.export_service = WorkbookExportService()
        self.recent_repository = RecentSessionRepository()
        self.session_file: Path | None = None
        self.progress_file: Path | None = None
        self.boq_file: Path | None = None
        self.boq_sheet: str | None = None
        self.progress_path_var = tk.StringVar(value="No Progress workbook loaded")
        self.boq_path_var = tk.StringVar(value="No BOQ workbook loaded")
        self.boq_sheet_var = tk.StringVar()
        self.input_status_var = tk.StringVar(value="Load both workbooks to begin mapping.")
        self.activity_filter = tk.StringVar()
        self.boq_filter = tk.StringVar()
        self.boq_wbs2_filter = tk.StringVar(value="All")
        self.boq_wbs3_filter = tk.StringVar(value="All")
        self.summary_var = tk.StringVar(value="Mapped 0.00 / 0.00 | Remaining 0.00 | Items 0/0")
        self.share_var = tk.StringVar(value="100")
        self.activity_page_var = tk.StringVar(value="Rows 0-0 of 0")
        self.boq_page_var = tk.StringVar(value="Rows 0-0 of 0")
        self.boq_selection_var = tk.StringVar(value="Selected 0 items | 0.00")
        self.boq_mapping_detail_var = tk.StringVar(value="Select a BOQ item to view all mapped activities.")
        self._wbs_header_paths: dict[str, tuple[tuple[str, str], ...]] = {}
        self.session_status_var = tk.StringVar(value="Project: unsaved")
        self.operation_status_var = tk.StringVar(value="Ready")
        self.activity_empty_var = tk.StringVar(value="No Progress workbook loaded")
        self.boq_empty_var = tk.StringVar(value="No BOQ worksheet loaded")
        self.layout_repository = LayoutPreferencesRepository()
        self.layout_preferences = self.layout_repository.load()
        self.inputs_collapsed = self.layout_preferences.mapping_inputs_collapsed
        self._build()
        self.after_idle(self._restore_mapping_sash)
        self.bind("<Destroy>", self._on_destroy, add="+")
        self.after_idle(self._bind_shortcuts)

    def _build(self) -> None:
        inputs_header = ttk.Frame(self)
        inputs_header.pack(fill="x", pady=(0, 4))
        self.inputs_toggle_button = ttk.Button(
            inputs_header, command=self._toggle_inputs, width=3
        )
        self.inputs_toggle_button.pack(side="left")
        ttk.Label(inputs_header, text="Workbook Inputs", font=("Segoe UI", 10, "bold")).pack(
            side="left", padx=(6, 0)
        )
        ttk.Label(inputs_header, textvariable=self.input_status_var, foreground=PALETTE.muted).pack(
            side="right"
        )

        self.inputs_frame = ttk.LabelFrame(self, text="Mapping Inputs", padding=8)
        self.inputs_frame.columnconfigure(1, weight=1)

        ttk.Label(self.inputs_frame, text="Progress workbook").grid(row=0, column=0, sticky="w")
        ttk.Entry(self.inputs_frame, textvariable=self.progress_path_var, state="readonly").grid(
            row=0, column=1, sticky="ew", padx=8
        )
        ttk.Button(self.inputs_frame, text="Load Progress...", command=self._browse_progress).grid(row=0, column=2)

        ttk.Label(self.inputs_frame, text="BOQ workbook").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(self.inputs_frame, textvariable=self.boq_path_var, state="readonly").grid(
            row=1, column=1, sticky="ew", padx=8, pady=(6, 0)
        )
        ttk.Button(self.inputs_frame, text="Load BOQ...", command=self._browse_boq).grid(row=1, column=2, pady=(6, 0))

        ttk.Label(self.inputs_frame, text="BOQ worksheet").grid(row=2, column=0, sticky="w", pady=(6, 0))
        self.boq_sheet_combo = ttk.Combobox(
            self.inputs_frame, textvariable=self.boq_sheet_var, state="disabled", values=()
        )
        self.boq_sheet_combo.grid(row=2, column=1, sticky="ew", padx=8, pady=(6, 0))
        self.load_sheet_button = ttk.Button(
            self.inputs_frame, text="Load selected sheet", command=self._load_selected_boq_sheet, state="disabled"
        )
        self.load_sheet_button.grid(row=2, column=2, pady=(6, 0))

        # Project commands live in the application shell. This workspace only
        # contains context actions for the selected Progress/BOQ nodes.

        self.body = ttk.Panedwindow(self, orient="horizontal")
        self.body.pack(fill="both", expand=True)
        left = ttk.Frame(self.body, padding=(0, 0, 6, 0))
        right = ttk.Frame(self.body, padding=(6, 0, 0, 0))
        self.body.add(left, weight=1)
        self.body.add(right, weight=1)

        activity_header = ttk.Frame(left)
        activity_header.pack(fill="x")
        ttk.Label(activity_header, text="Progress Activities", font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Button(activity_header, text="Add WBS", command=self._add_supplemental_wbs).pack(side="left", padx=(8, 0))
        ttk.Button(activity_header, text="Add Activity", command=self._add_supplemental_activity).pack(side="left", padx=(4, 0))
        ttk.Button(activity_header, text="Edit", command=self._edit_progress_node).pack(side="left", padx=(4, 0))
        ttk.Button(activity_header, text="Delete", command=self._delete_progress_node).pack(side="left", padx=(4, 0))
        ttk.Button(activity_header, text="Up", command=lambda: self._move_progress_node(-1)).pack(side="left", padx=(4, 0))
        ttk.Button(activity_header, text="Down", command=lambda: self._move_progress_node(1)).pack(side="left", padx=(4, 0))
        ttk.Button(activity_header, text="Indent", command=self._indent_progress_node).pack(side="left", padx=(4, 0))
        ttk.Button(activity_header, text="Outdent", command=self._outdent_progress_node).pack(side="left", padx=(4, 0))
        ttk.Button(activity_header, text="Move...", command=self._reparent_progress_node).pack(side="left", padx=(4, 0))
        ttk.Button(activity_header, text="Undo edit", command=self._undo_tree_edit).pack(side="left", padx=(8, 0))
        ttk.Button(activity_header, text="Redo edit", command=self._redo_tree_edit).pack(side="left", padx=(4, 0))
        ttk.Button(activity_header, text="Expand all", command=self._expand_all_wbs).pack(side="right")
        ttk.Button(activity_header, text="Collapse all", command=self._collapse_all_wbs).pack(side="right", padx=(0, 4))
        self.activity_empty_label = ttk.Label(left, textvariable=self.activity_empty_var, style="Empty.TLabel", anchor="center")
        self.activity_empty_label.pack(fill="x", pady=(12, 4))
        self._build_search(left, self.activity_filter, self._apply_activity_filter)
        self.activity_tree = self._tree(
            left,
            ("check", "id", "description", "amount"),
            ("✓", "Activity ID / WBS", "Description", "Amount"),
            (38, 125, 390, 105),
        )
        self.activity_tree.tag_configure("wbs_level_1", font=("Segoe UI", 9, "bold"))
        self.activity_tree.tag_configure("wbs_level_2", font=("Segoe UI", 9, "bold"))
        self.activity_tree.tag_configure("wbs_level_3", font=("Segoe UI", 9, "bold"))
        self.activity_tree.tag_configure("selected_wbs", background=PALETTE.selection)
        self.activity_tree.bind("<Button-1>", self._activity_click)
        self._build_pager(left, self.activity_page_var, self._activity_prev, self._activity_next)

        ttk.Label(right, text="BOQ Items", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.boq_empty_label = ttk.Label(right, textvariable=self.boq_empty_var, style="Empty.TLabel", anchor="center")
        self.boq_empty_label.pack(fill="x", pady=(12, 4))
        filter_row = ttk.Frame(right)
        filter_row.pack(fill="x", pady=(4, 4))
        ttk.Label(filter_row, text="WBS-2").pack(side="left")
        self.boq_wbs2_combo = ttk.Combobox(filter_row, textvariable=self.boq_wbs2_filter, state="readonly", width=24, values=("All",))
        self.boq_wbs2_combo.pack(side="left", padx=(5, 8))
        self.boq_wbs2_combo.bind("<<ComboboxSelected>>", self._on_wbs2_filter)
        ttk.Label(filter_row, text="WBS-3").pack(side="left")
        self.boq_wbs3_combo = ttk.Combobox(filter_row, textvariable=self.boq_wbs3_filter, state="readonly", width=28, values=("All",))
        self.boq_wbs3_combo.pack(side="left", padx=(5, 0))
        self.boq_wbs3_combo.bind("<<ComboboxSelected>>", self._on_wbs3_filter)
        self._build_search(right, self.boq_filter, self._apply_boq_filter)
        selection_row = ttk.Frame(right)
        selection_row.pack(fill="x", pady=(0, 4))
        ttk.Button(selection_row, text="Select page", command=self._select_boq_page).pack(side="left")
        ttk.Button(selection_row, text="Select all filtered", command=self._select_all_filtered_boq).pack(side="left", padx=(4, 0))
        ttk.Button(selection_row, text="Clear selection", command=self._clear_boq_selection).pack(side="left", padx=(4, 0))
        ttk.Label(selection_row, textvariable=self.boq_selection_var, foreground=PALETTE.muted).pack(side="right")
        self.boq_tree = self._tree(
            right,
            (
                "check", "wbs2", "wbs3", "wbs4", "description",
                "amount", "allocated", "remaining", "status", "mapped",
            ),
            (
                "✓", "WBS-2", "WBS-3", "WBS-4", "Description",
                "Amount", "Allocated", "Remaining %", "Status", "Mapped To",
            ),
            (38, 75, 105, 95, 310, 92, 92, 88, 72, 125),
        )
        self.boq_tree.bind("<Button-1>", self._boq_click)
        mapping_detail = ttk.Frame(right, padding=(0, 4, 0, 0))
        mapping_detail.pack(fill="x")
        ttk.Label(mapping_detail, text="Mapped activities", font=("Segoe UI", 9, "bold")).pack(side="left")
        ttk.Label(
            mapping_detail,
            textvariable=self.boq_mapping_detail_var,
            foreground=PALETTE.muted,
            anchor="w",
        ).pack(side="left", fill="x", expand=True, padx=(8, 0))
        self._build_pager(right, self.boq_page_var, self._boq_prev, self._boq_next)

        actions = ttk.Frame(self)
        actions.pack(fill="x", pady=(6, 0))
        ttk.Label(actions, text="Share").pack(side="left")
        share_entry = ttk.Entry(actions, textvariable=self.share_var, width=8, justify="right")
        share_entry.pack(side="left", padx=(5, 2))
        ttk.Label(actions, text="%").pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Map", command=self._map).pack(side="left")
        ttk.Button(actions, text="Undo", command=self._undo).pack(side="left", padx=(6, 0))
        ttk.Button(actions, text="Unmap", command=self._unmap).pack(side="left", padx=(6, 0))
        ttk.Button(actions, text="Clear all", command=self._clear_all).pack(side="left", padx=(6, 0))
        ttk.Label(
            actions,
            text="Share applies to every selected BOQ item.",
            foreground=PALETTE.muted,
        ).pack(side="left", padx=(12, 0))

        self.mapping_status_bar = ttk.Frame(self, style="StatusBar.TFrame", padding=(8, 4))
        self.mapping_status_bar.pack(fill="x", pady=(6, 0))
        ttk.Label(self.mapping_status_bar, textvariable=self.operation_status_var, style="StatusReady.TLabel").pack(side="left")
        ttk.Label(self.mapping_status_bar, textvariable=self.session_status_var, style="Status.TLabel").pack(side="left", padx=(12, 0))
        ttk.Label(self.mapping_status_bar, textvariable=self.summary_var, style="Status.TLabel").pack(side="right")

        self.loading_overlay = ttk.Frame(self, style="Loading.TFrame", padding=24)
        ttk.Label(self.loading_overlay, text="Working...", style="LoadingTitle.TLabel").pack()
        self.loading_progress = ttk.Progressbar(self.loading_overlay, mode="indeterminate", length=260)
        self.loading_progress.pack(pady=(10, 0))

        self._set_inputs_collapsed(self.inputs_collapsed, persist=False)

    # Application-shell command surface. Internal repository names remain
    # session-based for backward compatibility, while users work with Projects.
    def open_progress(self) -> None:
        self._browse_progress()

    def open_boq(self) -> None:
        self._browse_boq()

    def open_project(self) -> None:
        self._load_session()

    def open_recent_project(self) -> None:
        self._load_recent_session()

    def save_project(self) -> None:
        if self.session_file and self.session_file.suffix == ".progressstudio":
            try:
                self._write_session(self.session_file, show_message=False)
                self._notify("Project saved")
            except Exception as exc:
                messagebox.showerror("Progress Studio", str(exc))
        else:
            self.save_project_as()

    def save_project_as(self) -> None:
        self._save_session()

    def undo_mapping(self) -> None:
        self._undo()

    def map_selection(self) -> None:
        self._map()

    def unmap_selection(self) -> None:
        self._unmap()

    def export_workbook(self) -> None:
        self._export()

    def _bind_shortcuts(self) -> None:
        top = self.winfo_toplevel()
        for sequence in ("<Control-o>", "<Command-o>"):
            top.bind(sequence, lambda _event: self._browse_progress(), add="+")
        for sequence in ("<Control-Shift-O>", "<Command-Shift-O>"):
            top.bind(sequence, lambda _event: self._browse_boq(), add="+")
        for sequence in ("<Control-s>", "<Command-s>"):
            top.bind(sequence, lambda _event: self._save_session(), add="+")
        for sequence in ("<Control-z>", "<Command-z>"):
            top.bind(sequence, lambda _event: self._undo(), add="+")
        for sequence in ("<Control-f>", "<Command-f>"):
            top.bind(sequence, self._focus_search, add="+")
        top.bind("<Alt-Up>", lambda _event: self._move_progress_node(-1), add="+")
        top.bind("<Alt-Down>", lambda _event: self._move_progress_node(1), add="+")
        top.bind("<Alt-Right>", lambda _event: self._indent_progress_node(), add="+")
        top.bind("<Alt-Left>", lambda _event: self._outdent_progress_node(), add="+")
        top.bind("<Delete>", self._keyboard_unmap, add="+")
        top.bind("<BackSpace>", self._keyboard_unmap, add="+")

    def _focus_search(self, _event=None) -> str:
        self.activity_search_entry.focus_set()
        self.activity_search_entry.selection_range(0, "end")
        return "break"

    def _keyboard_unmap(self, event) -> str | None:
        widget_class = event.widget.winfo_class()
        if widget_class in {"Entry", "TEntry", "Text", "TCombobox", "Spinbox"}:
            return None
        self._unmap()
        return "break"

    def _notify(self, message: str, kind: str = "success") -> None:
        self.operation_status_var.set(message)
        top = self.winfo_toplevel()
        toast = tk.Toplevel(top)
        toast.overrideredirect(True)
        toast.attributes("-topmost", True)
        background = PALETTE.success if kind == "success" else PALETTE.warning
        label = tk.Label(toast, text=message, bg=background, fg="white", padx=16, pady=10)
        label.pack()
        top.update_idletasks()
        toast.update_idletasks()
        x = top.winfo_rootx() + max(0, top.winfo_width() - toast.winfo_width() - 24)
        y = top.winfo_rooty() + max(0, top.winfo_height() - toast.winfo_height() - 48)
        toast.geometry(f"+{x}+{y}")
        toast.after(2200, toast.destroy)

    def _toggle_inputs(self) -> None:
        self._set_inputs_collapsed(not self.inputs_collapsed)

    def _set_inputs_collapsed(self, collapsed: bool, persist: bool = True) -> None:
        self.inputs_collapsed = collapsed
        self.inputs_toggle_button.configure(text="▶" if collapsed else "▼")
        if collapsed:
            self.inputs_frame.pack_forget()
        elif not self.inputs_frame.winfo_manager():
            self.inputs_frame.pack(fill="x", after=self.inputs_toggle_button.master, pady=(0, 6))
        if persist:
            self._save_layout_preferences()

    def _restore_mapping_sash(self) -> None:
        if self.layout_preferences.mapping_sash is None:
            return
        try:
            self.body.sashpos(0, self.layout_preferences.mapping_sash)
        except tk.TclError:
            pass

    def _save_layout_preferences(self) -> None:
        sash = None
        try:
            sash = self.body.sashpos(0)
        except (AttributeError, tk.TclError):
            pass
        self.layout_preferences = LayoutPreferences(
            mapping_inputs_collapsed=self.inputs_collapsed,
            generator_collapsed=self.layout_preferences.generator_collapsed,
            sidebar_collapsed=self.layout_preferences.sidebar_collapsed,
            focus_mapping=False,
            mapping_sash=sash,
        )
        try:
            self.layout_repository.save(self.layout_preferences)
        except OSError:
            pass

    def _on_destroy(self, event) -> None:
        if event.widget is self:
            self._save_layout_preferences()

    def _build_search(self, parent, variable: tk.StringVar, command) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(4, 6))
        entry = ttk.Entry(row, textvariable=variable)
        entry.pack(side="left", fill="x", expand=True)
        if variable is self.activity_filter:
            self.activity_search_entry = entry
        elif variable is self.boq_filter:
            self.boq_search_entry = entry
        entry.bind("<Return>", lambda _event: command())
        ttk.Button(row, text="Search", command=command).pack(side="left", padx=(6, 0))
        ttk.Button(row, text="Clear", command=lambda: (variable.set(""), command())).pack(side="left", padx=(4, 0))

    @staticmethod
    def _build_pager(parent, variable: tk.StringVar, previous, next_) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(5, 0))
        ttk.Label(row, textvariable=variable).pack(side="left")
        ttk.Button(row, text="Previous", command=previous).pack(side="right")
        ttk.Button(row, text="Next", command=next_).pack(side="right", padx=(0, 5))

    @staticmethod
    def _tree(parent, columns, headings, widths):
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="none")
        y = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=y.set)
        for name, heading, width in zip(columns, headings, widths):
            tree.heading(name, text=heading)
            tree.column(name, width=width, minwidth=35, stretch=name not in {"check", "amount", "allocated", "remaining", "status", "mapped"})
        tree.grid(row=0, column=0, sticky="nsew")
        y.grid(row=0, column=1, sticky="ns")

        # Scroll only the table currently under the pointer. This keeps the
        # Activity and BOQ panes independent on Windows and macOS.
        def on_mousewheel(event):
            delta = -1 if event.delta > 0 else 1
            tree.yview_scroll(delta * 3, "units")
            return "break"

        tree.bind("<Enter>", lambda _event: tree.bind_all("<MouseWheel>", on_mousewheel))
        tree.bind("<Leave>", lambda _event: tree.unbind_all("<MouseWheel>"))
        return tree

    def _detach_session(self) -> None:
        self.session_file = None
        self.session_status_var.set("Session: not saved")

    def _browse_progress(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select Progress workbook",
            filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")],
        )
        if selected:
            self.set_progress_workbook(Path(selected))

    def set_progress_workbook(self, path: Path) -> None:
        self.progress_file = Path(path)
        self._detach_session()
        try:
            self._busy(True)
            rows = self.service.read_activities(self.progress_file)
            self.store.load_activities(rows)
            self.progress_path_var.set(str(self.progress_file))
            self._render_activities()
            self._update_summary()
            self._update_input_status()
            if self.progress_file and self.boq_file and self.boq_sheet:
                self._set_inputs_collapsed(True)
        except Exception as exc:
            self.progress_file = None
            self.progress_path_var.set("No Progress workbook loaded")
            messagebox.showerror("Amount Mapping", str(exc))
        finally:
            self._busy(False)

    def _browse_boq(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select BOQ workbook",
            filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")],
        )
        if not selected:
            return
        candidate = Path(selected)
        self._detach_session()
        try:
            self._busy(True)
            sheet_names = self.service.list_boq_sheets(candidate)
            if not sheet_names:
                raise ValueError("The selected BOQ workbook has no worksheets.")
            self.boq_file = candidate
            self.boq_sheet = None
            self.boq_path_var.set(str(candidate))
            self.boq_sheet_combo.configure(values=sheet_names, state="readonly")
            self.boq_sheet_var.set(sheet_names[0])
            self.load_sheet_button.configure(state="normal")
            self.store.load_boq([])
            self._refresh_boq_filter_values()
            self._render_boq()
            self._update_summary()
            self._update_input_status()
        except Exception as exc:
            self.boq_file = None
            self.boq_path_var.set("No BOQ workbook loaded")
            self.boq_sheet_combo.configure(values=(), state="disabled")
            self.boq_sheet_var.set("")
            self.load_sheet_button.configure(state="disabled")
            messagebox.showerror("Amount Mapping", str(exc))
        finally:
            self._busy(False)

    def _load_selected_boq_sheet(self) -> None:
        if not self.boq_file:
            messagebox.showwarning("Amount Mapping", "Load a BOQ workbook first.")
            return
        sheet_name = self.boq_sheet_var.get().strip()
        if not sheet_name:
            messagebox.showwarning("Amount Mapping", "Select a BOQ worksheet.")
            return
        try:
            self._detach_session()
            self._busy(True)
            rows = self.service.read_boq(self.boq_file, sheet_name)
            self.boq_sheet = sheet_name
            self.store.load_boq(rows)
            self._refresh_boq_filter_values()
            self._render_boq()
            self._render_activities()
            self._update_summary()
            self._update_input_status()
            self._set_inputs_collapsed(True)
        except Exception as exc:
            messagebox.showerror("Amount Mapping", str(exc))
        finally:
            self._busy(False)

    def _update_input_status(self) -> None:
        activity_count = len(self.store.activity_order)
        boq_count = len(self.store.boq_order)
        progress_text = f"Progress: {activity_count:,} activities" if self.progress_file else "Progress: not loaded"
        if self.boq_file and self.boq_sheet:
            boq_text = f"BOQ: {boq_count:,} items from {self.boq_sheet}"
        elif self.boq_file:
            boq_text = "BOQ: workbook loaded; select a worksheet"
        else:
            boq_text = "BOQ: not loaded"
        self.input_status_var.set(f"{progress_text} | {boq_text}")

    def _busy(self, active: bool, message: str = "Working...") -> None:
        top = self.winfo_toplevel()
        top.configure(cursor="watch" if active else "")
        self.operation_status_var.set(message if active else "Ready")
        if active:
            self.loading_overlay.place(relx=0.5, rely=0.5, anchor="center")
            self.loading_overlay.lift()
            self.loading_progress.start(12)
        else:
            self.loading_progress.stop()
            self.loading_overlay.place_forget()
        top.update_idletasks()


    def _refresh_boq_filter_values(self) -> None:
        wbs2_values = ("All",) + self.store.boq_wbs2_values()
        self.boq_wbs2_combo.configure(values=wbs2_values)
        self.boq_wbs2_filter.set("All")
        wbs3_values = ("All",) + self.store.boq_wbs3_values()
        self.boq_wbs3_combo.configure(values=wbs3_values)
        self.boq_wbs3_filter.set("All")

    def _on_wbs2_filter(self, _event=None) -> None:
        selected = self.boq_wbs2_filter.get()
        self.store.boq_wbs2 = "" if selected == "All" else selected
        values = ("All",) + self.store.boq_wbs3_values(self.store.boq_wbs2)
        self.boq_wbs3_combo.configure(values=values)
        self.boq_wbs3_filter.set("All")
        self.store.boq_wbs3 = ""
        self.store.boq_page = 1
        self._render_boq()

    def _on_wbs3_filter(self, _event=None) -> None:
        selected = self.boq_wbs3_filter.get()
        self.store.boq_wbs3 = "" if selected == "All" else selected
        self.store.boq_page = 1
        self._render_boq()

    def _selected_parent_path(self) -> tuple[tuple[str, str], ...]:
        if self.store.selected_wbs_path:
            return self.store.selected_wbs_path
        selected = next(iter(self.store.selected_activity_ids), "")
        row = self.store.activities_by_id.get(selected)
        return row.wbs_path if row else ()

    def _add_supplemental_wbs(self) -> None:
        if not self.store.activities_by_id:
            messagebox.showwarning("Progress tree", "Load a Progress workbook first.")
            return
        parent_path = self._selected_parent_path()
        code = simpledialog.askstring("Add WBS", "WBS code:", parent=self)
        if code is None:
            return
        name = simpledialog.askstring("Add WBS", "WBS name:", parent=self)
        if name is None:
            return
        try:
            self.store.add_supplemental_wbs(parent_path=parent_path, code=code, name=name)
            self._render_activities()
            self._autosave_session()
            self._notify("WBS created. It is selected for the next sub-WBS or Activity.")
        except ValueError as exc:
            messagebox.showerror("Progress tree", str(exc))

    def _add_supplemental_activity(self) -> None:
        if not self.store.activities_by_id:
            messagebox.showwarning("Progress tree", "Load a Progress workbook first.")
            return
        parent_path = self._selected_parent_path()
        if not parent_path:
            messagebox.showwarning("Progress tree", "Select a WBS first.")
            return
        activity_id = simpledialog.askstring("Add Activity", "Activity ID (must be unique):", parent=self)
        if activity_id is None:
            return
        description = simpledialog.askstring("Add Activity", "Activity name:", parent=self)
        if description is None:
            return
        try:
            self.store.add_supplemental_activity(
                parent_path=parent_path,
                wbs_code=parent_path[-1][0],
                wbs_name=parent_path[-1][1],
                activity_id=activity_id,
                description=description,
            )
            # Keep the parent visible and select the newly created Activity.
            self.store.selected_wbs_path = ()
            self.store.collapsed_wbs_paths.discard(self.store.wbs_path_key(parent_path))
            self._render_activities()
            self._update_summary()
            self._autosave_session()
            self._notify("Activity created")
        except ValueError as exc:
            messagebox.showerror("Progress tree", str(exc))

    def _edit_progress_node(self) -> None:
        node = self.store.selected_working_node()
        if node is None:
            messagebox.showwarning("Progress tree", "Select a WBS or Activity to edit.")
            return
        title = "Edit WBS" if node.kind.value == "wbs" else "Edit Activity"
        code_label = "WBS code:" if node.kind.value == "wbs" else "Activity ID:"
        code = simpledialog.askstring(title, code_label, initialvalue=node.code, parent=self)
        if code is None:
            return
        name = simpledialog.askstring(title, "Name:", initialvalue=node.name, parent=self)
        if name is None:
            return
        try:
            self.store.edit_selected_node(code=code, name=name)
            self._render_activities()
            self._render_boq()
            self._update_summary()
            self._autosave_session()
            self._notify("Progress tree updated")
        except ValueError as exc:
            messagebox.showerror("Progress tree", str(exc))

    def _delete_progress_node(self) -> None:
        node = self.store.selected_working_node()
        if node is None:
            messagebox.showwarning("Progress tree", "Select a WBS or Activity to delete.")
            return
        label = f"{node.code} — {node.name}"
        if not messagebox.askyesno(
            "Delete progress node",
            f"Delete {label}?\n\nChild nodes will also be removed from the working tree. "
            "The source workbook will not be overwritten.",
        ):
            return
        try:
            self.store.delete_selected_node()
            self._render_activities()
            self._render_boq()
            self._update_summary()
            self._autosave_session()
            self._notify("Progress node deleted")
        except ValueError as exc:
            messagebox.showerror("Progress tree", str(exc))

    def _move_progress_node(self, offset: int) -> None:
        try:
            if not self.store.move_selected_node(offset):
                self._notify("Node is already at the edge of its group")
                return
            self._render_activities()
            self._autosave_session()
            self._notify("Progress node reordered")
        except ValueError as exc:
            messagebox.showerror("Progress tree", str(exc))

    def _indent_progress_node(self) -> None:
        try:
            self.store.indent_selected_node()
            self._render_activities()
            self._render_boq()
            self._update_summary()
            self._autosave_session()
            self._notify("Progress node indented")
        except ValueError as exc:
            messagebox.showerror("Progress tree", str(exc))

    def _outdent_progress_node(self) -> None:
        try:
            self.store.outdent_selected_node()
            self._render_activities()
            self._render_boq()
            self._update_summary()
            self._autosave_session()
            self._notify("Progress node outdented")
        except ValueError as exc:
            messagebox.showerror("Progress tree", str(exc))

    def _reparent_progress_node(self) -> None:
        node = self.store.selected_working_node()
        if node is None:
            messagebox.showwarning("Progress tree", "Select a WBS or Activity to move.")
            return
        choices = [
            item for item in self.store.working_tree_nodes()
            if item.kind.value == "wbs" and item.node_id != node.node_id
        ]
        if not choices:
            messagebox.showwarning("Progress tree", "No destination WBS is available.")
            return
        dialog = tk.Toplevel(self)
        dialog.title("Move progress node")
        dialog.transient(self.winfo_toplevel())
        dialog.grab_set()
        dialog.geometry("520x340")
        ttk.Label(dialog, text=f"Move {node.code} — {node.name} under:").pack(anchor="w", padx=12, pady=(12, 6))
        box = tk.Listbox(dialog, height=13)
        box.pack(fill="both", expand=True, padx=12)
        ids: list[str] = []
        for item in choices:
            path = " > ".join(part.code for part in self.store.working_tree.path_for(item.node_id) if part.kind.value == "wbs")
            box.insert("end", f"{path} — {item.name}")
            ids.append(item.node_id)
        if ids:
            box.selection_set(0)
        selected: list[str] = []
        def accept() -> None:
            indexes = box.curselection()
            if indexes:
                selected.append(ids[indexes[0]])
                dialog.destroy()
        row = ttk.Frame(dialog)
        row.pack(fill="x", padx=12, pady=12)
        ttk.Button(row, text="Move", command=accept).pack(side="right")
        ttk.Button(row, text="Cancel", command=dialog.destroy).pack(side="right", padx=(0, 8))
        box.bind("<Double-1>", lambda _event: accept())
        dialog.wait_window()
        if not selected:
            return
        try:
            self.store.reparent_selected_node(selected[0])
            self._render_activities()
            self._autosave_session()
            self._notify("Progress node moved")
        except ValueError as exc:
            messagebox.showerror("Progress tree", str(exc))

    def _undo_tree_edit(self) -> None:
        if not self.store.undo_tree_edit():
            self._notify("Nothing to undo in the progress tree")
            return
        self._render_activities()
        self._render_boq()
        self._update_summary()
        self._autosave_session()
        self._notify("Tree edit undone")

    def _redo_tree_edit(self) -> None:
        if not self.store.redo_tree_edit():
            self._notify("Nothing to redo in the progress tree")
            return
        self._render_activities()
        self._render_boq()
        self._update_summary()
        self._autosave_session()
        self._notify("Tree edit redone")

    def _apply_activity_filter(self) -> None:
        self.store.activity_query = self.activity_filter.get()
        self.store.activity_page = 1
        self._render_activities()

    def _apply_boq_filter(self) -> None:
        self.store.boq_query = self.boq_filter.get()
        self.store.boq_page = 1
        self._render_boq()

    def _render_activities(self) -> None:
        """Render the recursive working tree in model order.

        WBS and Activity rows now come from one tree walk. Created Activities
        therefore appear below their selected parent instead of being appended
        after the workbook-derived flat activity list.
        """
        self.activity_tree.delete(*self.activity_tree.get_children())
        self._wbs_header_paths.clear()
        page = self.store.activity_page_data()
        page_activity_ids = set(page.ids)
        query_active = bool(self.store.activity_query.strip())

        relevant_node_ids: set[str] = set()
        for activity_id in page_activity_ids:
            node = self.store.working_tree.find_activity(activity_id)
            if node is None:
                continue
            relevant_node_ids.add(node.node_id)
            relevant_node_ids.update(item.node_id for item in self.store.working_tree.path_for(node.node_id))

        # Keep user-created WBS nodes visible even before they receive their
        # first Activity, so the editor remains usable as a true tree editor.
        for node in self.store.working_tree_nodes():
            if node.kind.value == "wbs" and node.origin.value == "user_created":
                relevant_node_ids.add(node.node_id)
                relevant_node_ids.update(item.node_id for item in self.store.working_tree.path_for(node.node_id))

        rendered_activities = 0

        def visit(parent_id: str | None, depth: int, ancestor_collapsed: bool = False) -> None:
            nonlocal rendered_activities
            for node in self.store.working_tree.children(parent_id):
                if node.node_id not in relevant_node_ids:
                    continue
                if ancestor_collapsed and not query_active:
                    continue

                if node.kind.value == "wbs":
                    path = tuple(
                        (item.code, item.name)
                        for item in self.store.working_tree.path_for(node.node_id)
                        if item.kind.value == "wbs"
                    )
                    path_key = self.store.wbs_path_key(path)
                    collapsed = path_key in self.store.collapsed_wbs_paths
                    header_id = f"__wbs__{node.node_id}"
                    self._wbs_header_paths[header_id] = path
                    indent = "    " * max(depth - 1, 0)
                    self.activity_tree.insert(
                        "",
                        "end",
                        iid=header_id,
                        values=("", f"{indent}{'▶' if collapsed else '▼'} {node.code}", node.name, ""),
                        tags=(
                            f"wbs_level_{min(depth, 3)}",
                            "wbs_header",
                            "selected_wbs" if node.node_id == self.store.selected_node_id else "",
                        ),
                    )
                    visit(node.node_id, depth + 1, ancestor_collapsed=collapsed and not query_active)
                    continue

                if node.code not in page_activity_ids:
                    continue
                row = self.store.activities_by_id.get(node.code)
                if row is None:
                    continue
                check = "☑" if node.code in self.store.selected_activity_ids else "☐"
                indent = "    " * max(depth - 1, 0)
                self.activity_tree.insert(
                    "",
                    "end",
                    iid=node.code,
                    values=(
                        check,
                        f"{indent}{node.code}",
                        node.name,
                        f"{self.store.activity_amount(node.code):,.2f}",
                    ),
                    tags=("selected_activity",) if node.node_id == self.store.selected_node_id else (),
                )
                rendered_activities += 1

        visit(None, 1)
        self.activity_empty_var.set(
            "" if rendered_activities else
            ("No matching activities" if self.progress_file else "No Progress workbook loaded")
        )
        self.activity_page_var.set(
            f"Rows {page.start}-{page.end} of {page.total} | Page {page.number}/{page.pages}"
        )

    def _boq_values(self, key: str) -> tuple[str, ...]:
        row = self.store.boq_by_id[key]
        return (
            "☑" if key in self.store.selected_boq_ids else "☐",
            row.wbs2,
            row.wbs3,
            row.wbs4,
            row.description,
            f"{row.amount:,.2f}",
            f"{self.store.boq_allocated_amount(key):,.2f}",
            f"{self.store.boq_remaining_percent(key):.0f}%",
            self.store.boq_status(key).value,
            self.store.mapped_to_compact_text(key),
        )

    def _render_boq(self) -> None:
        self.boq_tree.delete(*self.boq_tree.get_children())
        page = self.store.boq_page_data()
        for key in page.ids:
            self.boq_tree.insert("", "end", iid=key, values=self._boq_values(key))
        self.boq_empty_var.set("" if page.total else ("No matching BOQ items" if self.boq_sheet else "No BOQ worksheet loaded"))
        self.boq_page_var.set(f"Rows {page.start}-{page.end} of {page.total} | Page {page.number}/{page.pages}")
        self._update_boq_selection_status()

    def _refresh_changed_rows(self, change) -> None:
        """Update visible rows only; the memory store remains the source of truth."""
        for key in change.boq_keys:
            if self.boq_tree.exists(key):
                self.boq_tree.item(key, values=self._boq_values(key))
        for activity_id in change.activity_ids:
            if self.activity_tree.exists(activity_id):
                self.activity_tree.set(
                    activity_id,
                    "amount",
                    f"{self.store.activity_amount(activity_id):,.2f}",
                )
        self._update_summary()
        self._update_boq_selection_status()

    def _activity_click(self, event) -> str | None:
        if self.activity_tree.identify_region(event.x, event.y) != "cell":
            return None
        item = self.activity_tree.identify_row(event.y)
        if item in self._wbs_header_paths:
            path = self._wbs_header_paths[item]
            self.store.select_wbs(path)
            self.store.selected_activity_ids.clear()
            self.store.toggle_wbs(self.store.wbs_path_key(path))
            self._render_activities()
            return "break"
        if self.activity_tree.identify_column(event.x) != "#1":
            return None
        if item and item in self.store.activities_by_id:
            previous = set(self.store.selected_activity_ids)
            self.store.toggle_activity(item)
            for activity_id in previous | self.store.selected_activity_ids:
                if self.activity_tree.exists(activity_id):
                    self.activity_tree.set(activity_id, "check", "☑" if activity_id in self.store.selected_activity_ids else "☐")
        return "break"

    def _boq_click(self, event) -> str | None:
        if self.boq_tree.identify_region(event.x, event.y) != "cell" or self.boq_tree.identify_column(event.x) != "#1":
            return None
        item = self.boq_tree.identify_row(event.y)
        if not item:
            return "break"
        full_mapping = self.store.mapped_to_text(item)
        self.boq_mapping_detail_var.set(full_mapping or "Not mapped")
        previous = set(self.store.selected_boq_ids)
        shift = bool(event.state & 0x0001)
        additive = bool(event.state & 0x0004) or bool(event.state & 0x0008)
        if shift:
            self.store.select_boq_range(item)
        else:
            self.store.toggle_boq(item, additive=additive)
        for key in previous | self.store.selected_boq_ids:
            if self.boq_tree.exists(key):
                self.boq_tree.set(key, "check", "☑" if key in self.store.selected_boq_ids else "☐")
        self._update_boq_selection_status()
        return "break"

    def _update_boq_selection_status(self) -> None:
        self.boq_selection_var.set(
            f"Selected {len(self.store.selected_boq_ids):,} items | {self.store.selected_boq_amount:,.2f}"
        )

    def _select_boq_page(self) -> None:
        self.store.select_boq_page()
        self._render_boq()

    def _select_all_filtered_boq(self) -> None:
        self.store.select_all_filtered_boq()
        self._render_boq()

    def _clear_boq_selection(self) -> None:
        self.store.clear_boq_selection()
        self._render_boq()

    def _collapse_all_wbs(self) -> None:
        self.store.collapse_all_wbs()
        self._render_activities()

    def _expand_all_wbs(self) -> None:
        self.store.expand_all_wbs()
        self._render_activities()

    def _activity_prev(self) -> None:
        self.store.activity_page -= 1
        self._render_activities()

    def _activity_next(self) -> None:
        self.store.activity_page += 1
        self._render_activities()

    def _boq_prev(self) -> None:
        self.store.boq_page -= 1
        self._render_boq()

    def _boq_next(self) -> None:
        self.store.boq_page += 1
        self._render_boq()

    def _confirm_batch_mapping(self, action: str) -> bool:
        count = len(self.store.selected_boq_ids)
        if count < 50:
            return True
        activity_id = next(iter(self.store.selected_activity_ids), "")
        return messagebox.askyesno(
            "Amount Mapping",
            f"{action} {count:,} BOQ items for {activity_id}?\n"
            f"Selected amount: {self.store.selected_boq_amount:,.2f}",
        )

    def _map(self) -> None:
        try:
            if not self._confirm_batch_mapping("Map"):
                return
            change = self.store.map_selected(self.share_var.get())
            self._refresh_changed_rows(change)
            self._autosave_session()
            self._notify("Mapping saved")
        except ValueError as exc:
            messagebox.showwarning("Amount Mapping", str(exc))

    def _unmap(self) -> None:
        try:
            if not self._confirm_batch_mapping("Unmap"):
                return
            change = self.store.unmap_selected()
            self._refresh_changed_rows(change)
            self._autosave_session()
            self._notify("Mapping removed")
        except ValueError as exc:
            messagebox.showwarning("Amount Mapping", str(exc))

    def _undo(self) -> None:
        change = self.store.undo()
        if change is None:
            messagebox.showinfo("Amount Mapping", "Nothing to undo.")
            return
        self._refresh_changed_rows(change)
        self._autosave_session()
        self._notify("Undo complete")

    def _clear_all(self) -> None:
        if not messagebox.askyesno(
            "Amount Mapping",
            "Clear every mapping? You can use Undo immediately after this command.",
        ):
            return
        try:
            change = self.store.clear_all()
            self._refresh_changed_rows(change)
            self._autosave_session()
        except ValueError as exc:
            messagebox.showinfo("Amount Mapping", str(exc))

    def _update_summary(self) -> None:
        total = self.store.total_amount
        mapped = self.store.mapped_amount
        self.summary_var.set(
            f"Mapped {mapped:,.2f} / {total:,.2f} | Remaining {self.store.remaining_amount:,.2f} | "
            f"Items {self.store.mapped_item_count}/{len(self.store.boq_order)}"
        )
        ready = bool(self.progress_file and self.boq_file and self.boq_sheet and self.store.boq_order)
        # Export now lives in the application shell. Keep compatibility with
        # embedded/legacy layouts that still provide a local export button.
        export_button = getattr(self, "export_button", None)
        if export_button is not None:
            export_button.configure(state="normal" if ready else "disabled")

    def _default_session_path(self) -> Path | None:
        """Return the automatic working-session path for the current inputs."""
        if not self.progress_file or not self.boq_file or not self.boq_sheet:
            return None
        identity = "|".join((
            str(self.progress_file.expanduser().resolve()).lower(),
            str(self.boq_file.expanduser().resolve()).lower(),
            self.boq_sheet.strip().lower(),
        ))
        import hashlib
        key = hashlib.sha1(identity.encode("utf-8")).hexdigest()[:12]
        safe_stem = "".join(
            character if character.isalnum() or character in ("-", "_") else "_"
            for character in self.progress_file.stem
        ).strip("_") or "project"
        return user_data_dir() / "sessions" / f"{safe_stem}-{key}.progressstudio"

    def _write_session(self, path: Path, show_message: bool) -> Path:
        if not self.progress_file or not self.boq_file or not self.boq_sheet:
            raise ValueError("Load the Progress workbook and BOQ worksheet first.")
        session = self.session_repository.create(
            self.progress_file,
            self.boq_file,
            self.boq_sheet,
            self.store.allocation_records(),
            self.store.supplemental_activities(),
            self.store.supplemental_wbs_nodes,
            list(self.store.working_tree_nodes()),
        )
        saved = self.session_repository.save(path, session)
        self.session_file = saved
        self.recent_repository.remember(saved)
        self.session_status_var.set(f"Project: {saved.name} (saved)")
        if show_message:
            messagebox.showinfo("Amount Mapping", f"Project saved:\n{saved}")
        return saved

    def _save_session(self) -> None:
        initial = self.session_file or self._default_session_path()
        if initial is None:
            messagebox.showwarning("Amount Mapping", "Load a Progress workbook first.")
            return
        selected = filedialog.asksaveasfilename(
            title="Save Progress Studio project",
            defaultextension=".json",
            initialfile=initial.name,
            initialdir=str(initial.parent),
            filetypes=[("Progress Studio project", "*.progressstudio")],
        )
        if not selected:
            return
        try:
            self._write_session(Path(selected), show_message=True)
        except Exception as exc:
            messagebox.showerror("Amount Mapping", str(exc))

    def _autosave_session(self) -> None:
        if self.session_file is None:
            self.session_file = self._default_session_path()
        if self.session_file is None:
            self.session_status_var.set("Project: waiting for both workbooks")
            return
        try:
            self._write_session(self.session_file, show_message=False)
            self.session_status_var.set("Auto-saved")
        except Exception as exc:
            self.session_status_var.set("Project: auto-save failed")
            messagebox.showerror("Amount Mapping", f"Auto-save failed:\n{exc}")

    def _choose_recent_session(self, paths: list[Path]) -> Path | None:
        dialog = tk.Toplevel(self)
        dialog.title("Recent Projects")
        dialog.transient(self.winfo_toplevel())
        dialog.grab_set()
        dialog.geometry("680x260")
        selected: list[Path] = []
        ttk.Label(dialog, text="Select a project to continue:").pack(anchor="w", padx=12, pady=(12, 6))
        box = tk.Listbox(dialog, height=8)
        box.pack(fill="both", expand=True, padx=12)
        for path in paths:
            box.insert("end", str(path))
        box.selection_set(0)

        buttons = ttk.Frame(dialog)
        buttons.pack(fill="x", padx=12, pady=12)
        def open_selected() -> None:
            indexes = box.curselection()
            if indexes:
                selected.append(paths[indexes[0]])
                dialog.destroy()
        ttk.Button(buttons, text="Open", command=open_selected).pack(side="right")
        ttk.Button(buttons, text="Cancel", command=dialog.destroy).pack(side="right", padx=(0, 8))
        box.bind("<Double-1>", lambda _event: open_selected())
        dialog.wait_window()
        return selected[0] if selected else None

    def _load_recent_session(self) -> None:
        paths = self.recent_repository.list()
        if not paths:
            messagebox.showinfo("Amount Mapping", "No recent projects were found.")
            return
        selected = self._choose_recent_session(paths)
        if selected:
            self._restore_session(selected)

    def _load_session(self) -> None:
        selected = filedialog.askopenfilename(
            title="Open Progress Studio project",
            filetypes=[("Progress Studio project", "*.progressstudio"), ("All files", "*.*")],
        )
        if selected:
            self._restore_session(Path(selected))

    def _resolve_session_workbook(
        self, saved: WorkbookFingerprint, workbook_label: str
    ) -> Path:
        try:
            return self.session_repository.validate_workbook(saved)
        except SessionValidationError as original_error:
            browse = messagebox.askyesno(
                "Relink workbook",
                f"{workbook_label} workbook cannot be verified at its saved location.\n\n"
                f"Saved file: {saved.filename}\n\n"
                "Browse for the moved or renamed workbook?\n"
                "Only an identical workbook will be accepted.",
            )
            if not browse:
                raise original_error
            selected = filedialog.askopenfilename(
                title=f"Relink {workbook_label} workbook",
                initialfile=saved.filename,
                filetypes=[("Excel workbook", "*.xlsx *.xlsm"), ("All files", "*.*")],
            )
            if not selected:
                raise SessionValidationError(
                    f"Relink cancelled for {workbook_label} workbook: {saved.filename}"
                )
            return self.session_repository.validate_workbook(saved, Path(selected))

    def _restore_session(self, session_path: Path) -> None:
        try:
            self._busy(True)
            session = self.session_repository.load(session_path)
            progress_file = self._resolve_session_workbook(session.progress, "Progress")
            boq_file = self._resolve_session_workbook(session.boq, "BOQ")

            activities = self.service.read_activities(progress_file)
            sheet_names = self.service.list_boq_sheets(boq_file)
            if session.boq_sheet not in sheet_names:
                raise SessionValidationError(
                    f"BOQ worksheet was not found: {session.boq_sheet}"
                )
            boq_rows = self.service.read_boq(boq_file, session.boq_sheet)

            self.store.load_activities(activities)
            self.store.supplemental_wbs_nodes = list(session.supplemental_wbs)
            for row in session.supplemental_activities:
                self.store.activities_by_id[row.activity_id] = row
                self.store.activity_order.append(row.activity_id)
            self.store._rebuild_working_tree()
            if session.working_tree_nodes:
                self.store.restore_working_tree(session.working_tree_nodes)
            self.store.load_boq(boq_rows)
            self.store.restore_allocations(list(session.allocations))
            self.progress_file = progress_file
            self.boq_file = boq_file
            self.boq_sheet = session.boq_sheet
            self.progress_path_var.set(str(progress_file))
            self.boq_path_var.set(str(boq_file))
            self.boq_sheet_combo.configure(values=sheet_names, state="readonly")
            self.boq_sheet_var.set(session.boq_sheet)
            self.load_sheet_button.configure(state="normal")
            self.session_file = Path(session_path).resolve()
            self.recent_repository.remember(self.session_file)
            self.session_status_var.set(f"Project: {self.session_file.name} (loaded)")
            self._refresh_boq_filter_values()
            self._render_activities()
            self._render_boq()
            self._update_summary()
            self._update_input_status()
            self._set_inputs_collapsed(True)
        except (SessionValidationError, ValueError, OSError) as exc:
            messagebox.showerror("Amount Mapping", str(exc))
        finally:
            self._busy(False)

    def _export(self) -> None:
        if not self.progress_file:
            messagebox.showwarning("Amount Mapping", "Load a Progress workbook first.")
            return
        try:
            validation = self.export_service.validate(self.store)
        except ValueError as exc:
            messagebox.showerror("Export mapped workbook", str(exc))
            return

        summary = (
            f"Allocated: {validation.allocated_amount:,.2f} / {validation.total_boq_amount:,.2f} "
            f"({validation.allocated_percent:.2f}%)\n"
            f"Full BOQ items: {validation.full_boq_count}\n"
            f"Partial BOQ items: {validation.partial_boq_count}\n"
            f"Unmapped BOQ items: {validation.unmapped_boq_count}"
        )
        if not validation.is_complete:
            proceed = messagebox.askyesno(
                "Export partial mapping?",
                summary + "\n\nThe mapping is incomplete. Export the partial workbook anyway?",
            )
            if not proceed:
                return

        initial = self.progress_file.with_name(self.progress_file.stem + "_mapped.xlsx")
        selected = filedialog.asksaveasfilename(
            title="Export mapped Progress workbook",
            defaultextension=".xlsx",
            initialfile=initial.name,
            initialdir=str(initial.parent),
            filetypes=[("Excel workbook", "*.xlsx")],
            confirmoverwrite=True,
        )
        if not selected:
            return
        dialog = GenerationProgressDialog(self)

        def report(step: str, message: str, complete: bool) -> None:
            if complete:
                dialog.complete_step(step, message)
            else:
                dialog.update_step(step, message)

        try:
            self._busy(True, "Generating workbook...")
            result = self.export_service.export(
                self.progress_file,
                Path(selected),
                self.store,
                overwrite=True,
                progress_callback=report,
            )
            dialog.complete_step("finalize", "Workbook generated successfully.")
            self._notify("Workbook exported")
            messagebox.showinfo(
                "Export complete",
                f"Mapped workbook created:\n{result.output_file}\n\n"
                f"Amount rows updated: {result.amount_rows_updated}\n"
                f"Mapping rows written: {result.mapping_rows_written}\n\n{summary}",
            )
        except Exception as exc:
            dialog.fail(f"Generation failed: {exc}")
            messagebox.showerror("Export mapped workbook", str(exc))
        finally:
            dialog.close()
            self._busy(False)
