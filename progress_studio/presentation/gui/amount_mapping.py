from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from progress_studio.domain.mapping_session import WorkbookFingerprint
from progress_studio.services.boq_mapping_service import BOQMappingService
from progress_studio.services.mapping_store import MappingStore
from progress_studio.services.workbook_export_service import WorkbookExportService
from progress_studio.infrastructure.layout_preferences import (
    LayoutPreferences,
    LayoutPreferencesRepository,
)
from progress_studio.infrastructure.platform_paths import user_data_dir
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
        self.session_status_var = tk.StringVar(value="Session: not saved")
        self.layout_repository = LayoutPreferencesRepository()
        self.layout_preferences = self.layout_repository.load()
        self.inputs_collapsed = self.layout_preferences.mapping_inputs_collapsed
        self._build()
        self.after_idle(self._restore_mapping_sash)
        self.bind("<Destroy>", self._on_destroy, add="+")

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
        ttk.Label(inputs_header, textvariable=self.input_status_var, foreground="#52606d").pack(
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

        toolbar = ttk.Frame(self)
        toolbar.pack(fill="x", pady=(0, 6))
        self.export_button = ttk.Button(
            toolbar, text="Export...", command=self._export, state="disabled"
        )
        self.export_button.pack(side="left")
        ttk.Button(toolbar, text="Load Session...", command=self._load_session).pack(side="left", padx=(6, 0))
        self.save_session_button = ttk.Button(
            toolbar, text="Save Session...", command=self._save_session, state="disabled"
        )
        self.save_session_button.pack(side="left", padx=(6, 0))
        self.recent_session_button = ttk.Button(
            toolbar, text="Recent...", command=self._load_recent_session
        )
        self.recent_session_button.pack(side="left", padx=(6, 0))
        ttk.Label(toolbar, textvariable=self.session_status_var, foreground="#52606d").pack(side="left", padx=(10, 0))
        ttk.Label(toolbar, textvariable=self.summary_var).pack(side="right")

        self.body = ttk.Panedwindow(self, orient="horizontal")
        self.body.pack(fill="both", expand=True)
        left = ttk.Frame(self.body, padding=(0, 0, 6, 0))
        right = ttk.Frame(self.body, padding=(6, 0, 0, 0))
        self.body.add(left, weight=1)
        self.body.add(right, weight=1)

        ttk.Label(left, text="Progress Activities", font=("Segoe UI", 10, "bold")).pack(anchor="w")
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
        self.activity_tree.bind("<Button-1>", self._activity_click)
        self._build_pager(left, self.activity_page_var, self._activity_prev, self._activity_next)

        ttk.Label(right, text="BOQ Items", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        filter_row = ttk.Frame(right)
        filter_row.pack(fill="x", pady=(4, 4))
        ttk.Label(filter_row, text="WBS-2").pack(side="left")
        self.boq_wbs2_combo = ttk.Combobox(filter_row, textvariable=self.boq_wbs2_filter, state="readonly", width=16, values=("All",))
        self.boq_wbs2_combo.pack(side="left", padx=(5, 8))
        self.boq_wbs2_combo.bind("<<ComboboxSelected>>", self._on_wbs2_filter)
        ttk.Label(filter_row, text="WBS-3").pack(side="left")
        self.boq_wbs3_combo = ttk.Combobox(filter_row, textvariable=self.boq_wbs3_filter, state="readonly", width=16, values=("All",))
        self.boq_wbs3_combo.pack(side="left", padx=(5, 0))
        self.boq_wbs3_combo.bind("<<ComboboxSelected>>", self._on_wbs3_filter)
        self._build_search(right, self.boq_filter, self._apply_boq_filter)
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
            (38, 75, 85, 85, 310, 92, 92, 88, 72, 95),
        )
        self.boq_tree.bind("<Button-1>", self._boq_click)
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
            foreground="#52606d",
        ).pack(side="left", padx=(12, 0))

        self._set_inputs_collapsed(self.inputs_collapsed, persist=False)

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
            mapping_sash=sash,
        )
        try:
            self.layout_repository.save(self.layout_preferences)
        except OSError:
            pass

    def _on_destroy(self, event) -> None:
        if event.widget is self:
            self._save_layout_preferences()

    @staticmethod
    def _build_search(parent, variable: tk.StringVar, command) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(4, 6))
        entry = ttk.Entry(row, textvariable=variable)
        entry.pack(side="left", fill="x", expand=True)
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

    def _busy(self, active: bool) -> None:
        top = self.winfo_toplevel()
        top.configure(cursor="watch" if active else "")
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

    def _apply_activity_filter(self) -> None:
        self.store.activity_query = self.activity_filter.get()
        self.store.activity_page = 1
        self._render_activities()

    def _apply_boq_filter(self) -> None:
        self.store.boq_query = self.boq_filter.get()
        self.store.boq_page = 1
        self._render_boq()

    def _render_activities(self) -> None:
        self.activity_tree.delete(*self.activity_tree.get_children())
        page = self.store.activity_page_data()
        previous_path: tuple[tuple[str, str], ...] = ()
        header_counter = 0

        for activity_id in page.ids:
            row = self.store.activities_by_id[activity_id]

            # Repeat the full hierarchy at the start of every page and only
            # insert the levels that change between consecutive activities.
            common = 0
            for old_level, new_level in zip(previous_path, row.wbs_path):
                if old_level != new_level:
                    break
                common += 1
            for level, (code, name) in enumerate(row.wbs_path[common:], start=common + 1):
                header_counter += 1
                indent = "    " * (level - 1)
                header_id = f"__wbs__{page.number}_{header_counter}_{code}"
                self.activity_tree.insert(
                    "",
                    "end",
                    iid=header_id,
                    values=("", f"{indent}▾ {code}", name, ""),
                    tags=(f"wbs_level_{min(level, 3)}", "wbs_header"),
                )

            check = "☑" if activity_id in self.store.selected_activity_ids else "☐"
            activity_indent = "    " * len(row.wbs_path)
            self.activity_tree.insert(
                "",
                "end",
                iid=activity_id,
                values=(
                    check,
                    f"{activity_indent}{row.activity_id}",
                    row.description,
                    f"{self.store.activity_amount(activity_id):,.2f}",
                ),
            )
            previous_path = row.wbs_path

        self.activity_page_var.set(f"Rows {page.start}-{page.end} of {page.total} | Page {page.number}/{page.pages}")

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
            self.store.mapped_to_text(key),
        )

    def _render_boq(self) -> None:
        self.boq_tree.delete(*self.boq_tree.get_children())
        page = self.store.boq_page_data()
        for key in page.ids:
            self.boq_tree.insert("", "end", iid=key, values=self._boq_values(key))
        self.boq_page_var.set(f"Rows {page.start}-{page.end} of {page.total} | Page {page.number}/{page.pages}")

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

    def _activity_click(self, event) -> str | None:
        if self.activity_tree.identify_region(event.x, event.y) != "cell" or self.activity_tree.identify_column(event.x) != "#1":
            return None
        item = self.activity_tree.identify_row(event.y)
        if item and item in self.store.activities_by_id:
            previous = set(self.store.selected_activity_ids)
            self.store.toggle_activity(item)
            for activity_id in previous | self.store.selected_activity_ids:
                if self.activity_tree.exists(activity_id):
                    row = self.store.activities_by_id[activity_id]
                    self.activity_tree.set(activity_id, "check", "☑" if activity_id in self.store.selected_activity_ids else "☐")
        return "break"

    def _boq_click(self, event) -> str | None:
        if self.boq_tree.identify_region(event.x, event.y) != "cell" or self.boq_tree.identify_column(event.x) != "#1":
            return None
        item = self.boq_tree.identify_row(event.y)
        if item:
            self.store.toggle_boq(item)
            self.boq_tree.set(item, "check", "☑" if item in self.store.selected_boq_ids else "☐")
        return "break"

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

    def _map(self) -> None:
        try:
            change = self.store.map_selected(self.share_var.get())
            self._refresh_changed_rows(change)
            self._autosave_session()
        except ValueError as exc:
            messagebox.showwarning("Amount Mapping", str(exc))

    def _unmap(self) -> None:
        try:
            change = self.store.unmap_selected()
            self._refresh_changed_rows(change)
            self._autosave_session()
        except ValueError as exc:
            messagebox.showwarning("Amount Mapping", str(exc))

    def _undo(self) -> None:
        change = self.store.undo()
        if change is None:
            messagebox.showinfo("Amount Mapping", "Nothing to undo.")
            return
        self._refresh_changed_rows(change)
        self._autosave_session()

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
        self.export_button.configure(state="normal" if ready else "disabled")
        self.save_session_button.configure(state="normal" if ready else "disabled")

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
        return user_data_dir() / "sessions" / f"{safe_stem}-{key}.mapping.json"

    def _write_session(self, path: Path, show_message: bool) -> Path:
        if not self.progress_file or not self.boq_file or not self.boq_sheet:
            raise ValueError("Load the Progress workbook and BOQ worksheet first.")
        session = self.session_repository.create(
            self.progress_file,
            self.boq_file,
            self.boq_sheet,
            self.store.allocation_records(),
        )
        saved = self.session_repository.save(path, session)
        self.session_file = saved
        self.recent_repository.remember(saved)
        self.session_status_var.set(f"Session: {saved.name} (saved)")
        if show_message:
            messagebox.showinfo("Amount Mapping", f"Mapping session saved:\n{saved}")
        return saved

    def _save_session(self) -> None:
        initial = self.session_file or self._default_session_path()
        if initial is None:
            messagebox.showwarning("Amount Mapping", "Load a Progress workbook first.")
            return
        selected = filedialog.asksaveasfilename(
            title="Save mapping session",
            defaultextension=".json",
            initialfile=initial.name,
            initialdir=str(initial.parent),
            filetypes=[("Progress Studio mapping session", "*.json")],
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
            self.session_status_var.set("Session: waiting for both workbooks")
            return
        try:
            self._write_session(self.session_file, show_message=False)
            self.session_status_var.set(f"Session: {self.session_file.name} (auto-saved)")
        except Exception as exc:
            self.session_status_var.set("Session: auto-save failed")
            messagebox.showerror("Amount Mapping", f"Auto-save failed:\n{exc}")

    def _choose_recent_session(self, paths: list[Path]) -> Path | None:
        dialog = tk.Toplevel(self)
        dialog.title("Recent Mapping Sessions")
        dialog.transient(self.winfo_toplevel())
        dialog.grab_set()
        dialog.geometry("680x260")
        selected: list[Path] = []
        ttk.Label(dialog, text="Select a mapping session to continue:").pack(anchor="w", padx=12, pady=(12, 6))
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
            messagebox.showinfo("Amount Mapping", "No recent mapping sessions were found.")
            return
        selected = self._choose_recent_session(paths)
        if selected:
            self._restore_session(selected)

    def _load_session(self) -> None:
        selected = filedialog.askopenfilename(
            title="Load mapping session",
            filetypes=[("Progress Studio mapping session", "*.json"), ("All files", "*.*")],
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
            self.session_status_var.set(f"Session: {self.session_file.name} (loaded)")
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
        try:
            self._busy(True)
            result = self.export_service.export(
                self.progress_file, Path(selected), self.store, overwrite=True
            )
            messagebox.showinfo(
                "Export complete",
                f"Mapped workbook created:\n{result.output_file}\n\n"
                f"Amount rows updated: {result.amount_rows_updated}\n"
                f"Mapping rows written: {result.mapping_rows_written}\n\n{summary}",
            )
        except Exception as exc:
            messagebox.showerror("Export mapped workbook", str(exc))
        finally:
            self._busy(False)
