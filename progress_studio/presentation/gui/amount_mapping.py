from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from progress_studio.services.boq_mapping_service import BOQMappingService
from progress_studio.services.mapping_store import MappingStore


class AmountMappingFrame(ttk.Frame):
    ACTIVITY_PAGE_SIZE = 200
    BOQ_PAGE_SIZE = 300

    def __init__(self, master, service: BOQMappingService | None = None) -> None:
        super().__init__(master, padding=8)
        self.service = service or BOQMappingService()
        self.store = MappingStore(self.ACTIVITY_PAGE_SIZE, self.BOQ_PAGE_SIZE)
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
        self.activity_page_var = tk.StringVar(value="Rows 0-0 of 0")
        self.boq_page_var = tk.StringVar(value="Rows 0-0 of 0")
        self._build()

    def _build(self) -> None:
        inputs = ttk.LabelFrame(self, text="Mapping Inputs", padding=8)
        inputs.pack(fill="x", pady=(0, 8))
        inputs.columnconfigure(1, weight=1)

        ttk.Label(inputs, text="Progress workbook").grid(row=0, column=0, sticky="w")
        ttk.Entry(inputs, textvariable=self.progress_path_var, state="readonly").grid(
            row=0, column=1, sticky="ew", padx=8
        )
        ttk.Button(inputs, text="Load Progress...", command=self._browse_progress).grid(row=0, column=2)

        ttk.Label(inputs, text="BOQ workbook").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(inputs, textvariable=self.boq_path_var, state="readonly").grid(
            row=1, column=1, sticky="ew", padx=8, pady=(6, 0)
        )
        ttk.Button(inputs, text="Load BOQ...", command=self._browse_boq).grid(row=1, column=2, pady=(6, 0))

        ttk.Label(inputs, text="BOQ worksheet").grid(row=2, column=0, sticky="w", pady=(6, 0))
        self.boq_sheet_combo = ttk.Combobox(
            inputs, textvariable=self.boq_sheet_var, state="disabled", values=()
        )
        self.boq_sheet_combo.grid(row=2, column=1, sticky="ew", padx=8, pady=(6, 0))
        self.load_sheet_button = ttk.Button(
            inputs, text="Load selected sheet", command=self._load_selected_boq_sheet, state="disabled"
        )
        self.load_sheet_button.grid(row=2, column=2, pady=(6, 0))

        ttk.Label(inputs, textvariable=self.input_status_var, foreground="#52606d").grid(
            row=3, column=0, columnspan=3, sticky="w", pady=(7, 0)
        )

        toolbar = ttk.Frame(self)
        toolbar.pack(fill="x", pady=(0, 8))
        self.export_button = ttk.Button(
            toolbar, text="Export mapped workbook...", command=self._export, state="disabled"
        )
        self.export_button.pack(side="left")
        ttk.Label(toolbar, textvariable=self.summary_var).pack(side="right")

        body = ttk.Panedwindow(self, orient="horizontal")
        body.pack(fill="both", expand=True)
        left = ttk.Frame(body, padding=(0, 0, 6, 0))
        right = ttk.Frame(body, padding=(6, 0, 0, 0))
        body.add(left, weight=1)
        body.add(right, weight=1)

        ttk.Label(left, text="Progress Activities", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self._build_search(left, self.activity_filter, self._apply_activity_filter)
        self.activity_tree = self._tree(
            left,
            ("check", "id", "description", "amount"),
            ("✓", "Activity ID / WBS", "Description", "Amount"),
            (38, 125, 330, 95),
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
        self.boq_wbs2_combo = ttk.Combobox(filter_row, textvariable=self.boq_wbs2_filter, state="readonly", width=20, values=("All",))
        self.boq_wbs2_combo.pack(side="left", padx=(5, 10))
        self.boq_wbs2_combo.bind("<<ComboboxSelected>>", self._on_wbs2_filter)
        ttk.Label(filter_row, text="WBS-3").pack(side="left")
        self.boq_wbs3_combo = ttk.Combobox(filter_row, textvariable=self.boq_wbs3_filter, state="readonly", width=20, values=("All",))
        self.boq_wbs3_combo.pack(side="left", padx=(5, 0))
        self.boq_wbs3_combo.bind("<<ComboboxSelected>>", self._on_wbs3_filter)
        self._build_search(right, self.boq_filter, self._apply_boq_filter)
        self.boq_tree = self._tree(
            right,
            ("check", "wbs2", "wbs3", "wbs4", "description", "amount", "mapped"),
            ("✓", "WBS-2", "WBS-3", "WBS-4", "Description", "Amount", "Mapped To"),
            (38, 95, 105, 105, 235, 95, 90),
        )
        self.boq_tree.bind("<Button-1>", self._boq_click)
        self._build_pager(right, self.boq_page_var, self._boq_prev, self._boq_next)

        actions = ttk.Frame(self)
        actions.pack(fill="x", pady=(8, 0))
        ttk.Button(actions, text="Map selected", command=self._map).pack(side="left")
        ttk.Button(actions, text="Undo", command=self._undo).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Unmap selected", command=self._unmap).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Clear all", command=self._clear).pack(side="left", padx=(8, 0))

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
        tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="none")
        y = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=y.set)
        for name, heading, width in zip(columns, headings, widths):
            tree.heading(name, text=heading)
            tree.column(name, width=width, minwidth=35, stretch=name not in {"check", "amount"})
        tree.pack(side="left", fill="both", expand=True)
        y.pack(side="right", fill="y")
        return tree

    def _browse_progress(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select Progress workbook",
            filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")],
        )
        if selected:
            self.set_progress_workbook(Path(selected))

    def set_progress_workbook(self, path: Path) -> None:
        self.progress_file = Path(path)
        try:
            self._busy(True)
            rows = self.service.read_activities(self.progress_file)
            self.store.load_activities(rows)
            self.progress_path_var.set(str(self.progress_file))
            self._render_activities()
            self._update_summary()
            self._update_input_status()
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
            self._busy(True)
            rows = self.service.read_boq(self.boq_file, sheet_name)
            self.boq_sheet = sheet_name
            self.store.load_boq(rows)
            self._refresh_boq_filter_values()
            self._render_boq()
            self._render_activities()
            self._update_summary()
            self._update_input_status()
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

    def _render_boq(self) -> None:
        self.boq_tree.delete(*self.boq_tree.get_children())
        page = self.store.boq_page_data()
        for key in page.ids:
            row = self.store.boq_by_id[key]
            check = "☑" if key in self.store.selected_boq_ids else "☐"
            self.boq_tree.insert("", "end", iid=key, values=(
                check, row.wbs2, row.wbs3, row.wbs4, row.description,
                f"{row.amount:,.2f}", self.store.assignments.get(key, ""),
            ))
        self.boq_page_var.set(f"Rows {page.start}-{page.end} of {page.total} | Page {page.number}/{page.pages}")

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
            self.store.map_selected()
            self._render_boq(); self._render_activities(); self._update_summary()
        except ValueError as exc:
            messagebox.showwarning("Amount Mapping", str(exc))

    def _unmap(self) -> None:
        try:
            self.store.unmap_selected()
            self._render_boq(); self._render_activities(); self._update_summary()
        except ValueError as exc:
            messagebox.showwarning("Amount Mapping", str(exc))

    def _undo(self) -> None:
        if not self.store.undo():
            messagebox.showinfo("Amount Mapping", "Nothing to undo.")
            return
        self._render_boq(); self._render_activities(); self._update_summary()

    def _clear(self) -> None:
        if self.store.assignments and not messagebox.askyesno("Amount Mapping", "Clear all mappings?"):
            return
        self.store.clear_all()
        self._render_boq(); self._render_activities(); self._update_summary()

    def _update_summary(self) -> None:
        total = self.store.total_amount
        mapped = self.store.mapped_amount
        self.summary_var.set(
            f"Mapped {mapped:,.2f} / {total:,.2f} | Remaining {total - mapped:,.2f} | "
            f"Items {len(self.store.assignments)}/{len(self.store.boq_order)}"
        )
        self.export_button.configure(state="normal" if self.progress_file and self.store.boq_order else "disabled")

    def _export(self) -> None:
        if not self.progress_file:
            messagebox.showwarning("Amount Mapping", "Load a Progress workbook first.")
            return
        initial = self.progress_file.with_name(self.progress_file.stem + "_mapped.xlsx")
        selected = filedialog.asksaveasfilename(
            title="Export mapped Progress workbook",
            defaultextension=".xlsx",
            initialfile=initial.name,
            initialdir=str(initial.parent),
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not selected:
            return
        try:
            output = self.service.export(
                self.progress_file,
                Path(selected),
                list(self.store.boq_by_id.values()),
                dict(self.store.assignments),
            )
            messagebox.showinfo("Amount Mapping", f"Mapped workbook created:\n{output}")
        except Exception as exc:
            messagebox.showerror("Amount Mapping", str(exc))
