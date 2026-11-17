"""BOQ reports: review sources, select catalog rows, preview and publish."""
from concurrent.futures import ThreadPoolExecutor
from dataclasses import replace
from threading import Event
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter.scrolledtext import ScrolledText

from progress_studio.presentation.gui.strings import tr
from progress_studio.services.boq_payment_export import BOQPaymentExportService
from progress_studio.services.boq_word_report import BOQReportCatalog, display_number, export_boq_reports
from progress_studio.services.payment_plan_export import ExportCancelled


def scrolling_tree(frame, **options):
    box = ttk.Frame(frame)
    box.pack(fill='both', expand=True, pady=5)
    box.rowconfigure(0, weight=1)
    box.columnconfigure(0, weight=1)
    tree = ttk.Treeview(box, **options)
    tree.grid(row=0, column=0, sticky='nsew')
    for axis, row, column, sticky in [('y', 0, 1, 'ns'), ('x', 1, 0, 'ew')]:
        bar = ttk.Scrollbar(box, orient='vertical' if axis == 'y' else 'horizontal',
                            command=getattr(tree, axis + 'view'))
        bar.grid(row=row, column=column, sticky=sticky)
        tree.configure(**{axis + 'scrollcommand': bar.set})
    return tree


class BOQWordDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title(tr('bpr.title'))
        self.transient(parent.winfo_toplevel())
        self.minsize(min(680, self.winfo_screenwidth()-48), min(395, self.winfo_screenheight()-105))
        self.geometry(f'{min(940, self.winfo_screenwidth()-48)}x{min(620, self.winfo_screenheight()-105)}')
        self.executor = ThreadPoolExecutor(max_workers=1)
        self.cancel = Event()
        self.future = self.catalog = self.plan = None
        self.keys = set()
        self.stage = 0
        self.paths = [tk.StringVar(self) for _ in range(3)]
        self.search = tk.StringVar(self)
        self.scope = tk.StringVar(self)
        self.status = tk.StringVar(self)
        self.ack = tk.BooleanVar(self)
        self.cumulative = tk.BooleanVar(self)
        self.period = tk.StringVar(self)
        self.pages = []
        self.steps = []
        body = ttk.Frame(self, padding=12)
        body.pack(fill='both', expand=True)
        nav = ttk.Frame(body)
        nav.pack(fill='x')
        for label in ('bpr.sources', 'bpr.work', 'bpr.review'):
            step = ttk.Label(nav, text=tr(label), padding=(5, 3))
            step.pack(side='left')
            self.steps.append(step)
        ttk.Separator(body).pack(fill='x', pady=(3, 6))
        footer = ttk.Frame(body)
        footer.pack(side='bottom', fill='x')
        self.primary = ttk.Button(footer, style='Accent.TButton', command=self.forward)
        self.primary.pack(side='right')
        self.back = ttk.Button(footer, text=tr('pu.back'), command=lambda: self.show(self.stage-1))
        self.back.pack(side='right', padx=5)
        self.exit_button = ttk.Button(footer, command=self.close)
        self.exit_button.pack(side='left')
        ttk.Label(body, textvariable=self.status, wraplength=680).pack(side='bottom', anchor='w', pady=3)
        for build in (self.source_page, self.select_page, self.review_page):
            page = ttk.Frame(body)
            self.pages.append(page)
            build(page)
        self.protocol('WM_DELETE_WINDOW', self.close)
        self.show(0)

    def source_page(self, page):
        ttk.Label(page, text=tr('bp2.1')).pack(anchor='w', pady=6)
        fields = ttk.Frame(page)
        fields.pack(fill='x')
        fields.columnconfigure(1, weight=1)
        self.inputs = []
        for i, key in enumerate(('bp2.2', 'bp2.3', 'bp2.4')):
            ttk.Label(fields, text=tr(key)).grid(row=i, column=0, sticky='w')
            entry = ttk.Entry(fields, textvariable=self.paths[i])
            entry.grid(row=i, column=1, sticky='ew', padx=5, pady=4)
            button = ttk.Button(fields, text=tr('Browse...'), command=lambda i=i: self.pick(i))
            button.grid(row=i, column=2)
            self.inputs.extend((entry, button))
            self.paths[i].trace_add('write', lambda *_: self.sync())

    def pick(self, index):
        pattern = '*.progressstudio' if index == 0 else '*.xlsx'
        path = filedialog.askopenfilename(parent=self, filetypes=[(pattern, pattern)])
        if path:
            self.paths[index].set(path)

    def select_page(self, page):
        sources = ttk.Frame(page)
        sources.pack(fill='x', pady=(0, 4))
        self.loaded_source = ttk.Label(sources)
        self.loaded_source.pack(side='left', fill='x', expand=True)
        ttk.Button(sources, text=tr('pr2.problems'), command=self.diagnostics).pack(side='right')
        top = ttk.Frame(page)
        top.pack(fill='x')
        ttk.Label(top, text=tr('bpr.search')).pack(side='left')
        ttk.Entry(top, textvariable=self.search).pack(side='left', fill='x', expand=True, padx=5)
        ttk.Button(top, text=tr('pr2.whole'), command=lambda: self.choose({self.catalog.root.key})).pack(side='left')
        ttk.Button(top, text=tr('pr2.clear'), command=lambda: self.choose(set())).pack(side='left')
        self.tree = scrolling_tree(page, columns=('role',), selectmode='browse')
        self.tree.heading('#0', text=tr('pr2.name'))
        self.tree.heading('role', text=tr('pr2.state'))
        self.tree.column('#0', width=480)
        self.tree.column('role', width=135)
        self.tree.bind('<ButtonRelease-1>', self.click)
        self.tree.bind('<space>', self.toggle_focus)
        self.search.trace_add('write', lambda *_: self.fill_tree())
        tools = ttk.Frame(page)
        tools.pack(fill='x')
        ttk.Label(tools, textvariable=self.scope).pack(side='left')
        for label, action in [('pr2.restore', False), ('pr2.save', True)]:
            ttk.Button(tools, text=tr(label), command=lambda save=action: self.selection_file(save)).pack(side='right')

    def review_page(self, page):
        bar = ttk.Frame(page)
        bar.pack(fill='x')
        self.period_box = ttk.Combobox(bar, textvariable=self.period, state='readonly', width=9)
        self.period_box.pack(side='left')
        self.period_box.bind('<<ComboboxSelected>>', lambda _: self.preview())
        for key, value in [('pr2.period', False), ('pr2.cumulative', True)]:
            ttk.Radiobutton(bar, text=tr(key), variable=self.cumulative, value=value,
                            command=self.preview).pack(side='left', padx=4)
        self.row_mode = ttk.Combobox(bar, state='readonly', width=19,
                                     values=(tr('bpr.active_rows'), tr('bpr.all_rows')))
        self.row_mode.current(0)
        self.row_mode.pack(side='left', padx=6)
        self.row_mode.bind('<<ComboboxSelected>>', lambda _: self.preview())
        ttk.Button(bar, text=tr('pr2.problems'), command=self.diagnostics).pack(side='right')
        self.date_label = ttk.Label(page, wraplength=760)
        self.date_label.pack(anchor='w', pady=(6, 0))
        self.grid_view = scrolling_tree(page, columns=('name','state','pct','amount'), show='headings')
        for column, key, width in [('name','pr2.name',340), ('state','pr2.status',175),
                                   ('pct','pr2.percent',105), ('amount','pr2.money',145)]:
            self.grid_view.heading(column, text=tr(key))
            self.grid_view.column(column, width=width, anchor='e' if column in ('pct','amount') else 'w')
        self.grid_view.tag_configure('total', background='#e5f0ed')
        self.basis = ttk.Label(page, wraplength=740)
        self.basis.pack(anchor='w')
        self.check = ttk.Checkbutton(page, variable=self.ack, text=tr('bp2.5'), command=self.sync)

    def show(self, stage):
        if self.future:
            return
        self.stage = stage
        for i, page in enumerate(self.pages):
            page.pack_forget()
            self.steps[i].configure(style='Section.TLabel' if i == stage else 'Muted.TLabel')
        self.pages[stage].pack(fill='both', expand=True)
        self.status.set('')
        self.sync()

    def sync(self):
        busy = self.future is not None
        self.row_mode.configure(state='disabled' if busy else 'readonly')
        self.back.configure(state='normal' if self.stage and not busy else 'disabled')
        self.exit_button.configure(text=tr('bp2.14' if busy else 'bpr.close'))
        for widget in self.inputs:
            widget.configure(state='disabled' if busy else 'normal')
        if self.stage == 0:
            ready = all(v.get().strip() for v in self.paths[:2])
            label = 'bpr.next'
        elif self.stage == 1:
            ready = bool(self.keys)
            label = 'pu.to_review'
        else:
            ready = not self.catalog.review.revision_required or self.ack.get()
            label = 'pr3.export'
        self.primary.configure(text=tr(label), state='normal' if ready and not busy else 'disabled')

    def forward(self):
        if self.future:
            return
        if self.stage == 0:
            paths = tuple(v.get().strip() for v in self.paths)
            def review():
                return BOQReportCatalog(BOQPaymentExportService().review(*paths, cancelled=self.cancel.is_set))
            def loaded(catalog):
                if self.catalog is None or catalog.review.sources != self.catalog.review.sources:
                    self.keys.clear()
                    self.plan = None
                    self.ack.set(False)
                    self.search.set('')
                self.catalog = catalog
                session, payment = catalog.review.sources[:2]
                self.loaded_source.configure(text=f'{session[0].stem}  ·  {payment[0].name}')
                self.fill_tree()
                self.show(1)
            self.run(review, loaded, 'bp2.8')
        elif self.stage == 1:
            try:
                self.plan = self.catalog.select(self.keys)
            except ValueError as exc:
                self.status.set(str(exc))
                return
            self.period_box.configure(values=self.catalog.review.projection.periods)
            if self.period.get() not in self.catalog.review.projection.periods:
                self.period_box.current(0)
            self.preview()
            if self.catalog.review.revision_required:
                self.check.pack(anchor='w')
            else:
                self.check.pack_forget()
            self.show(2)
            if self.catalog.review.revision_required and not self.ack.get():
                self.status.set(tr('bpr.revision'))
        else:
            parent = filedialog.askdirectory(parent=self, title=tr('pr3.destination'))
            if not parent:
                return
            name = simpledialog.askstring(tr('bpr.title'), tr('pr3.folder'), parent=self,
                                          initialvalue=self.catalog.review.sources[0][0].stem + '_BOQ')
            if name:
                plan, ack = self.plan, self.ack.get()
                self.run(lambda: export_boq_reports(plan, parent, name, acknowledge_revision=ack,
                                                    cancelled=self.cancel.is_set),
                         lambda folder: self.status.set(tr('pr3.done', folder=folder)), 'bpr.progress')

    def fill_tree(self):
        if self.catalog is None:
            return
        tree, catalog = self.tree, self.catalog
        old = {k for k in catalog.by_key if tree.exists(k) and tree.item(k, 'open')}
        had_rows = bool(tree.get_children())
        tree.delete(*tree.get_children())
        query = self.search.get().strip().casefold()
        visible = set()
        if query:
            for r in catalog.rows:
                if query in (r.wbs+' '+r.name).casefold():
                    key = r.key
                    while key and key not in visible:
                        visible.add(key)
                        key = catalog.parents[key]
        for r in catalog.rows:
            if not query or r.key in visible:
                tree.insert(catalog.parents[r.key] or '', 'end', iid=r.key,
                            open=bool(query or r.key in old or (not had_rows and r is catalog.root)))
        self.mark_rows()

    def mark_rows(self):
        catalog = self.catalog
        for r in catalog.rows:
            if not self.tree.exists(r.key):
                continue
            parent = catalog.parents[r.key]
            while parent is not None and parent not in self.keys:
                parent = catalog.parents[parent]
            checked = r.key in self.keys
            role = tr('pr2.selected') if checked else ''
            if parent is not None:
                owner = catalog.by_key[parent]
                role = tr('pr2.included' if checked else 'pr2.represented', parent=owner.wbs or owner.name)
            self.tree.item(r.key, text=('☑  ' if checked else '☐  ') + r.wbs+'  '+r.name, values=(role,))
        leaves = {k for key in self.keys for k in catalog.members[key]}
        self.scope.set(tr('bpr.scope', rows=len(self.keys), count=len(leaves), all=len(catalog.entries)))
        self.sync()

    def choose(self, keys):
        self.keys = set(keys)
        self.plan = None
        self.mark_rows()

    def click(self, event):
        tree = self.tree
        key = tree.identify_row(event.y)
        if key and tree.identify_region(event.x, event.y) in ('cell','tree') and 'indicator' not in tree.identify_element(event.x, event.y):
            self.choose(self.keys ^ {key})

    def toggle_focus(self, _event):
        key = self.tree.focus()
        if key:
            self.choose(self.keys ^ {key})
        return 'break'

    def selection_file(self, save):
        ask = filedialog.asksaveasfilename if save else filedialog.askopenfilename
        path = ask(parent=self, defaultextension='.json', filetypes=[(tr('bpr.file'), '*.json')])
        if path:
            try:
                if save:
                    self.catalog.save(path, self.keys)
                else:
                    self.choose(self.catalog.restore(path))
                self.status.set(tr('pr2.saved' if save else 'pr2.restored'))
            except (OSError, ValueError) as exc:
                messagebox.showerror(tr('bpr.title'), str(exc), parent=self)

    def preview(self):
        index = self.period_box.current()
        session, payment = self.catalog.review.sources[:2]
        self.plan = replace(self.plan, active_only=self.row_mode.current() == 0)
        heading = self.plan.period_title(index)
        self.date_label.configure(text=f'{session[0].stem}  ·  {payment[0].name}  ·  {heading}')
        self.grid_view.delete(*self.grid_view.get_children())
        page = self.plan.page(index, self.cumulative.get())
        if not page.details:
            self.grid_view.insert('', 'end', values=(tr('bpr.no_activity'),))
        for values in page.details:
            self.grid_view.insert('', 'end', values=values)
        for values in page.totals:
            self.grid_view.insert('', 'end', values=values, tags=('total',))
        if page.diagnostics:
            self.grid_view.insert('', 'end', values=(tr('bpr.omitted', count=len(page.diagnostics)),))
        self.basis.configure(text=self.scope.get()+'\n'+tr('pr2.basis',
            selected=display_number(self.plan.total.amount), project=display_number(self.catalog.root.amount),
            coverage=display_number(self.plan.coverage, True)))

    def diagnostics(self):
        review = self.catalog.review
        window = tk.Toplevel(self)
        window.title(tr('pr2.problems'))
        window.transient(self)
        text = ScrolledText(window, width=80, height=20, wrap='word')
        text.pack(fill='both', expand=True)
        if self.stage == 2 and self.plan is not None:
            index = self.period_box.current()
            notes = self.plan.page(index, self.cumulative.get()).diagnostics
            if notes:
                text.insert('end', self.plan.period_title(index) + '\n' + '\n'.join(notes) + '\n\n')
        for path, revision in review.sources:
            text.insert('end', f'{path.name}\n{revision}\n\n')
        for issue in review.projection.diagnostics:
            text.insert('end', f'{issue.code} · {issue.identity}\n{issue.detail}\n\n')
        text.configure(state='disabled')

    def run(self, work, success, label):
        self.cancel.clear()
        self.future = self.executor.submit(work)
        self.status.set(tr(label))
        self.sync()
        def poll():
            if not self.future.done():
                self.after(80, poll)
                return
            finished, self.future = self.future, None
            try:
                value = finished.result()
                if label == 'bp2.8' and self.cancel.is_set():
                    raise InterruptedError()
                success(value)
            except (ExportCancelled, InterruptedError):
                self.status.set(tr('pr3.cancelled'))
            except Exception as exc:
                self.status.set(str(exc))
            self.sync()
        self.after(80, poll)

    def close(self):
        if self.future:
            self.cancel.set()
            self.status.set(tr('bp2.13'))
            return
        self.executor.shutdown(wait=False)
        self.destroy()
