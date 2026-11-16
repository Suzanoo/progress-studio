"""Select reporting rows and preview the two Payment Plan views."""
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter.scrolledtext import ScrolledText

from progress_studio.infrastructure.excel.payment_plan_source import read_payment_plan_source
from progress_studio.presentation.gui.payment_plan_selection import PlanSelection, SUFFIX
from progress_studio.presentation.gui.strings import tr


def number(value, percent=False):
    if value is None:
        return tr('pr2.unknown')
    if percent:
        return f'{value:.2%}'
    return f'{value:,.2f}'


def scroll_tree(parent, columns, title, headings, widths):
    frame = ttk.Frame(parent)
    frame.pack(fill='both', expand=True)
    tree = ttk.Treeview(frame, columns=columns, selectmode='browse', height=9)
    tree.heading('#0', text=title)
    tree.column('#0', width=370, minwidth=180)
    for col, label, width in zip(columns, headings, widths):
        tree.heading(col, text=label)
        tree.column(col, width=width, minwidth=65, stretch=False,
                    anchor='e' if col in ('percent','money') else 'w')
    vertical = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
    horizontal = ttk.Scrollbar(frame, orient='horizontal', command=tree.xview)
    tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
    tree.grid(row=0,column=0,sticky='nsew')
    vertical.grid(row=0,column=1,sticky='ns')
    horizontal.grid(row=1,column=0,sticky='ew')
    frame.rowconfigure(0,weight=1)
    frame.columnconfigure(0,weight=1)
    return tree


class PaymentPlanDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title(tr('pr2.title'))
        self.transient(parent.winfo_toplevel())
        scale = self.winfo_fpixels('1i') / 96
        width = min(round(980 * scale), self.winfo_screenwidth() - 36)
        height = min(round(650 * scale), self.winfo_screenheight() - 90)
        self.geometry(f'{width}x{height}')
        self.minsize(min(width, 640), min(height, 460))
        self.model = None
        self.reviewed = None
        self.plan_label = tk.StringVar(self,tr("pr4.original"))
        self.workbook_path = None
        self.export_cancel = None
        self.events = queue.Queue()
        self.busy = False
        self.poll_id = self.search_id = None
        self.query = tk.StringVar(self)
        self.mode = tk.StringVar(self, 'period')
        self.source_label = tk.StringVar(self, tr('pr2.intro'))
        self.status = tk.StringVar(self)
        self.scope = tk.StringVar(self)
        self.basis = tk.StringVar(self)
        self.date = tk.StringVar(self)
        self.total = tk.StringVar(self)
        self.retained = tk.StringVar(self)
        self.expanded = set()
        self.filtering = False
        self.controls = []
        self.protocol('WM_DELETE_WINDOW', self.destroy)
        body = ttk.Frame(self, padding=10, style='Surface.TFrame')
        body.pack(fill='both', expand=True)
        top = ttk.Frame(body)
        top.pack(fill='x')
        ttk.Label(top,text=tr('pr2.title'),style='WorkspaceTitle.TLabel').pack(side='left')
        self.open_button = ttk.Button(top,text=tr('pr2.browse'),command=self.browse,style='Accent.TButton')
        self.open_button.pack(side='right')
        self.action(top,'pr2.problems',self.details,'right')
        self.line(body,self.source_label)
        self.line(body,self.status)
        footer = ttk.Frame(body)
        footer.pack(side='bottom', fill='x', pady=(4,0))
        ttk.Button(footer,text=tr('xw.close'),command=self.destroy).pack(side='left')
        self.next_button = ttk.Button(footer,command=self.next_stage,style='Accent.TButton')
        self.export_button = ttk.Button(footer,text=tr('pr3.export'),command=self.export,style='Accent.TButton')
        self.back_button = ttk.Button(footer,text=tr('pu.back'),command=self.previous_stage)
        self.tabs = ttk.Notebook(body)
        self.tabs.pack(fill='both', expand=True, pady=6)
        select = ttk.Frame(self.tabs,padding=8)
        adjust = ttk.Frame(self.tabs,padding=10)
        preview = ttk.Frame(self.tabs,padding=8)
        self.tabs.add(select,text=tr('pu.select'))
        self.tabs.add(adjust,text=tr('pu.adjust'))
        self.tabs.add(preview,text=tr('pu.review'))
        self.tabs.enable_traversal()
        self.tabs.bind('<<NotebookTabChanged>>',self.stage_controls)
        tools = ttk.Frame(select)
        tools.pack(fill='x',pady=(0,6))
        ttk.Label(tools,text=tr('pr2.search')).pack(side='left')
        entry = ttk.Entry(tools,textvariable=self.query,width=8)
        entry.pack(side='left',padx=6,fill='x',expand=True)
        self.action(tools,'pr2.whole',lambda:self.apply({'PROJECT'}))
        self.action(tools,'pr2.clear',lambda:self.apply(set()))
        ttk.Label(select,text=tr('pr2.pick'),style='Muted.TLabel').pack(anchor='w',pady=2)
        selection_footer = ttk.Frame(select)
        selection_footer.pack(side='bottom',fill='x')
        self.line(selection_footer,self.scope)
        self.line(selection_footer,self.basis)
        row = ttk.Frame(selection_footer)
        row.pack(fill='x')
        self.action(row,'pr2.review',self.review_rows)
        self.action(row,'pr2.save',lambda:self.selection_file(True))
        self.action(row,'pr2.restore',lambda:self.selection_file(False))
        self.tree = scroll_tree(select,('role',),tr('pr2.name'),[tr('pr2.state')],[235])
        self.tree.bind('<Button-1>',self.tree_click)
        self.tree.bind('<space>',self.toggle_focused)
        self.tree.bind('<Return>',self.toggle_focused)
        bar = ttk.Frame(preview)
        bar.pack(fill='x',pady=3)
        self.period = ttk.Combobox(bar,state='readonly',width=7)
        self.period.pack(side='left',padx=(0,10))
        self.period.bind('<<ComboboxSelected>>',lambda _:self.draw_preview())
        for m in ('period','cumulative'):
            ttk.Radiobutton(bar,text=tr('pr2.'+m),value=m,variable=self.mode,
                            command=self.draw_preview).pack(side='left')
        ttk.Label(bar,textvariable=self.date).pack(side='right')
        summary = ttk.Frame(preview)
        summary.pack(side='bottom',fill='x')
        for variable in (self.total,self.retained,self.scope,self.basis):
            self.line(summary,variable)
        self.line(preview,self.plan_label)
        self.preview = scroll_tree(preview,('percent','money','state'),tr('pr2.name'),
                                   [tr('pr2.percent'),tr('pr2.money'),tr('pr2.status')],[105,140,190])

        ttk.Label(adjust,text=tr('pu.excel'),style='Section.TLabel').pack(anchor='w')
        ttk.Label(adjust,text=tr('pu.instructions'),wraplength=540,justify='left').pack(anchor='w',pady=(2,8))
        row = ttk.Frame(adjust)
        row.pack(fill='x')
        self.action(row,'pr4.create',self.create_plan)
        self.action(row,'pr4.open',self.open_plan)
        ttk.Separator(adjust).pack(fill='x',pady=12)
        ttk.Label(adjust,text=tr('pu.current'),style='Section.TLabel').pack(anchor='w')
        self.line(adjust,self.plan_label)
        row = ttk.Frame(adjust)
        row.pack(fill='x',pady=4)
        self.action(row,'pu.use_original',self.review_original)
        self.query.trace_add('write',self.schedule_search)
        self.enable()

    def review_original(self):
        self.original_plan()
        self.tabs.select(2)

    def next_stage(self):
        self.tabs.select(min(2,self.tabs.index('current')+1))

    def previous_stage(self):
        self.tabs.select(max(0,self.tabs.index('current')-1))

    def stage_controls(self, _event=None):
        stage = self.tabs.index('current')
        for button in (self.next_button,self.export_button,self.back_button):
            button.pack_forget()
        if stage == 2 or self.export_cancel is not None:
            self.export_button.pack(side='right')
        else:
            self.next_button.configure(text=tr('pu.to_adjust') if stage==0 else tr('pu.to_review'))
            self.next_button.pack(side='right')
        if stage:
            self.back_button.pack(side='right',padx=4)

    @staticmethod
    def line(parent,var):
        label = ttk.Label(parent,textvariable=var,justify='left')
        label.pack(fill='x',pady=1)
        label.bind('<Configure>',lambda e: label.configure(wraplength=max(120,e.width-4)))

    def action(self,parent,key,command,side='left'):
        btn=ttk.Button(parent,text=tr(key),command=command)
        btn.pack(side=side,padx=3)
        self.controls.append(btn)

    def enable(self):
        self.stage_controls()
        if self.export_cancel is not None:
            self.export_button.configure(text=tr('pr3.cancel'))
            self.export_button.state(['!disabled'])
        else:
            self.export_button.configure(text=tr('pr3.export'))
            ready = self.model and self.model.keys and self.workbook_path and not self.busy
            ready = ready and not (self.reviewed and self.reviewed.incomplete)
            self.export_button.state(['!disabled' if ready else 'disabled'])
        self.open_button.state(['disabled' if self.busy else '!disabled'])
        for btn in self.controls:
            btn.state(['!disabled' if self.model and not self.busy else 'disabled'])

    def browse(self):
        name=filedialog.askopenfilename(parent=self,title=tr('pr2.browse'),
                                        filetypes=[(tr('pr2.book'),'*.xlsx *.xlsm')])
        if name:
            self.load(Path(name))

    def load(self,path):
        if self.busy:
            return
        self.busy=True
        self.status.set(tr('pr2.loading',file=path.name))
        self.enable()
        def worker():
            try:
                self.events.put((PlanSelection(read_payment_plan_source(path)),None))
            except Exception as exc:
                self.events.put((None,str(exc)))
        threading.Thread(target=worker,daemon=True).start()
        def poll():
            try:
                candidate,error=self.events.get_nowait()
            except queue.Empty:
                self.poll_id=self.after(50,poll)
                return
            self.poll_id=None
            self.busy=False
            self.enable()
            if error:
                self.problem(tr('pr2.failed',file=path.name,error=error))
                return
            same=self.model and self.model.source.revision==candidate.source.revision
            if same:
                candidate.select(self.model.keys)
            elif self.model and self.model.keys and not messagebox.askyesno(
                    tr('pr2.title'),tr('pr2.replace'),parent=self):
                self.source_status()
                return
            self.workbook_path = path
            self.clear_adjustment()
            self.model=candidate
            source=candidate.source
            self.source_label.set(tr('pr2.source',project=source.project,file=source.filename,
                                     count=len(source.periods),origin=source.monetary_origin))
            self.period.configure(values=[p.id for p in source.periods])
            self.period.current(0)
            self.query.set('')
            self.expanded={'PROJECT'}
            self.filtering=False
            self.tree.delete(*self.tree.get_children())
            self.draw_tree()
            self.refresh()
            self.source_status()
            self.enable()
        self.poll_id=self.after(25,poll)

    def problem(self,text):
        self.status.set(text.splitlines()[0][:145])
        messagebox.showerror(tr('pr2.title'),text,parent=self)

    def issues(self):
        source=self.model.source
        issues=list(source.diagnostics)
        for activity,leaf in source.leaves.items():
            if leaf.schedule and leaf.schedule[0].origin == 'Unknown':
                issues.append(tr('pr2.missing_schedule',activity=activity))
        issues.extend(p.id+'  '+issue for p in source.periods for issue in p.issues)
        if self.reviewed:
            issues.extend(self.reviewed.issues)
        return issues

    def source_status(self):
        issues=self.issues()
        self.status.set(tr('pr2.warning',n=len(issues)) if issues else tr('pr2.ready'))

    def details(self):
        source=self.model.source
        win=tk.Toplevel(self)
        win.title(tr('pr2.problems'))
        win.transient(self)
        text=ScrolledText(win,width=78,height=16,wrap='word')
        text.pack(fill='both',expand=True,padx=10,pady=10)
        text.insert('end','\n'.join([self.source_label.get(),tr('pr2.revision_id')+'  '+source.revision,
                                    *self.issues()]))
        text.configure(state='disabled')

    def review_rows(self):
        win=tk.Toplevel(self)
        win.title(tr('pr2.review'))
        win.transient(self)
        win.geometry('880x390')
        frame=ttk.Frame(win,padding=12)
        frame.pack(fill='both',expand=True)
        ttk.Label(frame,text=self.scope.get()).pack(anchor='w',pady=3)
        result=self.result
        ttk.Label(frame,text=tr('pr2.reconcile',selected=number(result.selected.basis),
                                remaining=number(result.unselected.basis),
                                project=number(result.project.basis))).pack(anchor='w',pady=3)
        tree=scroll_tree(frame,('basis','role'),tr('pr2.name'),
                         [tr('pr2.row_basis'),tr('pr2.state')],[155,235])
        for row in result.rows:
            ancestor=result.included_in.get(row.key)
            role=tr('pr2.selected') if ancestor is None else tr(
                'pr2.included',parent=self.label(self.model.by_key[ancestor]))
            tree.insert('','end',iid=row.key,text=self.label(row),
                        values=(number(result.row_values[row.key].basis),role))
        if not result.rows:
            tree.insert('','end',text=tr('pr2.empty'))
        win.grab_set()

    def schedule_search(self,*_):
        if self.search_id is not None:
            self.after_cancel(self.search_id)
        self.search_id=self.after(100,self.draw_tree)

    def draw_tree(self):
        if self.search_id is not None:
            self.after_cancel(self.search_id)
            self.search_id=None
        if not self.model:
            return
        tree=self.tree
        if not self.filtering and tree.get_children():
            self.expanded={r.key for r in self.model.source.rows if tree.exists(r.key)
                           and tree.item(r.key,'open')}
        self.filtering=bool(self.query.get().strip())
        tree.delete(*tree.get_children())
        for r in self.model.matching(self.query.get()):
            tree.insert(r.parent or '', 'end',iid=r.key,open=self.filtering or r.key in self.expanded)
        self.paint_checks()

    def label(self,row):
        identifier=row.activities[0] if row.activity else row.code
        return (identifier+'   '+row.name).strip()

    def paint_checks(self):
        model=self.model
        for r in model.source.rows:
            if not self.tree.exists(r.key):
                continue
            checked=r.key in model.keys
            ancestor=r.parent
            while ancestor and ancestor not in model.keys:
                ancestor=model.by_key[ancestor].parent
            role=tr('pr2.selected') if checked else ''
            if ancestor:
                role=tr('pr2.included' if checked else 'pr2.represented',
                        parent=self.label(model.by_key[ancestor]))
            self.tree.item(r.key,text=('☑  ' if checked else '☐  ')+self.label(r),values=(role,))

    def tree_click(self,event):
        tree=self.tree
        key=tree.identify_row(event.y)
        if key and tree.identify_column(event.x)=='#0' and 'indicator' not in tree.identify_element(event.x,event.y):
            tree.focus(key)
            tree.selection_set(key)
            tree.focus_set()
            self.toggle(key)
            return 'break'

    def toggle_focused(self,_):
        key=self.tree.focus()
        if key:
            self.toggle(key)
        return 'break'

    def toggle(self,key):
        if self.model and not self.busy:
            self.apply(self.model.keys ^ {key})

    def apply(self,keys):
        self.clear_adjustment()
        self.model.select(keys)
        self.paint_checks()
        self.refresh()

    def refresh(self):
        self.enable()
        r=self.result
        state='empty' if not r.rows else 'full' if r.full_project else 'partial'
        self.scope.set(tr('pr2.scope',scope=tr('pr2.'+state),rows=len(r.rows),count=len(r.selected.activities),
                          all=len(r.project.activities),overlap=len(r.included_in)))
        coverage=None
        if r.selected.basis is not None and r.project.basis:
            coverage=r.selected.basis/r.project.basis
        self.basis.set(tr('pr2.basis',selected=number(r.selected.basis),project=number(r.project.basis),
                          coverage=number(coverage,True)))
        self.draw_preview()

    def draw_preview(self):
        if not self.model or self.period.current()<0:
            return
        result=self.result
        j=self.period.current()
        p=self.model.source.periods[j]
        self.date.set(tr('pr2.eligible')+'   '+(p.eligible.isoformat() if p.eligible else tr('pr2.unknown')))
        cumulative=self.mode.get()=='cumulative'
        tree=self.preview
        tree.delete(*tree.get_children())
        for row in result.rows:
            v=result.row_values[row.key].values[j]
            if not (v.cumulative_visible if cumulative else v.period_visible):
                continue
            notes=[]
            if v.transition!='open':
                notes.append(tr('pr2.'+v.transition))
            if v.correction:
                notes.append(tr('pr2.decrease'))
            if row.key in result.included_in:
                notes.append(tr('pr2.included',parent=self.label(self.model.by_key[result.included_in[row.key]])))
            tree.insert('','end',iid=row.key,text=self.label(row),values=(
                number(v.display_fraction if cumulative else v.delta,True),
                number(v.money if cumulative else v.movement),' · '.join(notes)))
        if not tree.get_children():
            tree.insert('','end',text=tr('pr2.no_detail' if result.rows else 'pr2.empty'))
        v=result.selected.values[j]
        self.total.set(tr('pr2.total' if cumulative else 'pr2.known_period',
                          total=number(v.money if cumulative else v.movement),
                          percent=number(v.display_fraction if cumulative else v.delta,True),known=number(v.known_money),
                          basis=number(v.known_basis),n=v.known_leaves,all=len(result.selected.activities)))
        self.retained.set(tr('pr2.retained',visible=number(result.visible_known_money[j]),
                             kept=number(result.retained_money[j])) if cumulative else '')

    @property
    def result(self):
        return self.reviewed.selection if self.reviewed else self.model.result

    def clear_adjustment(self):
        self.reviewed = None
        self.plan_label.set(tr('pr4.original'))

    def original_plan(self):
        self.clear_adjustment()
        self.refresh()
        self.source_status()

    def create_plan(self):
        from progress_studio.infrastructure.excel.payment_plan_adjustment import create_adjustment
        if not self.model.keys:
            self.problem(tr('pr2.empty'))
            return
        path = filedialog.asksaveasfilename(parent=self,defaultextension='.xlsx',filetypes=[('Excel','*.xlsx')])
        if path:
            try:
                create_adjustment(path,self.model.source,self.model.result)
                self.status.set(tr('pr4.created',path=path))
            except Exception as exc:
                self.problem(str(exc))

    def open_plan(self):
        from hashlib import sha256
        from progress_studio.infrastructure.excel.payment_plan_adjustment import read_adjustment
        path = filedialog.askopenfilename(parent=self,filetypes=[('Excel','*.xlsx')])
        if not path:
            return
        try:
            if sha256(Path(self.workbook_path).read_bytes()).hexdigest()!=self.model.source.revision:
                raise ValueError(tr('pr4.stale'))
            plan = read_adjustment(path,self.model.source)
        except Exception as exc:
            self.problem(str(exc))
            return
        self.model.select([row.key for row in plan.selection.rows])
        self.reviewed = plan
        self.plan_label.set(tr('pr4.adjusted',name=plan.path.name))
        self.paint_checks()
        self.refresh()
        self.source_status()

    def selection_file(self,save):
        kwargs=dict(parent=self,filetypes=[(tr('pr2.selection_file'),'*'+SUFFIX)])
        if save:
            path=filedialog.asksaveasfilename(**kwargs,defaultextension=SUFFIX,
                                              initialfile=Path(self.model.source.filename).stem+SUFFIX)
        else:
            path=filedialog.askopenfilename(**kwargs)
        if not path:
            return
        try:
            if save:
                self.model.save(path)
            else:
                self.model.restore(path)
                self.clear_adjustment()
                self.paint_checks()
                self.refresh()
            self.source_status()
            self.status.set(tr('pr2.saved' if save else 'pr2.restored')+'  •  '+self.status.get())
        except (ValueError,OSError) as exc:
            self.problem(str(exc))

    def export(self):
        if self.export_cancel is not None:
            self.export_cancel.set()
            self.status.set(tr('pr3.cancelling'))
            return
        if self.busy or not self.model or not self.model.keys or not self.workbook_path:
            return
        parent = filedialog.askdirectory(parent=self, title=tr('pr3.destination'), mustexist=True)
        if not parent:
            return
        name = simpledialog.askstring(tr('pr3.export'), tr('pr3.folder'),
                                      initialvalue=self.model.source.project, parent=self)
        if name is None:
            return
        source, selection, workbook = self.model.source, self.result, self.workbook_path
        reviewed = self.reviewed
        cancel = threading.Event()
        self.export_cancel = cancel
        self.busy = True
        self.enable()
        self.status.set(tr('pr3.start'))
        events = queue.Queue()

        def worker():
            try:
                from progress_studio.services.payment_plan_export import export_payment_plans, ExportCancelled
                try:
                    folder = export_payment_plans(source, selection, workbook, parent, name,
                                                  reviewed=reviewed, cancelled=cancel.is_set,
                                                  progress=lambda stage: events.put(('stage', stage)))
                except ExportCancelled:
                    events.put(('cancelled', None))
                    return
                events.put(('done', folder))
            except Exception as exc:
                events.put(('error', str(exc)))

        def poll():
            while True:
                try:
                    event, value = events.get_nowait()
                except queue.Empty:
                    self.poll_id = self.after(60, poll)
                    return
                if event == 'stage':
                    if not cancel.is_set():
                        self.status.set(tr('pr3.period' if value == 'PERIOD' else 'pr3.cumulative'))
                    continue
                self.poll_id = None
                self.export_cancel = None
                self.busy = False
                self.enable()
                if event == 'error':
                    self.problem(tr('pr3.failed', error=value))
                elif event == 'cancelled':
                    self.status.set(tr('pr3.cancelled'))
                else:
                    self.status.set(tr('pr3.done', folder=value))
                    messagebox.showinfo(tr('pr3.export'), self.status.get(), parent=self)
                return

        threading.Thread(target=worker, daemon=False).start()
        self.poll_id = self.after(40, poll)

    def destroy(self):
        if self.export_cancel is not None:
            self.export_cancel.set()
            self.status.set(tr('pr3.cancelling'))
            return
        for job in (self.search_id,self.poll_id):
            if job is not None:
                self.after_cancel(job)
        super().destroy()
