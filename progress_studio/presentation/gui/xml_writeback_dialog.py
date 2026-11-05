"""Desktop orchestration only; monetary/XML decisions belong to XW-1.

Workers receive immutable requests/assessments and never call Tk. Only the UI
thread consumes their queue. Closing during work requests cancellation and keeps
the dialog alive until the actual result is displayed (publication may have won).
"""
from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import filedialog, ttk
from tkinter.scrolledtext import ScrolledText

from progress_studio.services.xml_writeback_service import (
    XmlWritebackService, WritebackRequest, WritebackCancelled,
)
from progress_studio.presentation.gui.strings import tr


def review_rows(assessment):
    """Present engine evidence without interpreting monetary or milestone rules."""
    for row in assessment.rows:
        yield (row.activity_id, row.name,
               tr('xw.missing') if row.before_raw is None else row.before_raw,
               tr('xw.missing') if row.after_raw is None else row.after_raw,
               tr('xw.' + row.status))
    for identity in assessment.target_only_ids:
        yield (identity, tr('xw.not_assessed'), tr('xw.not_assessed'),
               tr('xw.not_assessed'), tr('xw.target_only'))


class XmlWritebackDialog(tk.Toplevel):
    def __init__(self, parent, service=None):
        super().__init__(parent)
        self.service = service or XmlWritebackService()
        self.assessment = None
        self.fields = ()
        self.busy = False
        self.events = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker = None
        self.workbook = tk.StringVar(self)
        self.target = tk.StringVar(self)
        self.field_choice = tk.StringVar(self)
        self.monetary = tk.BooleanVar(self, False)
        self.recalculated = tk.BooleanVar(self, False)
        self.status = tk.StringVar(self, tr('xw.start'))
        self.controls = []
        self.title(tr('xw.title'))
        self.transient(parent.winfo_toplevel())
        self.geometry('1080x780')
        self.minsize(760, 620)
        self._build_ui()
        self.protocol('WM_DELETE_WINDOW', self.request_close)
        self.bind('<Escape>', lambda _e: self.request_close())
        self.grab_set()
        self._sync_controls()

    def _build_ui(self):
        body = ttk.Frame(self, padding=12)
        body.pack(fill='both', expand=True)
        ttk.Label(body, text=tr('xw.help'), wraplength=940, justify='left').pack(fill='x')
        inputs = ttk.Frame(body)
        inputs.pack(fill='x', pady=6)
        inputs.columnconfigure(1, weight=1)
        for index, (key, variable, kind) in enumerate((
                ('xw.workbook', self.workbook, 'workbook'), ('xw.target', self.target, 'target'))):
            ttk.Label(inputs, text=tr(key)).grid(row=index, column=0, sticky='w', padx=(0, 8))
            ttk.Entry(inputs, textvariable=variable, state='readonly').grid(row=index, column=1, sticky='ew')
            button = ttk.Button(inputs, text=tr('Browse...'), command=lambda k=kind: self._pick(k))
            button.grid(row=index, column=2, padx=6, pady=3)
            self.controls.append((button, 'normal'))
        ttk.Label(inputs, text=tr('xw.field')).grid(row=2, column=0, sticky='w')
        self.field_box = ttk.Combobox(inputs, textvariable=self.field_choice, state='readonly')
        self.field_box.grid(row=2, column=1, sticky='ew')
        self.field_box.bind('<<ComboboxSelected>>', lambda _e: self._invalidate())
        self.controls.append((self.field_box, 'readonly'))
        details = ttk.Button(inputs, text=tr('xw.field_details'), command=self._field_details)
        details.grid(row=2, column=2, padx=6)
        self.controls.append((details, 'normal'))
        monetary = ttk.Checkbutton(body, text=tr('xw.monetary'), variable=self.monetary, command=self._invalidate)
        monetary.pack(anchor='w')
        self.controls.append((monetary, 'normal'))
        self.formula_check = ttk.Checkbutton(body, text=tr('xw.formula_confirm'),
                                           variable=self.recalculated, command=self._invalidate)
        self.controls.append((self.formula_check, 'normal'))
        self.formula_host = body
        self.summary = ttk.Label(body, text='', wraplength=960, justify='left')
        self.summary.pack(fill='x', pady=6)
        pane = ttk.Frame(body)
        pane.pack(fill='both', expand=True)
        pane.rowconfigure(0, weight=1)
        pane.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(pane, columns=('id', 'name', 'before', 'after', 'status'), show='headings')
        for col, key, width in zip(self.tree['columns'],
                ('xw.id', 'Description', 'xw.before', 'xw.after', 'xw.state'), (150, 260, 180, 180, 200)):
            self.tree.heading(col, text=tr(key))
            self.tree.column(col, width=width, minwidth=100, stretch=True)
        self.tree.grid(row=0, column=0, sticky='nsew')
        vs = ttk.Scrollbar(pane, orient='vertical', command=self.tree.yview)
        hs = ttk.Scrollbar(pane, orient='horizontal', command=self.tree.xview)
        vs.grid(row=0, column=1, sticky='ns'); hs.grid(row=1, column=0, sticky='ew')
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.bind('<<TreeviewSelect>>', self._row_details)
        self.diagnostics = ScrolledText(body, height=7, wrap='word', state='disabled')
        self.diagnostics.pack(fill='x', pady=6)
        ttk.Label(body, textvariable=self.status, wraplength=950, justify='left').pack(fill='x')
        self.progress = ttk.Progressbar(body, mode='indeterminate')
        self.progress.pack(fill='x', pady=4)
        actions = ttk.Frame(body)
        actions.pack(fill='x')
        self.review_button = ttk.Button(actions, text=tr('xw.review'), command=self._review)
        self.review_button.pack(side='left')
        self.export_button = ttk.Button(actions, text=tr('xw.export'), command=self._export)
        self.export_button.pack(side='right')
        self.close_button = ttk.Button(actions, text=tr('xw.close'), command=self.request_close)
        self.close_button.pack(side='right', padx=6)

    def _text(self, text):
        self.diagnostics.configure(state='normal')
        self.diagnostics.delete('1.0', 'end')
        self.diagnostics.insert('1.0', text)
        self.diagnostics.configure(state='disabled')

    def _invalidate(self):
        self.assessment = None
        self.tree.delete(*self.tree.get_children())
        self.summary.configure(text='')
        self._review_diagnostics = ''
        self._text('')
        self.status.set(tr('xw.changed'))
        self._sync_controls()

    def _sync_controls(self):
        for widget, state in self.controls:
            widget.configure(state='disabled' if self.busy else state)
        ready = bool(self.workbook.get() and self.target.get())
        self.review_button.configure(state='normal' if ready and not self.busy else 'disabled')
        valid = self.assessment is not None and self.assessment.valid
        self.export_button.configure(state='normal' if valid and not self.busy else 'disabled')
        self.close_button.configure(text=tr('xw.cancel_work') if self.busy else tr('xw.close'))

    def _pick(self, kind):
        if self.busy:
            return
        types = [(tr('xw.workbook'), '*.xlsx *.xlsm')] if kind == 'workbook' else [('P6 XML', '*.xml')]
        path = filedialog.askopenfilename(parent=self, title=tr('xw.' + kind), filetypes=types)
        if not path:
            return
        getattr(self, kind).set(path)
        self.monetary.set(False)
        self.recalculated.set(False)
        self.field_choice.set('')
        self.fields = ()
        self.field_box.configure(values=())
        self.formula_check.pack_forget()
        self._invalidate()
        if self.workbook.get() and self.target.get():
            self._start('discover', self._request())

    def _request(self):
        index = self.field_box.current()
        field = self.fields[index] if 0 <= index < len(self.fields) else None
        return WritebackRequest(Path(self.workbook.get()), Path(self.target.get()),
            field_identity=field.identity if field else None,
            create_progress_studio_amount=index == len(self.fields) and index >= 0,
            monetary_confirmed=self.monetary.get(), recalculated_and_saved=self.recalculated.get())

    def _review(self):
        if self.busy or not (self.workbook.get() and self.target.get()):
            return
        request = self._request()
        self._invalidate()
        self._start('review', request)

    def _export(self):
        if self.busy or self.assessment is None or not self.assessment.valid:
            return
        assessment = self.assessment
        path = filedialog.asksaveasfilename(parent=self, title=tr('xw.save'),
            defaultextension='.xml', filetypes=[('P6 XML', '*.xml')],
            initialfile=assessment.request.target_xml.stem + '_amounts.xml')
        if not path:
            return
        # Native dialogs run a nested event loop. Never use a replaced review.
        if self.busy or self.assessment is not assessment:
            return
        self._start('export', assessment, Path(path))

    def _start(self, action, payload, output=None):
        if self.busy:
            return
        self.busy = True
        self.cancel_event.clear()
        self.status.set(tr('xw.exporting') if action == 'export' else tr('xw.assessing'))
        self._sync_controls()
        self.progress.start()

        # Do not capture the dialog/Tk variables in the worker closure.
        service, events, cancelled = self.service, self.events, self.cancel_event.is_set

        def run():
            try:
                value = (service.export(payload, output, reviewed=True, cancelled=cancelled)
                         if action == 'export' else service.assess(payload))
                events.put((action, value, None))
            except Exception as exc:
                events.put((action, None, exc))

        self.worker = threading.Thread(target=run, daemon=False)
        try:
            self.worker.start()
        except Exception as exc:
            self.events.put((action, None, exc))
        self.after(75, self._poll)

    def _poll(self):
        try:
            action, value, error = self.events.get_nowait()
        except queue.Empty:
            self.after(75, self._poll)
            return
        self.busy = False
        self.progress.stop()
        if error is not None:
            self.assessment = None
            self.status.set(tr('xw.cancelled') if isinstance(error, WritebackCancelled) else tr('xw.failed'))
            self._review_diagnostics = str(error) + '\n\n' + tr('xw.recovery')
            self._text(self._review_diagnostics)
        elif action == 'export':
            # Success wins over a late cancellation request; always show real output.
            self.assessment = None
            self.status.set(tr('xw.success', path=value.output))
            self._review_diagnostics = '\n'.join(value.warnings) + '\n' + tr('xw.p6_notice')
            self._text(self._review_diagnostics)
        elif self.cancel_event.is_set():
            self.assessment = None
            self.status.set(tr('xw.review_cancelled'))
        else:
            self._show_assessment(value, discovery=action == 'discover')
        self._sync_controls()

    def _show_assessment(self, value, *, discovery=False):
        if value.formula_count:
            self.formula_check.pack(before=self.summary, anchor='w')
        else:
            self.formula_check.pack_forget()
        # Display only Cost candidates. Engine still decides actual compatibility.
        fields = tuple(f for f in value.available_fields if f.data_type == 'Cost')
        if discovery or fields != self.fields:
            self.fields = fields
            labels = [f'{i + 1}. {f.name} — Activity / Cost' for i, f in enumerate(fields)]
            self.field_box.configure(values=labels + [tr('xw.create')])
            self.field_choice.set('')
            discovery = True
        self.assessment = None if discovery else value
        self.tree.delete(*self.tree.get_children())
        for row in review_rows(value):
            self.tree.insert('', 'end', values=row)
        field_name = value.field.name if value.field else tr('xw.choose')
        self.summary.configure(text=tr('xw.summary', project=value.project_id or '—', field=field_name,
            source=value.source_count, matched=len(value.rows), replaced=value.replacement_count,
            unchanged=value.unchanged_count, milestones=value.preserved_milestone_count,
            other=len(value.target_only_ids), zeros=value.zero_count, formulas=value.formula_count,
            before=value.before_known_total, missing=value.before_missing_count, after=value.write_total))
        notes = [tr('xw.discovery') if discovery else tr('xw.valid') if value.valid else tr('xw.blocked')]
        if value.creates_field:
            notes.append(tr('xw.new_field'))
        if value.valid and not value.replacement_count and not discovery:
            notes.append(tr('xw.no_changes'))
        notes.extend([tr('xw.blockers'), *value.errors, tr('xw.warnings'), *value.warnings,
                      tr('xw.totals_notice')])
        self._review_diagnostics = '\n'.join(notes)
        self._text(self._review_diagnostics)
        self.status.set(notes[0])

    def _field_details(self):
        if self.busy:
            return
        lines = [tr('xw.details_notice')]
        lines.extend(f'{i + 1}. {f.name} | Activity / {f.data_type} | {f.identity}'
                     for i, f in enumerate(self.fields))
        if self.assessment and self.assessment.field:
            lines.append(self.assessment.field.label)
        # Keep blockers/warnings visible, even when inspecting details.
        self._text(getattr(self, '_review_diagnostics', '') + '\n\n' + '\n'.join(lines))

    def _row_details(self, _event=None):
        selection = self.tree.selection()
        if selection and not self.busy:
            values = self.tree.item(selection[0], 'values')
            self._text(getattr(self, '_review_diagnostics', '') + '\n\n' + '\n'.join(map(str, values)))

    def request_close(self):
        if self.busy:
            self.cancel_event.set()
            self.status.set(tr('xw.cancelling'))
            return
        self.destroy()
