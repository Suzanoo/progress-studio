"""Explicit custom XML field selection and pre-Create preview."""
from pathlib import Path
import hashlib
import tkinter as tk
from tkinter import ttk

from progress_studio.infrastructure.schedule_xml import NormalizedScheduleXmlReader
from progress_studio.services.xml_amount_service import preview_amounts


class AmountFieldDialog(tk.Toplevel):
    def __init__(self, parent, source_xml: Path):
        # Read before opening a modal window, so parse failures leave no orphan dialog.
        self.source_digest = hashlib.sha256(source_xml.read_bytes()).hexdigest()
        _, self.rows, self.fields = NormalizedScheduleXmlReader().read_with_amount_fields(source_xml)
        super().__init__(parent)
        self.title('Select XML Amount field')
        self.geometry('1060x600')
        self.transient(parent)
        self.result = None
        self.preview = None
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)
        ttk.Label(self, text='Select a numeric custom field. Values become monetary weights; milestones remain zero.').grid(
            row=0, column=0, sticky='w', padx=12, pady=10)
        self.choice = ttk.Combobox(self, state='readonly', values=[f.label for f in self.fields])
        self.choice.grid(row=1, column=0, sticky='ew', padx=12)
        self.choice.bind('<<ComboboxSelected>>', self._preview)
        self.summary = tk.StringVar(value='Choose a field to preview.' if self.fields else
                                    'No declared numeric activity/task custom fields were found in this XML.')
        ttk.Label(self, textvariable=self.summary, wraplength=1000).grid(row=2, column=0, sticky='w', padx=12, pady=10)
        table = ttk.Frame(self)
        table.grid(row=3, column=0, sticky='nsew', padx=12)
        table.columnconfigure(0, weight=1); table.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(table, columns=('id', 'name', 'raw', 'weight', 'status'), show='headings')
        for key, title, width in [('id', 'Activity ID', 115), ('name', 'Activity', 270),
                                  ('raw', 'XML value (unrounded)', 210), ('weight', 'Weight (display)', 140),
                                  ('status', 'Validation', 255)]:
            self.tree.heading(key, text=title); self.tree.column(key, width=width, minwidth=80)
        self.tree.grid(row=0, column=0, sticky='nsew')
        scrollbar = ttk.Scrollbar(table, orient='vertical', command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky='ns'); self.tree.configure(yscrollcommand=scrollbar.set)
        horizontal = ttk.Scrollbar(table, orient='horizontal', command=self.tree.xview)
        horizontal.grid(row=1, column=0, sticky='ew'); self.tree.configure(xscrollcommand=horizontal.set)
        actions = ttk.Frame(self); actions.grid(row=4, column=0, sticky='e', padx=12, pady=12)
        ttk.Button(actions, text='Cancel', command=self.destroy).pack(side='left', padx=6)
        self.accept_button = ttk.Button(actions, text='Use selected field', command=self._accept, state='disabled')
        self.accept_button.pack(side='left')
        self.bind('<Escape>', lambda _: self.destroy())
        self.grab_set()

    def _preview(self, _event=None):
        index = self.choice.current()
        if index < 0:
            return
        self.accept_button.configure(state='disabled')
        self.tree.delete(*self.tree.get_children())
        try:
            p = preview_amounts(self.rows, self.fields, self.fields[index].identity)
        except ValueError as exc:
            self.preview = None; self.summary.set(str(exc)); return
        self.preview = p
        for row in p.rows:
            self.tree.insert('', 'end', values=(row.activity_id, row.name,
                             row.raw_value if row.raw_value is not None else '(missing)',
                             f'{row.value:,.2f}' if row.value is not None else '', row.status))
        message = (f'Ordinary: {p.ordinary_count} | Positive: {p.positive_count} | Zero: {p.zero_count} | '
                   f'Milestones: {p.milestone_count} (zero weight) | '
                   f'{"Valid-value subtotal" if p.errors else "Total"}: {p.total:,.2f}')
        if p.errors:
            message += f'\nCreate blocked: {len(p.errors)} issue(s). {p.errors[0]}'
        else:
            message += '\nValues are retained at workbook numeric precision; two decimals are display only.'
        self.summary.set(message)
        self.accept_button.configure(state='normal' if p.valid else 'disabled')

    def _accept(self):
        if self.preview is not None and self.preview.valid:
            self.result = self.preview.field
            self.destroy()
