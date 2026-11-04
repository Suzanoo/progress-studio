"""Review only: closing this dialog never mutates a project."""
import tkinter as tk
from tkinter import ttk

from progress_studio.presentation.gui.strings import tr


def review_boq_update(parent, preview) -> bool:
    dialog = tk.Toplevel(parent)
    dialog.title(tr('boq_update.title'))
    dialog.transient(parent.winfo_toplevel())
    dialog.geometry('900x620')
    dialog.minsize(650, 420)
    accepted = False
    ttk.Label(dialog, text=tr('boq_update.source', file=preview.candidate_path.name,
                             sheet=preview.original.boq_sheet), wraplength=840).pack(anchor='w', padx=12, pady=(12, 4))
    ttk.Label(dialog, text=tr('boq_update.summary', changed=len(preview.item_changes),
                             unchanged=len(preview.new_items)-len(preview.item_changes),
                             mapped=preview.mapped_count,
                             unmapped=len(preview.new_items)-preview.mapped_count,
                             old=preview.before_total, new=preview.after_total,
                             old_mapped=preview.before_mapped, new_mapped=preview.after_mapped),
              wraplength=840, justify='left').pack(anchor='w', padx=12, pady=4)
    ttk.Label(dialog, text=tr('boq_update.formulas', old=preview.old_formula_count,
                             new=preview.new_formula_count), wraplength=840,
              justify='left').pack(anchor='w', padx=12, pady=4)
    tabs = ttk.Notebook(dialog)
    tabs.pack(fill='both', expand=True, padx=12, pady=8)
    for title, changes in ((tr('boq_update.items'), preview.item_changes),
                           (tr('boq_update.activities'), preview.activity_changes)):
        pane = ttk.Frame(tabs)
        tabs.add(pane, text=title)
        pane.rowconfigure(0, weight=1)
        pane.columnconfigure(0, weight=1)
        columns = ('id', 'description', 'old', 'new', 'delta')
        tree = ttk.Treeview(pane, columns=columns, show='headings')
        labels = (tr('boq_update.identity'), tr('Description'), tr('boq_update.old'),
                  tr('boq_update.new'), tr('boq_update.delta'))
        for col, label, width in zip(columns, labels, (150, 310, 130, 130, 130)):
            tree.heading(col, text=label)
            tree.column(col, width=width, minwidth=90, anchor='w' if col in ('id', 'description') else 'e')
        tree.grid(row=0, column=0, sticky='nsew')
        vertical = ttk.Scrollbar(pane, orient='vertical', command=tree.yview)
        vertical.grid(row=0, column=1, sticky='ns')
        horizontal = ttk.Scrollbar(pane, orient='horizontal', command=tree.xview)
        horizontal.grid(row=1, column=0, sticky='ew')
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        for change in changes:
            tree.insert('', 'end', values=(change.identity, change.description,
                        f'{change.before:,.2f}', f'{change.after:,.2f}', f'{change.after-change.before:+,.2f}'))
    ttk.Label(dialog, text=tr('boq_update.save_notice'), wraplength=840).pack(anchor='w', padx=12)
    buttons = ttk.Frame(dialog)
    buttons.pack(fill='x', padx=12, pady=12)

    def confirm():
        nonlocal accepted
        accepted = True
        dialog.destroy()

    ttk.Button(buttons, text=tr('boq_update.apply'), command=confirm,
               state='normal' if preview.item_changes else 'disabled').pack(side='right')
    ttk.Button(buttons, text=tr('Cancel'), command=dialog.destroy).pack(side='right', padx=8)
    dialog.bind('<Escape>', lambda _event: dialog.destroy())
    dialog.grab_set()
    dialog.wait_window()
    return accepted
