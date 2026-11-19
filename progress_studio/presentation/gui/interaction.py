"""Small Tk interaction helpers; business commands remain in their workspaces."""
from __future__ import annotations
import tkinter as tk
from .strings import tr

EDITORS = {"Entry", "TEntry", "Text", "TCombobox", "Spinbox", "TSpinbox"}


def is_editor(widget):
    return widget is not None and widget.winfo_class() in EDITORS


def active_operations(widget):
    """Inspect live widgets, including untracked report/prepare Toplevels."""
    result = []
    values = vars(widget)
    busy = any(values.get(key) is True for key in ("busy", "_busy", "operation_active"))
    future = values.get("future")
    workers = (values.get(key) for key in ("worker", "_worker", "_export_worker"))
    if busy or future is not None or any(w is not None and w.is_alive() for w in workers):
        result.append(widget)
    for child in widget.winfo_children():
        result.extend(active_operations(child))
    return result


def local_wheel(widget):
    """Keep wheel ownership local; never replace or delete another widget's bindings."""
    def scroll(event):
        delta = getattr(event, 'delta', 0)
        num = getattr(event, 'num', 0)
        if not delta and num not in (4, 5):
            return
        amount = -1 if delta > 0 or num == 4 else 1
        widget.yview_scroll(amount * 3, 'units')
        return 'break'
    for sequence in ('<MouseWheel>', '<Button-4>', '<Button-5>'):
        widget.bind(sequence, scroll)


def copy_tree(tree, rows=None):
    keys = list(rows if rows is not None else tree.selection())
    if not keys and tree.focus():
        keys = [tree.focus()]
    lines = []
    for key in keys:
        if tree.exists(key):
            item = tree.item(key)
            lines.append('\t'.join(str(v) for v in (item['text'], *item['values'])))
    if lines:
        tree.clipboard_clear()
        tree.clipboard_append('\n'.join(lines))
    return 'break'


def install_text_menu(widget):
    if not is_editor(widget):
        return
    menu = tk.Menu(widget, tearoff=False)
    for label, event in [('ux.cut', '<<Cut>>'), ('ux.copy', '<<Copy>>'), ('ux.paste', '<<Paste>>'), ('ux.select_all', '<<SelectAll>>')]:
        menu.add_command(label=tr(label), command=lambda ev=event: widget.event_generate(ev))
    def show(event):
        widget.focus_set()
        readonly = widget.cget('state') in ('readonly', 'disabled')
        for index in (0, 2):
            menu.entryconfigure(index, state='disabled' if readonly else 'normal')
        try:
            menu.tk_popup(event.x_root or widget.winfo_rootx(), event.y_root or widget.winfo_rooty())
        finally:
            menu.grab_release()
        return 'break'
    widget.bind('<Button-3>', show)
    widget.bind('<Shift-F10>', show)


def install_native_controls(root):
    """Install once per widget, retaining the toolkit class binding chain."""
    for widget in root.winfo_children():
        if not getattr(widget, '_ux_native', False):
            widget._ux_native = True
            install_text_menu(widget)
            if widget.winfo_class() == 'Treeview' and not getattr(widget, '_mapping_tree', False):
                widget.bind('<Control-c>', lambda e: copy_tree(e.widget))
        install_native_controls(widget)


def scroll_content(canvas, body):
    """Wheel/focus scrolling scoped to descendants of a single content panel."""
    tag = 'UXScroll' + str(canvas)
    def wheel(event):
        if event.widget.winfo_class() in ('Treeview', 'Text', 'Listbox', 'TCombobox'):
            return
        delta, num = getattr(event, 'delta', 0), getattr(event, 'num', 0)
        if delta or num in (4, 5):
            canvas.yview_scroll(-3 if delta > 0 or num == 4 else 3, 'units')
            return 'break'
    def reveal(event):
        try:
            relative = event.widget.winfo_rooty() - body.winfo_rooty()
            height = max(1, body.winfo_height())
            top = canvas.canvasy(0)
            bottom = top + canvas.winfo_height()
            end = relative + event.widget.winfo_height()
            if relative < top:
                canvas.yview_moveto(relative / height)
            elif end > bottom:
                canvas.yview_moveto((end - canvas.winfo_height()) / height)
        except tk.TclError:
            pass
    for sequence in ('<MouseWheel>', '<Button-4>', '<Button-5>'):
        canvas.bind_class(tag, sequence, wheel)
    canvas.bind_class(tag, '<FocusIn>', reveal)
    def attach(widget):
        if tag not in widget.bindtags():
            tags = list(widget.bindtags())
            tags.insert(1, tag)
            widget.bindtags(tuple(tags))
        for child in widget.winfo_children():
            attach(child)
    body.bind('<Configure>', lambda e: attach(body), add='+')
    attach(body)
    def cleanup(event):
        if event.widget is canvas:
            for sequence in ('<MouseWheel>', '<Button-4>', '<Button-5>', '<FocusIn>'):
                canvas.unbind_class(tag, sequence)
    canvas.bind('<Destroy>', cleanup, add='+')
