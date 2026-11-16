"""Publish two editable reports together by renaming a staged sibling directory."""
import hashlib
import os
import re
from pathlib import Path
from shutil import rmtree
from tempfile import mkdtemp
from zipfile import ZipFile

from progress_studio.infrastructure.word.payment_plan import render


class ExportCancelled(Exception):
    pass


def export_payment_plans(source, selection, workbook, parent, name, *, cancelled=lambda: False,
                         progress=lambda stage: None, reviewed=None):
    if reviewed is not None:
        from dataclasses import replace
        if selection is not reviewed.selection:
            raise ValueError("Review the adjusted selection before export.")
        source = replace(source, diagnostics=(*source.diagnostics, f"Adjusted plan: {reviewed.path.name} | SHA-256 {reviewed.revision}", *reviewed.issues))
    if not selection.rows or selection.revision != source.revision:
        raise ValueError('Select reporting rows from the current workbook.')
    if (not name or name != name.strip() or name.endswith('.') or len(name) > 80
            or re.search(r'[<>:"/\\|?*\x00-\x1f]', name)
            or name.split('.')[0].upper() in {'CON','NUL','PRN','AUX', *(f'COM{i}' for i in range(1,10)),
                                                *(f'LPT{i}' for i in range(1,10))}):
        raise ValueError('Use a short, valid folder name without path separators.')
    parent = Path(parent).resolve(strict=True)
    target = parent / name
    def check():
        if cancelled():
            raise ExportCancelled()
    def validate():
        check()
        if reviewed is not None:
            reviewed.validate()
        if os.path.lexists(target):
            raise FileExistsError(f'Choose a new report folder. Already exists: {target}')
        if hashlib.sha256(Path(workbook).read_bytes()).hexdigest() != source.revision:
            raise ValueError('The source workbook changed. Reopen it and review the selection.')
    validate()
    stage = Path(mkdtemp(prefix='.payment-plan-', dir=parent))
    try:
        for kind in ('PERIOD', 'CUMULATIVE'):
            progress(kind)
            check()
            path = stage / f'{name}_{kind}_PAYMENT_PLAN_REPORT.docx'
            render(path, source, selection, cumulative=kind == 'CUMULATIVE', check_cancel=check)
            with ZipFile(path) as package:
                if 'word/document.xml' not in package.namelist() or package.testzip() is not None:
                    raise ValueError(f'Invalid generated report: {path.name}')
        validate()
        stage.rename(target)
        return target
    finally:
        if stage.exists():
            rmtree(stage)
