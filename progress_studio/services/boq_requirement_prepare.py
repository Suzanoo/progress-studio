"""Explicit prepare/reconcile into a new source copy. No source mutation."""
from pathlib import Path
from dataclasses import replace
import os
import tempfile
from zipfile import ZipFile
from progress_studio.services.boq_payment_projection import BOQPaymentProjectionService
from progress_studio.services.boq_payment_export import ExportReview, digest
from progress_studio.domain.boq_payment_projection import Diagnostic
from progress_studio.infrastructure.excel.boq_requirement_input import read_input, render_input, eligible
from progress_studio.infrastructure.excel.boq_input_package import merge_input


class BOQRequirementPrepareService:
    def review(self, session, payment, boq=None, *, cancelled=lambda:False):
        paths=tuple(Path(p).resolve() for p in (session,payment,boq) if p)
        before=tuple((p,digest(p)) for p in paths)
        if cancelled():raise InterruptedError('Prepare cancelled')
        projection=BOQPaymentProjectionService().calculate(paths[0],paths[1],accept_revision=True,
            boq_path=paths[2] if len(paths)>2 else None)
        records=read_input(paths[1],projection,reconcile=True)
        if cancelled():raise InterruptedError('Prepare cancelled')
        if any(digest(p)!=h for p,h in before):raise ValueError('Source changed during Prepare review')
        projection=replace(projection,diagnostics=projection.diagnostics+(Diagnostic('PrepareBOQ', 'Planned requirements',
            f'{len(eligible(projection))} wholly unmapped rows; {len(records or ())} existing rows preserved; references refresh on Prepare'),))
        return ExportReview(projection,before,any(d.code=='SourceRevision' for d in projection.diagnostics))

    def export(self, review, destination, *, acknowledge_revision=False, cancelled=lambda:False):
        destination=Path(destination).expanduser().resolve()
        if destination.suffix.lower()!='.xlsx':raise ValueError('Prepare output must be a new .xlsx')
        for path,_ in review.sources:
            if str(path).casefold()==str(destination).casefold() or (destination.exists() and os.path.samefile(path,destination)):
                raise ValueError('Prepare output aliases source')
        if destination.exists():raise FileExistsError('Choose a new output; existing files are not overwritten')
        if review.revision_required and not acknowledge_revision:raise ValueError('Acknowledge source revision before Prepare')
        def check():
            if cancelled():raise InterruptedError('Prepare cancelled')
            if any(digest(p)!=h for p,h in review.sources):raise ValueError('Source changed since Prepare review')
        check();source=review.sources[1][0];records=read_input(source,review.projection,reconcile=True)
        temp=None
        try:
            fd,name=tempfile.mkstemp(prefix='.bp4-',suffix='.xlsx',dir=destination.parent);os.close(fd);temp=Path(name)
            wb=render_input(review.projection,records)
            try:wb.save(temp)
            finally:wb.close()
            check();merge_input(source,temp)
            with ZipFile(temp) as z:
                if z.testzip():raise ValueError('Invalid prepared workbook ZIP')
            verified=read_input(temp,review.projection)
            if {r.key:r.values for r in records or ()}.items() <= {r.key:r.values for r in verified}.items():
                check();os.link(temp,destination)
            else:raise ValueError('Prepare verification lost manual values')
            return destination
        finally:
            if temp is not None:temp.unlink(missing_ok=True)
