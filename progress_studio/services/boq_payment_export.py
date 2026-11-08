"""Reviewed BP-2 export. Source files are never opened for writing."""
from dataclasses import dataclass, replace
from pathlib import Path
import hashlib
import os
import tempfile
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from progress_studio.services.boq_payment_projection import BOQPaymentProjectionService
from progress_studio.infrastructure.excel.boq_payment_workbook import render
from progress_studio.infrastructure.excel.payment_report_snapshot import inspect_snapshots, append_payment_snapshots
from progress_studio.infrastructure.excel.payment_workbook import PaymentWorkbookSnapshotter
from progress_studio.domain.boq_payment_projection import Diagnostic
from progress_studio.domain.boq_manual_requirements import compose
from progress_studio.infrastructure.excel.boq_requirement_input import read_input


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


@dataclass(frozen=True)
class ExportReview:
    projection: object
    sources: tuple[tuple[Path,str], ...]
    revision_required: bool


class BOQPaymentExportService:
    def review(self, session, payment, boq=None, *, cancelled=lambda:False):
        paths=tuple(Path(p).resolve() for p in (session,payment,boq) if p)
        before=tuple((p,digest(p)) for p in paths)
        if cancelled():raise InterruptedError('BOQ export cancelled')
        # Review may read a revised source; publication separately requires its
        # explicit acknowledgement. BP-1 compatibility/relink rules still apply.
        snapshot_info=inspect_snapshots(paths[1], cancelled)
        result=BOQPaymentProjectionService().calculate(paths[0],paths[1],accept_revision=True,boq_path=paths[2] if len(paths)>2 else None)
        result=compose(result,read_input(paths[1],result))
        if cancelled():raise InterruptedError('BOQ export cancelled')
        if any(digest(p)!=h for p,h in before):raise ValueError('Source changed during review')
        diagnostics = tuple(Diagnostic('PaymentSnapshot', i.sheet,
            f'Saved-value snapshot: {i.cells} stored cells; {i.formulas} cached formulas; {i.errors} source errors. '
            'Cache freshness is unverified. Recalculate/save source before export; no Payment calculation is run.') for i in snapshot_info)
        result=replace(result, diagnostics=result.diagnostics+diagnostics)
        return ExportReview(result,before,any(d.code=='SourceRevision' for d in result.diagnostics))

    def export(self, review, destination, *, acknowledge_revision=False, cancelled=lambda:False):
        destination=Path(destination).expanduser().resolve()
        if destination.suffix.lower()!='.xlsx':raise ValueError('Output must be a new .xlsx file')
        for source,_ in review.sources:
            if str(source).casefold()==str(destination).casefold() or (destination.exists() and os.path.samefile(source,destination)):
                raise ValueError('Output aliases an input source')
        if destination.exists():raise FileExistsError('Choose a new output filename; existing files are not overwritten')
        if review.revision_required and not acknowledge_revision:raise ValueError('Acknowledge the reviewed source revision before export')
        def check():
            if cancelled():raise InterruptedError('BOQ export cancelled')
            if any(digest(p)!=h for p,h in review.sources):raise ValueError('Source changed since review; review again')
        check();temp=None
        try:
            wb=render(review.projection,cancelled)
            try:
                if len(review.sources)<2:raise ValueError('Reviewed Payment source is required for snapshots')
                append_payment_snapshots(wb, review.sources[1][0], review.sources[1][1], cancelled)
                check()
                fd,name=tempfile.mkstemp(prefix='.bp2-',suffix='.xlsx',dir=destination.parent);os.close(fd);temp=Path(name)
                wb.save(temp)
            finally:wb.close()
            self.validate(temp)
            check()
            # Atomic no-clobber publication, including a destination appearing
            # after validation. A winning publication is success, not cancellation.
            os.link(temp,destination)
            return destination
        finally:
            if temp is not None:temp.unlink(missing_ok=True)

    @staticmethod
    def validate(path):
        ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        with ZipFile(path) as z:
            if z.testzip():raise ValueError('Invalid output ZIP')
            names=z.namelist()
            if any('vbaproject' in n.lower() or 'activex' in n.lower() or 'externallinks' in n.lower() for n in names):raise ValueError('Forbidden output dependency')
            root=ET.fromstring(z.read('xl/workbook.xml'));sheets=root.find('s:sheets',ns)
            if [x.attrib['name'] for x in sheets]!=['Percentage','Amount','Trace','Payment Input','Payment','_BP_Data']:raise ValueError('Invalid report schema')
            if sheets[-1].attrib.get('state')!='hidden':raise ValueError('Snapshot must be hidden')
            sheet_map=PaymentWorkbookSnapshotter._sheet_map(z)
            data=ET.fromstring(z.read(sheet_map['_BP_Data']))
            if data.find('s:sheetProtection',ns) is None or data.findall('.//s:f',ns):raise ValueError('Invalid snapshot protection/formulas')
            for name in ('Payment Input','Payment'):
                snapshot=ET.fromstring(z.read(sheet_map[name]))
                if snapshot.findall('.//s:f',ns) or snapshot.find('s:sheetProtection',ns) is None:
                    raise ValueError('Invalid Payment snapshot values/protection')
                if snapshot.find('s:dataValidations',ns) is not None or snapshot.find('s:hyperlinks',ns) is not None:
                    raise ValueError('Live input controls/links in Payment snapshot')
            for n in names:
                if n.startswith('xl/charts/') and n.endswith('.xml'):
                    if any(e.tag.rsplit('}',1)[-1]=='f' for e in ET.fromstring(z.read(n)).iter()):
                        raise ValueError('Chart dependency in Payment snapshot')
                if n.endswith('.rels'):
                    for rel in ET.fromstring(z.read(n)):
                        if rel.attrib.get('TargetMode')=='External':raise ValueError('External relationship in report')
