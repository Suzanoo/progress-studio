"""Reviewed BP-2 export. Source files are never opened for writing."""
from dataclasses import dataclass
from pathlib import Path
import hashlib
import os
import tempfile
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from progress_studio.services.boq_payment_projection import BOQPaymentProjectionService
from progress_studio.infrastructure.excel.boq_payment_workbook import render


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
        result=BOQPaymentProjectionService().calculate(paths[0],paths[1],accept_revision=True,boq_path=paths[2] if len(paths)>2 else None)
        if cancelled():raise InterruptedError('BOQ export cancelled')
        if any(digest(p)!=h for p,h in before):raise ValueError('Source changed during review')
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
            check()
            fd,name=tempfile.mkstemp(prefix='.bp2-',suffix='.xlsx',dir=destination.parent);os.close(fd);temp=Path(name)
            try:wb.save(temp)
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
            if [x.attrib['name'] for x in sheets]!=['Percentage','Amount','Trace','_BP_Data']:raise ValueError('Invalid report schema')
            if sheets[-1].attrib.get('state')!='hidden':raise ValueError('Snapshot must be hidden')
            data=ET.fromstring(z.read('xl/worksheets/sheet4.xml'))
            if data.find('s:sheetProtection',ns) is None or data.findall('.//s:f',ns):raise ValueError('Invalid snapshot protection/formulas')
            for n in names:
                if n.endswith('.rels'):
                    for rel in ET.fromstring(z.read(n)):
                        if rel.attrib.get('TargetMode')=='External':raise ValueError('External relationship in report')
