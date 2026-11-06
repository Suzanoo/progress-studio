"""BP-1 source association is read-only and distinct from session relinking."""
from dataclasses import replace
from pathlib import Path
from tempfile import TemporaryDirectory
import hashlib
from openpyxl import load_workbook
from progress_studio.domain.boq_payment_projection import project, Diagnostic, number
from progress_studio.infrastructure.excel.boq_projection_catalog import read_catalog
from progress_studio.infrastructure.excel.mapping_reader import ProgressActivityReader
from progress_studio.infrastructure.excel.payment_input_reader import PaymentInputSparseReader
from progress_studio.infrastructure.session.mapping_session_repository import MappingSessionRepository, materialize_workbook_snapshot
from progress_studio.services.mapping_store import MappingStore


def _main(path):
    wb=load_workbook(path,read_only=True,data_only=False)
    try:
        if 'main' not in wb.sheetnames:raise ValueError('main sheet missing')
        rows=iter(wb['main'].values);headers=None
        for _,row in zip(range(30),rows):
            h={str(v or '').strip().lower():i for i,v in enumerate(row)}
            if {'row type','activity id','p/a','description'} <= h.keys():headers=h;break
        if headers is None:raise ValueError('main headers missing')
        result={}
        for row in rows:
            get=lambda key:row[headers[key]] if key in headers and headers[key]<len(row) else None
            if str(get('row type') or '').lower()!='activity' or str(get('p/a') or '').upper()!='P':continue
            identity=str(get('activity id') or '').strip()
            if not identity or identity in result:raise ValueError('Blank or duplicate main Activity ID')
            result[identity]={k:get(k) for k in ('description','wbs','plan start','plan finish','amount')}
        return result
    finally:wb.close()


class BOQPaymentProjectionService:
    def calculate(self, session_path: Path, payment_path: Path, *, accept_revision=False, boq_path: Path | None=None):
        """Use selected current Payment with saved allocations, never relink.

        A changed Payment workbook requires explicit acknowledgement. The result
        always retains hashes and revision diagnostics even when acknowledged.
        A supplied BOQ must pass the existing session source validation contract.
        Embedded snapshots are materialized only into a disposable directory.
        """
        session_path=Path(session_path);payment_path=Path(payment_path)
        paths=[session_path,payment_path]+([Path(boq_path)] if boq_path else [])
        hashes={str(p.resolve()):hashlib.sha256(p.read_bytes()).hexdigest() for p in paths}
        repository=MappingSessionRepository();session=repository.load(session_path)
        if not session.progress_snapshot or not session.boq_snapshot:
            raise ValueError('Verified embedded source snapshots are required; use a supported saved project session')
        with TemporaryDirectory(prefix='bp1-projection-') as temp:
            source=materialize_workbook_snapshot(session.progress_snapshot,Path(temp))
            catalog_source=materialize_workbook_snapshot(session.boq_snapshot,Path(temp))
            if session.progress_snapshot.sha256!=session.progress.sha256 or session.boq_snapshot.sha256!=session.boq.sha256:
                raise ValueError('Session snapshot and fingerprint disagree')
            repository.validate_workbook(session.progress,source)
            repository.validate_workbook(session.boq,catalog_source)
            if boq_path:catalog_source=repository.validate_workbook(session.boq,Path(boq_path))
            catalog=read_catalog(catalog_source,session.boq_sheet)
            original=_main(source);current=_main(payment_path)
            store=MappingStore();store.load_activities(ProgressActivityReader().read(source))
            if session.supplemental_activities or session.supplemental_wbs:
                raise ValueError('Supplemental session records require explicit compatibility review in BP-1')
            if session.working_tree_nodes:store.restore_working_tree(session.working_tree_nodes)
            expected=store.activities_by_id
            for a in session.allocations:
                if a.activity_id not in expected or a.activity_id not in current:
                    raise ValueError('Allocation Activity not present in both selected sources')
                if a.activity_id not in original:
                    raise ValueError('Mapped Activity missing from saved source main')
                if str(original[a.activity_id]['description'] or '').strip()!=str(current[a.activity_id]['description'] or '').strip():
                    raise ValueError('Mapped Activity description changed; source association is ambiguous')
            diagnostics=[]
            for identity in sorted(set(expected) & set(original)):
                if expected[identity].description.strip()!=str(original[identity]['description'] or '').strip():
                    diagnostics.append(Diagnostic('MappingDescriptionRevision',identity,'Saved mapping label differs from saved main; identity matched against original main'))
            current_hash=hashes[str(payment_path.resolve())]
            if current_hash!=session.progress.sha256:
                if not accept_revision:raise ValueError('Current Payment differs from session Progress; explicitly accept revision diagnostics')
                diagnostics.append(Diagnostic('SourceRevision','Progress',f'Saved={session.progress.sha256}; selected={current_hash}; not a relink'))
            for identity in sorted(set(original)&set(current)):
                changes=[k for k in original[identity] if k!='amount' and original[identity][k]!=current[identity][k]]
                if changes:diagnostics.append(Diagnostic('ActivityRevision',identity,', '.join(changes)))
            for identity in sorted(set(current)-set(expected)):
                diagnostics.append(Diagnostic('AdditionalActivity',identity,'No mapping inferred'))
            payment=PaymentInputSparseReader().read(payment_path)
            result=project(catalog,session.allocations,payment,tuple(current))
            by_key={x.key:x for x in catalog};allocated={};unknown=set()
            for a in session.allocations:
                item=by_key.get(a.boq_key)
                if item is None:raise ValueError('Unknown BOQ reference')
                if item.amount is None or item.amount<0:unknown.add(a.activity_id);continue
                allocated[a.activity_id]=allocated.get(a.activity_id,number(0))+item.amount*number(a.share_percent)/100
            for identity in sorted({a.activity_id for a in session.allocations}):
                if identity in unknown:
                    diagnostics.append(Diagnostic('ReconciliationUnavailable',identity,'Some allocated BOQ Amounts unavailable'));continue
                amount=allocated.get(identity,number(0))
                try:actual=number(current[identity]['amount'])
                except ValueError:
                    diagnostics.append(Diagnostic('ReconciliationUnavailable',identity,'Current Activity Amount missing/invalid/formula'));continue
                difference=actual-amount
                if abs(difference)>number('.01'):
                    diagnostics.append(Diagnostic('ActivityAmountDifference',identity,f'Allocated={amount}; current={actual}; difference={difference}; informational, not redistributed'))
            for p in paths:
                if hashlib.sha256(p.read_bytes()).hexdigest()!=hashes[str(p.resolve())]:raise ValueError('Source changed during calculation')
            provenance=tuple(sorted(hashes.items()))+(('saved_progress_sha256',session.progress.sha256),('saved_boq_sha256',session.boq.sha256),('mapping_saved_at',session.saved_at),('boq_sheet',session.boq_sheet))
            return replace(result,diagnostics=result.diagnostics+tuple(diagnostics),provenance=provenance)
