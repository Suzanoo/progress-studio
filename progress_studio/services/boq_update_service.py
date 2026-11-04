"""BU-1: stage a revision without mutating the live project; persist before publish."""
from copy import deepcopy
from dataclasses import dataclass, replace
from datetime import datetime, timezone
import hashlib
import math
from pathlib import Path
from tempfile import TemporaryDirectory
from uuid import uuid4

from progress_studio.domain.mapping_models import BOQRow
from progress_studio.domain.mapping_session import MappingSessionData
from progress_studio.infrastructure.excel.boq_revision_reader import read_boq_revision
from progress_studio.infrastructure.session.mapping_session_repository import (
    MappingSessionRepository, materialize_workbook_snapshot,
)
from progress_studio.services.mapping_store import MappingStore


def _digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


@dataclass(frozen=True)
class AmountChange:
    identity: str
    description: str
    before: float
    after: float


@dataclass(frozen=True)
class BOQUpdatePreview:
    original: MappingSessionData
    updated: MappingSessionData
    project_path: Path
    project_digest: str
    candidate_path: Path
    candidate_digest: str
    old_items: tuple[BOQRow, ...]
    new_items: tuple[BOQRow, ...]
    item_changes: tuple[AmountChange, ...]
    activity_changes: tuple[AmountChange, ...]
    old_formula_count: int
    new_formula_count: int
    before_total: float
    after_total: float
    before_mapped: float
    after_mapped: float
    mapped_count: int


class BOQUpdateService:
    def __init__(self, repository: MappingSessionRepository | None = None):
        self.repository = repository or MappingSessionRepository()

    def preview(self, session: MappingSessionData, store: MappingStore,
                project_path: Path, candidate_path: Path) -> BOQUpdatePreview:
        project_path, candidate_path = Path(project_path), Path(candidate_path)
        project_digest = _digest(project_path)
        # Take a byte snapshot once. All parsing and the replacement fingerprint
        # refer to these bytes, even if Excel is open elsewhere.
        snapshot = self.repository.snapshot(candidate_path)
        with TemporaryDirectory(prefix='progress-studio-boq-') as directory:
            temp = Path(directory)
            if session.boq_snapshot:
                old = materialize_workbook_snapshot(session.boq_snapshot, temp)
                self.repository.validate_workbook(session.boq, old)
            else:
                old = self.repository.validate_workbook(session.boq)
                old = materialize_workbook_snapshot(self.repository.snapshot(old), temp)
            new = materialize_workbook_snapshot(snapshot, temp)
            before = read_boq_revision(old, session.boq_sheet)
            after = read_boq_revision(new, session.boq_sheet)
            if before.structure != after.structure:
                raise ValueError('Only Amount changes are supported. BOQ rows, columns, identity, descriptions, units, quantities and hierarchy must remain unchanged.')
            old_amounts, new_amounts = dict(before.amounts), dict(after.amounts)
            if old_amounts.keys() != new_amounts.keys():
                raise ValueError('BOQ row insertion or deletion is outside Amount-only update.')
            for row, value in old_amounts.items():
                other = new_amounts[row]
                if value != other and not (value is not None and value > 0 and other is not None and other > 0):
                    raise ValueError(f'BOQ row {row}: eligibility transition is unsupported; both old and new Amount must be positive.')
            if [replace(r, amount=0) for r in before.items] != [replace(r, amount=0) for r in after.items]:
                raise ValueError('BOQ item identity changed; no automatic remapping is allowed.')
            if tuple(store.boq_by_id[k] for k in store.boq_order) != before.items:
                raise ValueError('Loaded BOQ differs from the verified project snapshot. Reopen the project before updating.')
            current = self._stage(store, session, before.items)
            staged = self._stage(store, session, after.items)
            if not all(math.isfinite(v) for v in (staged.total_amount, staged.mapped_amount)):
                raise ValueError('BOQ totals exceed the supported numeric range.')
            # Semantic workbook identity ignores formula caches. Pin reviewed
            # BOQ bytes so changed caches trigger verified snapshot fallback.
            fingerprint = replace(self.repository.fingerprint(new),
                                  path=str(candidate_path.resolve()), filename=candidate_path.name,
                                  semantic_sha256="", identity_kind="binary-sha256")
        if _digest(candidate_path) != snapshot.sha256:
            raise ValueError('BOQ changed during review preparation. Recalculate, Save and review again.')
        changes = tuple(AmountChange(a.key, a.description, a.amount, b.amount)
                        for a, b in zip(before.items, after.items) if a.amount != b.amount)
        activity_changes = tuple(
            AmountChange(key, current.activities_by_id[key].description,
                         current.activity_amount(key), staged.activity_amount(key))
            for key in current.activity_order
            if current.activity_amount(key) != staged.activity_amount(key))
        updated = replace(session, boq=fingerprint, boq_snapshot=snapshot,
                          saved_at=datetime.now(timezone.utc).isoformat())
        return BOQUpdatePreview(session, updated, project_path, project_digest,
                                candidate_path, snapshot.sha256, before.items, after.items,
                                changes, activity_changes, len(before.formula_rows), len(after.formula_rows),
                                current.total_amount, staged.total_amount, current.mapped_amount,
                                staged.mapped_amount, current.mapped_item_count)

    @staticmethod
    def _stage(store: MappingStore, session: MappingSessionData, items: tuple[BOQRow, ...]) -> MappingStore:
        if (tuple(store.allocation_records()) != session.allocations
                or tuple(store.working_tree_nodes()) != session.working_tree_nodes):
            raise ValueError('Mapping changed since review. Review the BOQ update again.')
        staged = deepcopy(store)
        staged.load_boq(list(items))
        staged.restore_allocations(list(session.allocations))
        return staged

    def apply(self, preview: BOQUpdatePreview, current: MappingSessionData,
              store: MappingStore) -> tuple[MappingStore, Path | None, Path | None]:
        if replace(current, saved_at=preview.original.saved_at) != preview.original:
            raise ValueError('Project changed since review. Review the BOQ update again.')
        if _digest(preview.candidate_path) != preview.candidate_digest:
            raise ValueError('BOQ changed since review. Recalculate, Save and review again.')
        if _digest(preview.project_path) != preview.project_digest:
            raise ValueError('Project file changed since review. Reopen it before updating.')
        if tuple(store.boq_by_id[k] for k in store.boq_order) != preview.old_items:
            raise ValueError('Loaded BOQ changed since review.')
        staged = self._stage(store, current, preview.new_items)
        if not preview.item_changes:
            return store, None, None
        # Materialize the exact reviewed bytes before committing. Publication
        # never depends on a source workbook that can change after confirmation.
        materialized = self.repository.materialize_snapshot(preview.updated.boq_snapshot)
        backup = preview.project_path.with_name(
            f'{preview.project_path.stem}.before-boq-update-{uuid4().hex[:12]}.progressstudio')
        self.repository.save(backup, current)
        # Guard stale writers again after writing the recovery copy.
        if _digest(preview.project_path) != preview.project_digest:
            raise ValueError('Project file changed during update. Original project retained; reopen it.')
        self.repository.save(preview.project_path, preview.updated)
        return staged, backup, materialized
