"""XW-1 reviewed, all-or-nothing Workbook -> P6 monetary export. No Desktop API."""
from dataclasses import dataclass
from decimal import Decimal
import hashlib
import os
from pathlib import Path
import tempfile
from typing import Callable

from progress_studio.domain.amount_field import AmountField
from progress_studio.domain.xml_writeback import ActivityAmountChange, exact_total, monetary_decimal
from progress_studio.infrastructure.excel.xml_writeback_reader import read_workbook_amounts
from progress_studio.infrastructure.schedule_xml.p6_amount_writer import P6AmountDocument


@dataclass(frozen=True)
class WritebackRequest:
    workbook: Path
    target_xml: Path
    field_identity: str | None = None
    create_progress_studio_amount: bool = False
    monetary_confirmed: bool = False
    recalculated_and_saved: bool = False


@dataclass(frozen=True)
class WritebackAssessment:
    request: WritebackRequest
    workbook_sha256: str
    xml_sha256: str
    project_id: str
    field: AmountField | None
    creates_field: bool
    available_fields: tuple[AmountField, ...]
    rows: tuple[ActivityAmountChange, ...]
    source_count: int
    target_only_ids: tuple[str, ...]
    formula_count: int
    zero_count: int
    source_total: Decimal
    write_total: Decimal
    before_known_total: Decimal
    before_missing_count: int
    errors: tuple[str, ...]
    warnings: tuple[str, ...]

    @property
    def valid(self):
        return not self.errors

    @property
    def replacement_count(self):
        return sum(r.status == 'replace' for r in self.rows)

    @property
    def unchanged_count(self):
        return sum(r.status == 'unchanged' for r in self.rows)

    @property
    def preserved_milestone_count(self):
        return sum(r.status == 'preserve_milestone' for r in self.rows)


@dataclass(frozen=True)
class WritebackResult:
    output: Path
    sha256: str
    replacement_count: int
    warnings: tuple[str, ...]


class WritebackCancelled(ValueError):
    pass


def _sha(data):
    return hashlib.sha256(data).hexdigest()


class XmlWritebackService:
    def assess(self, request: WritebackRequest) -> WritebackAssessment:
        request = WritebackRequest(Path(request.workbook).resolve(), Path(request.target_xml).resolve(),
                                   request.field_identity, request.create_progress_studio_amount,
                                   request.monetary_confirmed, request.recalculated_and_saved)
        errors, warnings, changes, before = [], [], [], []
        workbook_digest = xml_digest = project_id = ''
        selected = None
        create = False
        available, target_only = (), ()
        source_count = formula_count = zeros = missing = 0
        source_total = write_total = Decimal(0)
        if request.monetary_confirmed is not True:
            errors.append('Explicitly confirm that current main Activity Amounts are monetary values.')
        try:
            workbook_data = request.workbook.read_bytes()
            xml_data = request.target_xml.read_bytes()
            workbook_digest, xml_digest = _sha(workbook_data), _sha(xml_data)
            source = read_workbook_amounts(workbook_data)
            errors.extend(source.errors)
            source_count = len(source.rows)
            formula_count = sum(r.formula for r in source.rows)
            if formula_count:
                warnings.append('Saved numeric formula results are used. Cache freshness cannot be independently guaranteed; recalculate in Excel and Save before assessment.')
                if request.recalculated_and_saved is not True:
                    errors.append('Confirm Excel recalculation and Save before using formula Amounts.')
            source_total = exact_total(r.value for r in source.rows if r.value is not None)
            target = P6AmountDocument(xml_data)
            project_id = target.text(target.project, 'Id')
            available = target.fields()
            selected, create = target.resolve_field(request.field_identity, request.create_progress_studio_amount)
            ids = {r.activity_id for r in source.rows}
            target_only = tuple(aid for aid in target.activities if aid not in ids)
            for row in source.rows:
                try:
                    if row.activity_id not in target.activities:
                        raise ValueError('No exact case-preserved Activity ID match in target Project.')
                    milestone = target.milestone(row.activity_id)
                    if row.milestone is None:
                        raise ValueError('Unknown source milestone classification; explicit workbook milestone evidence is required.')
                    if row.milestone != milestone:
                        raise ValueError('Conflicting Workbook/XML milestone classifications.')
                    _, _, raw = target.value_node(row.activity_id, selected.identity)
                    if row.value is None:
                        continue  # Already a blocking source error, never partial acceptance.
                    name = target.text(target.activities[row.activity_id], 'Name')
                    if milestone:
                        changes.append(ActivityAmountChange(row.activity_id, name, row.row, row.value,
                                                           raw, raw, 'preserve_milestone'))
                        continue
                    old = None
                    if raw is not None:
                        old = monetary_decimal(raw)
                        before.append(old)
                    else:
                        missing += 1
                    status = 'unchanged' if old == row.value else 'replace'
                    changes.append(ActivityAmountChange(row.activity_id, name, row.row, row.value,
                                                       raw, raw if status == 'unchanged' else row.raw, status))
                    zeros += row.value == 0
                except ValueError as exc:
                    errors.append(f'{row.activity_id} (main row {row.row}): {exc}')
            write_total = exact_total(r.source_amount for r in changes if r.status != 'preserve_milestone')
            if any(r.status == 'preserve_milestone' for r in changes):
                warnings.append('Milestone monetary values are preserved verbatim in XML; main Progress Weights do not replace them. Totals for writes exclude milestones.')
            if create:
                warnings.append(f'Export will create Activity Cost UDF {selected.label}; review this new identity.')
        except (ValueError, OSError) as exc:
            errors.append(str(exc))
        return WritebackAssessment(request, workbook_digest, xml_digest, project_id, selected, create,
                                   available, tuple(changes), source_count, target_only, formula_count,
                                   zeros, source_total, write_total, exact_total(before), missing,
                                   tuple(errors), tuple(warnings))

    @staticmethod
    def _check_sources(assessment):
        if (_sha(assessment.request.workbook.read_bytes()) != assessment.workbook_sha256
                or _sha(assessment.request.target_xml.read_bytes()) != assessment.xml_sha256):
            raise ValueError('Workbook or XML changed since assessment; review again.')

    def export(self, assessment: WritebackAssessment, output: Path, *, reviewed: bool = False,
               cancelled: Callable[[], bool] | None = None) -> WritebackResult:
        if reviewed is not True:
            raise ValueError('Review and explicitly approve the assessment before export.')
        if not assessment.valid:
            raise ValueError('Invalid assessment; no output: ' + '\n'.join(assessment.errors))

        def checkpoint():
            if cancelled is not None and cancelled():
                raise WritebackCancelled('XML writeback cancelled; no output published.')

        checkpoint()
        self._check_sources(assessment)
        # Do not trust edited/reconstructed preview fields or caller-supplied bytes.
        if self.assess(assessment.request) != assessment:
            raise ValueError('Assessment changed or was modified; review again.')
        destination = Path(output).absolute()
        if destination.resolve() in (assessment.request.workbook, assessment.request.target_xml):
            raise ValueError('Output must be a new file, separate from both inputs.')
        if os.path.lexists(destination):
            raise FileExistsError('Output already exists; choose a new filename.')
        data = assessment.request.target_xml.read_bytes()
        if _sha(data) != assessment.xml_sha256:
            raise ValueError('XML changed since assessment; review again.')
        payload = P6AmountDocument(data).render(assessment.rows, assessment.field, assessment.creates_field)
        checkpoint()
        stage = None
        published = False
        warnings = list(assessment.warnings)
        try:
            fd, name = tempfile.mkstemp(prefix='.progress-studio-xml-', suffix='.tmp', dir=destination.parent)
            stage = Path(name)
            try:
                stream = os.fdopen(fd, 'wb')
            except BaseException:
                os.close(fd)
                raise
            with stream:
                stream.write(payload)
                stream.flush()
                os.fsync(stream.fileno())
            if stage.read_bytes() != payload:
                raise OSError('Staged XML integrity check failed.')
            checkpoint()
            self._check_sources(assessment)
            # Same-directory hard-link publication is atomic and fails if the
            # destination appeared meanwhile. Never use replace/overwrite.
            os.link(stage, destination)
            published = True
        finally:
            if stage is not None:
                try:
                    stage.unlink(missing_ok=True)
                except OSError:
                    if published:
                        warnings.append(f'Output completed; temporary file cleanup required: {stage}')
                    # Preserve original publication/write exception on failure.
        return WritebackResult(destination, _sha(payload), assessment.replacement_count, tuple(warnings))
