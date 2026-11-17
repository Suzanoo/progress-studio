"""Select stable BOQ catalog rows from one accepted BP export review."""
from dataclasses import dataclass, replace
from pathlib import Path
from tempfile import NamedTemporaryFile
import json

from progress_studio.domain.boq_payment_report import (
    amount_completion_states, known_projected_percentages, known_projected_totals, report_rows,
)
from progress_studio.services.boq_payment_export import digest
from progress_studio.services.payment_plan_export import publish_pair


class BOQReportCatalog:
    def __init__(self, review):
        self.review = review
        self.rows = report_rows(review.projection)
        self.by_key = {r.key: r for r in self.rows}
        if len(self.rows) != len(self.by_key):
            raise ValueError('Duplicate BOQ reporting identity.')
        self.root = self.rows[0]
        self.entries = {e.item.key: e for e in review.projection.items if e.item.role == 'item'}
        groups = {r.wbs: r.key for r in self.rows if r.group and r.wbs}
        self.parents = {}
        self.members = {r.key: [] for r in self.rows}
        for row in self.rows:
            code = row.wbs
            parent = None
            while '.' in code and parent is None:
                code = code.rsplit('.', 1)[0]
                parent = groups.get(code)
            self.parents[row.key] = parent or (self.root.key if row != self.root else None)
            if not row.group:
                key = row.key
                while key is not None:
                    self.members[key].append(row.key)
                    key = self.parents[key]

    def select(self, keys):
        keys = set(keys)
        if not keys or keys - self.by_key.keys():
            raise ValueError('Choose rows from the current BOQ catalog.')
        rows = tuple(r for r in self.rows if r.key in keys)
        covered = {leaf for r in rows for leaf in self.members[r.key]}
        if not covered:
            raise ValueError('The selected groups contain no BOQ items.')
        projection = self.review.projection
        subset = replace(projection, items=tuple(e for k, e in self.entries.items() if k in covered))
        total = report_rows(subset)[0]
        known = known_projected_totals(subset)
        fractions = known_projected_percentages(subset, known)
        flags = amount_completion_states(projection, rows)
        total_flags = amount_completion_states(subset, (total,))[0]
        details = []
        for row, complete in zip(rows, flags):
            parent = self.parents[row.key]
            while parent is not None and parent not in keys:
                parent = self.parents[parent]
            members = tuple(self.entries[k] for k in self.members[row.key])
            partial = known_projected_totals(replace(projection, items=members))
            details.append(Detail(row, self.by_key[parent] if parent else None, complete, partial))
        return SelectedBOQ(self, tuple(details), len(covered), total, total_flags, known, fractions)

    def selection_data(self, keys):
        return {'schema': 'boq-word-selection-v1',
                'revisions': [h for _, h in self.review.sources],
                'rows': [r.key for r in self.rows if r.key in keys]}

    def restore(self, path):
        data = json.loads(Path(path).read_text(encoding='utf-8'))
        if not isinstance(data, dict) or not isinstance(data.get('rows'), list):
            raise ValueError('Invalid BOQ selection file.')
        keys = data['rows']
        if (not all(isinstance(k, str) for k in keys) or len(set(keys)) != len(keys)
                or data != self.selection_data(set(keys))):
            raise ValueError('Selection is incompatible with this reviewed source or catalog.')
        return set(keys)

    def save(self, path, keys):
        target = Path(path)
        if target.suffix.lower() != '.json' or any(target.resolve() == p.resolve()
                or (target.exists() and target.samefile(p)) for p, _ in self.review.sources):
            raise ValueError('Choose a separate JSON selection file.')
        tmp = None
        try:
            with NamedTemporaryFile('w', dir=target.parent, delete=False, encoding='utf-8') as f:
                tmp = Path(f.name)
                json.dump(self.selection_data(keys), f, ensure_ascii=False, indent=2)
            tmp.replace(target)
        finally:
            if tmp:
                tmp.unlink(missing_ok=True)


@dataclass(frozen=True)
class Detail:
    row: object
    ancestor: object
    complete: tuple
    known: object


def display_number(value, percent=False):
    if value is None:
        return 'Unavailable'
    if percent and value != 1 and format(value, '.2%') == '100.00%':
        return '<100%' if value < 1 else '>100%'
    return format(value, '.2%' if percent else ',.2f')


@dataclass(frozen=True)
class ReportPage:
    details: tuple
    totals: tuple
    diagnostics: tuple


@dataclass(frozen=True)
class SelectedBOQ:
    catalog: BOQReportCatalog
    details: tuple
    count: int
    total: object
    complete: tuple
    known: object
    fractions: object
    active_only: bool = True

    @property
    def coverage(self):
        basis = self.catalog.root.amount
        return self.total.amount / basis if basis and self.total.amount is not None else None

    def period_title(self, index):
        projection = self.catalog.review.projection
        period = projection.periods[index]
        date = next((p.payment_date for p in projection.period_sources if p.period_id == period), None)
        return f'{period}  ·  {date.isoformat()}' if date else period

    def page(self, index, cumulative):
        lines = self.lines(index, cumulative)
        visible = []
        diagnostics = []
        for detail in self.details:
            line = next(lines)
            movement = detail.row.values[3][index]
            if self.active_only and movement is None:
                known = display_number(detail.known.period[index])
                diagnostics.append(f'{line[0]} · {detail.row.status} · Known period contribution: {known}')
            elif not self.active_only or movement != 0:
                visible.append(line)
        return ReportPage(tuple(visible), tuple(lines), tuple(diagnostics))

    def lines(self, index, cumulative):
        """Shared preview and Word cells. Raw projection and totals stay immutable."""
        mode = 0 if cumulative else 1

        def cells(row, complete, known, label, state):
            pct, amount = row.values[mode][index], row.values[mode + 2][index]
            repeated = cumulative and index > 0 and complete[index] and complete[index - 1]
            if repeated:
                return (label, 'Still complete', '—', '—')
            if amount is None and known is not None:
                state += ' · Known ' + display_number(known)
            return (label, state, display_number(pct, True), display_number(amount))

        for detail in self.details:
            row = detail.row
            label = (row.wbs + '  ' + row.name).strip()
            if detail.ancestor:
                label += ' · Included in ' + (detail.ancestor.wbs or detail.ancestor.name)
            entry = self.catalog.entries.get(row.key)
            if entry:
                period = entry.periods[index]
                state = getattr(entry, 'origin', '') or period.status
                if state == 'Complete':
                    state = 'Mapped · Full data'
                if getattr(entry, 'origin', ''):
                    state += ' · ' + period.state
                if period.status == 'PartialMapping':
                    state += ' ' + format(period.mapping_coverage, '.1%')
                if entry.item.amount is None or entry.item.amount < 0:
                    state += ' · ' + entry.item.amount_status
            else:
                state = 'Group'
            known = (detail.known.cumulative if cumulative else detail.known.period)[index]
            yield cells(row, detail.complete, known, label, state)
        yield cells(self.total, self.complete, None, 'Complete selected scope',
                    'Incomplete' if self.total.values[mode + 2][index] is None else 'Complete data')
        # A known subtotal alone cannot establish completion. Only the independent
        # scope predicate and unchanged, fully determined headline justify suppression.
        if cumulative and index > 0 and self.complete[index] and self.complete[index-1]:
            if all(self.fractions.cumulative[i] == 1 and self.known.cumulative[i] == self.total.amount
                   for i in (index-1, index)):
                yield ('Known contribution', 'Still complete', '—', '—')
                return
        yield ('Known contribution', 'Selected basis',
               display_number((self.fractions.cumulative if cumulative else self.fractions.period)[index], True),
               display_number((self.known.cumulative if cumulative else self.known.period)[index]))


def export_boq_reports(plan, parent, name, *, acknowledge_revision=False, cancelled=lambda: False):
    from progress_studio.infrastructure.word.boq_plan import render
    review = plan.catalog.review

    def validate():
        if review.revision_required and not acknowledge_revision:
            raise ValueError('Acknowledge the reviewed source revision.')
        if any(digest(path) != expected for path, expected in review.sources):
            raise ValueError('A source has changed. Review the BOQ sources again.')

    def write(path, cumulative, check):
        render(path, plan, cumulative=cumulative, check_cancel=check)

    return publish_pair(parent, name, 'BOQ_PAYMENT_PLAN_REPORT', write, validate, cancelled)
