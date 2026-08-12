
from __future__ import annotations

from datetime import date, datetime

from progress_studio.domain.activity_table import ActivityTableModel, ActivityTableRow
from progress_studio.domain.main_dataset import MainDataset, MainRow


def _row_identity(row: MainRow) -> tuple[str, str, str]:
    kind = row.row_type.strip().lower()
    if kind == "activity":
        return kind, row.activity_id.strip(), row.description.strip()
    return kind, row.wbs.strip(), row.description.strip()


def _progress(dataset: MainDataset, row: MainRow | None, cutoff: date | None) -> float:
    if row is None:
        return 0.0
    total = 0.0
    period_dates = {p.column: p.reporting_date for p in dataset.periods}
    for col, value in row.period_values:
        if value is None:
            continue
        reporting = period_dates.get(col)
        if cutoff is not None and reporting is not None and reporting.date() > cutoff:
            continue
        total += float(value)
    return total


def _outline_level(row: MainRow) -> int:
    kind = row.row_type.strip().lower()
    if kind == "project summary":
        return 0
    if row.outline_level is not None:
        return max(0, min(int(row.outline_level), 7))
    parts = [p for p in row.wbs.split(".") if p]
    return min(max(len(parts), 1), 7)




def _rollup_total(source_rows: list[MainRow], index: int, row: MainRow) -> float | None:
    """Return budget total for an Activity/WBS/Project Summary row.

    Activity totals come directly from the row Amount. Parent totals are the
    sum of descendant Plan activities until the outline climbs back to the
    parent's level. This keeps Dashboard roll-ups independent of whether a
    parent Amount cell happens to be populated in ``main``.
    """
    kind = row.row_type.strip().lower()
    if kind == "activity":
        return float(row.amount) if row.amount is not None else None

    parent_level = _outline_level(row)
    total = 0.0
    found = False
    for candidate in source_rows[index + 1:]:
        if candidate.pa.strip().upper() != "P":
            continue
        candidate_kind = candidate.row_type.strip().lower()
        if candidate_kind not in {"project summary", "wbs", "activity"}:
            continue
        candidate_level = _outline_level(candidate)
        if candidate_level <= parent_level:
            break
        if candidate_kind == "activity" and candidate.amount is not None:
            total += float(candidate.amount)
            found = True
    return total if found else (float(row.amount) if row.amount is not None else None)

def _status(plan_progress: float, actual_progress: float) -> str:
    if plan_progress <= 0 and actual_progress <= 0:
        return "Not Due"
    if plan_progress > 0 and actual_progress <= 0:
        return "No Progress"
    if actual_progress >= 1:
        return "Complete"
    if actual_progress < plan_progress:
        return "Behind"
    return "On Track"


class ActivityTableDeriver:
    """Derive the Dashboard Activity Table directly from MainDataset.

    LW-3 deliberately contains no workbook-library dependency.  It preserves the
    existing two-row Plan/Actual presentation contract while removing the Live
    path's dependency on the generated `progress_table` worksheet.
    """

    def derive(
        self,
        dataset: MainDataset,
        *,
        cutoff: date | datetime | None = None,
    ) -> ActivityTableModel:
        if isinstance(cutoff, datetime):
            cutoff = cutoff.date()

        source_rows = list(dataset.rows)
        rows: list[ActivityTableRow] = []

        index = 0
        while index < len(source_rows):
            plan = source_rows[index]
            if plan.pa.strip().upper() != "P" or plan.row_type.strip().lower() not in {
                "project summary", "wbs", "activity"
            }:
                index += 1
                continue

            actual = None
            if index + 1 < len(source_rows):
                candidate = source_rows[index + 1]
                if candidate.pa.strip().upper() == "A":
                    # Main grammar is Plan/Actual adjacency. Actual rows in legacy
                    # workbooks may intentionally have blank Row Type/Description.
                    # For activities, Activity ID is the identity guard.
                    if plan.row_type.strip().lower() != "activity" or (
                        candidate.activity_id.strip() == plan.activity_id.strip()
                    ):
                        actual = candidate

            plan_progress = _progress(dataset, plan, cutoff)
            actual_progress = _progress(dataset, actual, cutoff)
            total = _rollup_total(source_rows, index, plan)
            plan_amount = total * plan_progress if total is not None else None
            actual_amount = total * actual_progress if total is not None else None
            variance = actual_progress - plan_progress
            status = _status(plan_progress, actual_progress)
            level = _outline_level(plan)
            activity_name = plan.description.strip()
            activity_id = plan.activity_id.strip()
            wbs = plan.wbs.strip()
            kind = plan.row_type.strip().lower()

            rows.append(
                ActivityTableRow(
                    row_type=kind,
                    wbs=wbs,
                    activity=activity_name,
                    activity_id=activity_id,
                    type_label="Plan",
                    total=total,
                    amount=plan_amount,
                    progress=plan_progress,
                    variance=None,
                    status="",
                    outline_level=level,
                    source_plan_row=plan.row_number,
                    source_actual_row=actual.row_number if actual else None,
                )
            )
            rows.append(
                ActivityTableRow(
                    row_type=kind,
                    wbs=wbs,
                    activity="",
                    activity_id=activity_id,
                    type_label="Actual",
                    total=total,
                    amount=actual_amount,
                    progress=actual_progress,
                    variance=variance,
                    status=status,
                    outline_level=level,
                    source_plan_row=plan.row_number,
                    source_actual_row=actual.row_number if actual else None,
                )
            )

            index += 2 if actual is not None else 1

        return ActivityTableModel(cutoff=cutoff, rows=tuple(rows))
