from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from math import isclose

from progress_studio.domain.earned_value import (
    ActivityEarnedValue,
    BOQEarnedValue,
    EarnedValuePoint,
    EarnedValueResult,
)
from progress_studio.domain.main_dataset import MainDataset, MainRow
from progress_studio.domain.mapping_models import AllocationRecord, BOQRow


class EarnedValueInputError(ValueError):
    """Raised when EV input violates a frozen EV-0 contract."""


def _identity(row: MainRow) -> str:
    return row.activity_id.strip()


def _ratio(numerator: float | None, denominator: float | None) -> float | None:
    if numerator is None or denominator is None or denominator <= 0.0:
        return None
    return numerator / denominator


def _point(*, period_key: str, reporting_date: datetime | None, pv: float | None, ev: float | None) -> EarnedValuePoint:
    sv = None if pv is None or ev is None else ev - pv
    return EarnedValuePoint(
        period_key=period_key,
        reporting_date=reporting_date,
        planned_value=pv,
        earned_value=ev,
        schedule_variance=sv,
        schedule_performance_index=_ratio(ev, pv),
    )


class EarnedValueDeriver:
    def __init__(self, *, allocation_tolerance: float = 0.01) -> None:
        self._allocation_tolerance = float(allocation_tolerance)

    def derive(self, dataset: MainDataset, boq_rows, allocations, *, cutoff_date: datetime | None) -> EarnedValueResult:
        plan_rows = tuple(dataset.activities)
        plan_by_id = {_identity(row): row for row in plan_rows}
        actual_by_id: dict[str, MainRow] = {}
        for row in dataset.rows:
            if row.pa.strip().upper() == "A" and row.activity_id.strip():
                actual_by_id[_identity(row)] = row
        boq_by_key = {row.key: row for row in boq_rows}
        allocation_by_boq: dict[str, list[AllocationRecord]] = defaultdict(list)
        for record in allocations:
            if record.boq_key not in boq_by_key:
                raise EarnedValueInputError(f"Allocation references unknown BOQ item: {record.boq_key}")
            if record.activity_id.strip() not in plan_by_id:
                raise EarnedValueInputError(f"Allocation references unknown Activity ID: {record.activity_id} (BOQ {record.boq_key})")
            allocation_by_boq[record.boq_key].append(record)
        self._validate_full_allocation(tuple(boq_rows), allocation_by_boq)

        reporting_periods = self._reporting_periods(dataset, plan_rows)
        activity_progress = self._derive_activity_progress(
            periods=reporting_periods,
            plan_rows=plan_rows,
            actual_by_id=actual_by_id,
            cutoff_date=cutoff_date,
        )
        activities: list[ActivityEarnedValue] = []
        for plan in plan_rows:
            activity_id = _identity(plan)
            bac = float(plan.amount or 0.0)
            plan_acc, actual_acc = activity_progress[activity_id]
            points = tuple(
                _point(
                    period_key=period.key,
                    reporting_date=period.reporting_date,
                    pv=None if plan_acc[index] is None else bac * plan_acc[index],
                    ev=None if actual_acc[index] is None else bac * actual_acc[index],
                )
                for index, period in enumerate(reporting_periods)
            )
            activities.append(ActivityEarnedValue(activity_id=activity_id, description=plan.description, wbs=plan.wbs, bac=bac, points=points))

        project_bac = sum(item.bac for item in activities)
        project_points = self._aggregate_activity_points(reporting_periods, tuple(activities))
        boq_items: list[BOQEarnedValue] = []
        for boq in boq_rows:
            bac = float(boq.amount or 0.0)
            records = allocation_by_boq.get(boq.key, [])
            points: list[EarnedValuePoint] = []
            for index, period in enumerate(reporting_periods):
                pv_total = 0.0
                ev_total = 0.0
                pv_has_value = False
                ev_has_value = False
                for record in records:
                    share_amount = bac * float(record.share_percent) / 100.0
                    plan_acc, actual_acc = activity_progress[record.activity_id.strip()]
                    if plan_acc[index] is not None:
                        pv_has_value = True
                        pv_total += share_amount * plan_acc[index]
                    if actual_acc[index] is not None:
                        ev_has_value = True
                        ev_total += share_amount * actual_acc[index]
                points.append(_point(period_key=period.key, reporting_date=period.reporting_date, pv=pv_total if pv_has_value else None, ev=ev_total if ev_has_value else None))
            boq_items.append(BOQEarnedValue(boq_key=boq.key, stable_id=boq.stable_id, description=boq.description, bac=bac, points=tuple(points)))

        return EarnedValueResult(
            cutoff_date=cutoff_date,
            project_bac=project_bac,
            project_points=project_points,
            activities=tuple(activities),
            boq_items=tuple(boq_items),
        )

    def _validate_full_allocation(self, boq_rows: tuple[BOQRow, ...], allocation_by_boq: dict[str, list[AllocationRecord]]) -> None:
        incomplete: list[str] = []
        for boq in boq_rows:
            if float(boq.amount or 0.0) <= 0.0:
                continue
            total_share = sum(float(record.share_percent) for record in allocation_by_boq.get(boq.key, []))
            if not isclose(total_share, 100.0, rel_tol=0.0, abs_tol=self._allocation_tolerance):
                label = boq.stable_id or boq.key
                incomplete.append(f"{label}: {total_share:.2f}%")
        if incomplete:
            detail = "\n".join(incomplete)
            raise EarnedValueInputError(
                "Earned Value requires every positive-amount BOQ item to be allocated to 100%.\n" + detail
            )

    @staticmethod
    def _reporting_periods(dataset: MainDataset, plan_rows: tuple[MainRow, ...]) -> tuple:
        active_indices = [
            index
            for index, period in enumerate(dataset.periods)
            if any(plan.period_value(period.column) is not None for plan in plan_rows)
        ]
        if not active_indices:
            return ()
        first = min(active_indices)
        last = max(active_indices)
        return tuple(dataset.periods[first:last + 1])

    @staticmethod
    def _derive_activity_progress(*, periods: tuple, plan_rows: tuple[MainRow, ...], actual_by_id: dict[str, MainRow], cutoff_date: datetime | None):
        result: dict[str, tuple[list[float | None], list[float | None]]] = {}
        for plan in plan_rows:
            actual = actual_by_id.get(_identity(plan))
            plan_acc = 0.0
            actual_acc = 0.0
            plan_started = False
            actual_started = False
            plan_values: list[float | None] = []
            actual_values: list[float | None] = []
            for period in periods:
                p_week = plan.period_value(period.column)
                if p_week is not None:
                    plan_acc += float(p_week)
                    plan_started = True
                plan_values.append(plan_acc if plan_started else None)
                after_cutoff = (
                    cutoff_date is not None
                    and period.reporting_date is not None
                    and period.reporting_date > cutoff_date
                )
                if after_cutoff:
                    actual_values.append(None)
                    continue
                a_week = actual.period_value(period.column) if actual is not None else None
                if a_week is not None:
                    actual_acc += float(a_week)
                    actual_started = True
                # Reuse the existing Progress/overlay contract: before the first
                # reported Actual, cumulative Actual is a visible 0 baseline.
                actual_values.append(actual_acc if actual_started else 0.0)
            result[_identity(plan)] = (plan_values, actual_values)
        return result

    @staticmethod
    def _aggregate_activity_points(periods: tuple, activities: tuple[ActivityEarnedValue, ...]) -> tuple[EarnedValuePoint, ...]:
        points: list[EarnedValuePoint] = []
        for index, period in enumerate(periods):
            pv_values = [a.points[index].planned_value for a in activities]
            ev_values = [a.points[index].earned_value for a in activities]
            pv_has_value = any(value is not None for value in pv_values)
            ev_has_value = any(value is not None for value in ev_values)
            pv = sum(value or 0.0 for value in pv_values) if pv_has_value else None
            ev = sum(value or 0.0 for value in ev_values) if ev_has_value else None
            points.append(_point(period_key=period.key, reporting_date=period.reporting_date, pv=pv, ev=ev))
        return tuple(points)
