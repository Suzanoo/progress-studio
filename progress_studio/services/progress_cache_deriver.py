
from __future__ import annotations

from progress_studio.domain.main_dataset import MainDataset, MainRow
from progress_studio.domain.progress_cache import ProgressCache, ProgressCachePoint


def _identity(row: MainRow) -> tuple[str, str]:
    return row.activity_id.strip(), row.description.strip()


class ProgressCacheDeriver:
    """Build the LW tiny S-curve cache directly from MainDataset.

    Project progress is amount-weighted from Activity Plan/Actual rows.
    The denominator is the full Plan Amount of all activities, matching the
    existing main-sheet roll-up contract.  No workbook library is used here.
    """

    def derive(self, dataset: MainDataset) -> ProgressCache:
        plan_rows = tuple(dataset.activities)
        total_amount = sum(float(row.amount or 0.0) for row in plan_rows)

        actual_by_identity: dict[tuple[str, str], MainRow] = {}
        for row in dataset.rows:
            if (
                row.row_type.strip().lower() == "activity"
                and row.pa.strip().upper() == "A"
                and row.activity_id.strip()
            ):
                actual_by_identity[_identity(row)] = row

        period_columns = [period.column for period in dataset.periods]
        plan_weekly: list[float | None] = []
        actual_weekly: list[float | None] = []

        for column in period_columns:
            plan_has_value = False
            actual_has_value = False
            plan_weighted = 0.0
            actual_weighted = 0.0

            for plan in plan_rows:
                amount = float(plan.amount or 0.0)
                plan_value = plan.period_value(column)
                if plan_value is not None:
                    plan_has_value = True
                    plan_weighted += amount * float(plan_value)

                actual = actual_by_identity.get(_identity(plan))
                if actual is not None:
                    actual_value = actual.period_value(column)
                    if actual_value is not None:
                        actual_has_value = True
                        actual_weighted += amount * float(actual_value)

            if total_amount <= 0:
                plan_weekly.append(None if not plan_has_value else 0.0)
                actual_weekly.append(None if not actual_has_value else 0.0)
            else:
                plan_weekly.append(plan_weighted / total_amount if plan_has_value else None)
                actual_weekly.append(actual_weighted / total_amount if actual_has_value else None)

        actual_indices = [i for i, value in enumerate(actual_weekly) if value is not None]
        first_actual = min(actual_indices) if actual_indices else None
        last_actual = max(actual_indices) if actual_indices else None

        points: list[ProgressCachePoint] = []
        plan_acc = 0.0
        actual_acc = 0.0
        plan_started = False

        for index, period in enumerate(dataset.periods):
            p_week = plan_weekly[index]
            a_week = actual_weekly[index]

            if p_week is not None:
                plan_acc += p_week
                plan_started = True
            p_acc = plan_acc if plan_started else None

            if (
                first_actual is None
                or last_actual is None
                or index < first_actual
                or index > last_actual
            ):
                a_acc = None
            else:
                if a_week is not None:
                    actual_acc += a_week
                a_acc = actual_acc

            points.append(
                ProgressCachePoint(
                    period_key=period.key,
                    reporting_date=period.reporting_date,
                    plan_weekly=p_week,
                    plan_cumulative=p_acc,
                    actual_weekly=a_week,
                    actual_cumulative=a_acc,
                )
            )

        return ProgressCache(total_amount=total_amount, points=tuple(points))
