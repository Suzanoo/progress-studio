
from __future__ import annotations

from pathlib import Path

from openpyxl.utils import get_column_letter

from progress_studio.domain.main_dataset import MainDataset
from progress_studio.domain.payment_models import (
    ActivityProgress,
    ActivityProgressBucket,
    ActivityProgressIndex,
)
from progress_studio.infrastructure.excel.payment_workbook import PaymentWorkbookError


class MainDatasetPaymentProgressAdapter:
    """Create the compact Payment plan index from the already parsed MainDataset."""

    EPSILON = 1e-12

    def build(self, dataset: MainDataset, workbook_path: Path) -> ActivityProgressIndex:
        timescale = tuple(
            (
                period.column,
                get_column_letter(period.column),
                period.reporting_date.date(),
            )
            for period in dataset.periods
            if period.reporting_date is not None
        )
        if not timescale:
            raise PaymentWorkbookError("Weekly timescale columns were not found in 'main'.")

        activities: dict[str, ActivityProgress] = {}
        for row in dataset.activities:
            activity_id = row.activity_id.strip()
            if activity_id in activities:
                raise PaymentWorkbookError(f"Duplicate Plan Activity ID in main: {activity_id}")

            raw = []
            total = 0.0
            for col, letters, week_start in timescale:
                value = row.period_value(col)
                if value is None or abs(float(value)) <= self.EPSILON:
                    continue
                value = float(value)
                if value < -self.EPSILON:
                    raise PaymentWorkbookError(
                        f"Negative plan distribution found for {activity_id} at {letters}{row.row_number}."
                    )
                total += value
                raw.append((col, letters, week_start, value))

            buckets = []
            if total > self.EPSILON:
                cumulative = 0.0
                for col, letters, week_start, value in raw:
                    normalized = value / total
                    cumulative = min(cumulative + normalized, 1.0)
                    buckets.append(
                        ActivityProgressBucket(
                            column_index=col,
                            column_letter=letters,
                            week_start=week_start,
                            incremental_fraction=normalized,
                            cumulative_fraction=cumulative,
                        )
                    )
                if buckets:
                    last = buckets[-1]
                    buckets[-1] = ActivityProgressBucket(
                        column_index=last.column_index,
                        column_letter=last.column_letter,
                        week_start=last.week_start,
                        incremental_fraction=last.incremental_fraction,
                        cumulative_fraction=1.0,
                    )

            activities[activity_id] = ActivityProgress(
                activity_id=activity_id,
                row_number=row.row_number,
                buckets=tuple(buckets),
            )

        if not activities:
            raise PaymentWorkbookError("No Plan activity rows were found in 'main'.")

        return ActivityProgressIndex(
            workbook=Path(workbook_path),
            sheet="main",
            timescale_columns=timescale,
            activities=activities,
        )
