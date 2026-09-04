from __future__ import annotations

from dataclasses import dataclass
from math import isclose

from progress_studio.domain.main_dataset import MainDataset, MainPeriod
from progress_studio.services.payment_breakdown_service import (
    DerivedPaymentActivity,
    PaymentBreakdownService,
    PaymentBreakdownSourceActivity,
)


@dataclass(frozen=True, slots=True)
class PaymentBreakdownDatasetSnapshot:
    """Derived Payment-Breakdown data ready for an Excel renderer.

    `main` remains the source of truth.  The adapter only accepts Plan Activity
    rows that can satisfy the frozen Payment-Breakdown contract:
    - positive Amount;
    - non-blank Activity ID / Activity Name;
    - weekly Plan profile totals 100%;
    - exact Activity Name grouping only.
    """

    periods: tuple[MainPeriod, ...]
    activities: tuple[DerivedPaymentActivity, ...]
    eligible_source_count: int
    skipped_activity_ids: tuple[str, ...]


class MainDatasetPaymentBreakdownAdapter:
    """Translate MainDataset into the generic PB calculation contract."""

    def __init__(
        self,
        service: PaymentBreakdownService | None = None,
        *,
        progress_tolerance: float | None = None,
    ) -> None:
        self.service = service or PaymentBreakdownService()
        self.progress_tolerance = (
            float(progress_tolerance)
            if progress_tolerance is not None
            else self.service.PROGRESS_TOLERANCE
        )

    def derive(
        self,
        dataset: MainDataset,
        *,
        min_occurrences: int = 2,
    ) -> PaymentBreakdownDatasetSnapshot:
        sources: list[PaymentBreakdownSourceActivity] = []
        skipped: list[str] = []

        for row in dataset.activities:
            activity_id = row.activity_id.strip()
            activity_name = row.description.strip()
            amount = float(row.amount or 0.0)

            if not activity_id or not activity_name or amount <= self.service.EPSILON:
                if activity_id:
                    skipped.append(activity_id)
                continue

            profile = tuple(
                float(row.period_value(period.column) or 0.0)
                for period in dataset.periods
            )
            if (
                not profile
                or any(value < -self.service.EPSILON for value in profile)
                or not isclose(
                    sum(profile),
                    1.0,
                    rel_tol=0.0,
                    abs_tol=self.progress_tolerance,
                )
            ):
                skipped.append(activity_id)
                continue

            sources.append(
                PaymentBreakdownSourceActivity(
                    activity_id=activity_id,
                    activity_name=activity_name,
                    amount=amount,
                    period_progress=profile,
                    wbs=row.wbs.strip() or None,
                )
            )

        activities = self.service.derive_exact_name_groups(
            sources,
            min_occurrences=min_occurrences,
        )
        return PaymentBreakdownDatasetSnapshot(
            periods=tuple(dataset.periods),
            activities=activities,
            eligible_source_count=len(sources),
            skipped_activity_ids=tuple(skipped),
        )
