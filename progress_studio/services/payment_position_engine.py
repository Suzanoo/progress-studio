from __future__ import annotations

from bisect import bisect_left

from progress_studio.domain.payment_models import (
    ActivityProgressIndex,
    PaymentInputData,
    PaymentPositionIssue,
    PaymentPositionResult,
    PaymentResolvedPeriod,
    PaymentResolvedPoint,
)


class PaymentPositionEngine:
    """Resolve sparse payment requirements onto weekly cell boundaries."""

    EPSILON = 1e-12

    @staticmethod
    def _boundary(point: PaymentResolvedPoint) -> int:
        return point.timescale_column if point.boundary_edge == "left" else point.timescale_column + 1

    def resolve(self, payment: PaymentInputData, progress: ActivityProgressIndex) -> PaymentPositionResult:
        periods: list[PaymentResolvedPeriod] = []
        issues: list[PaymentPositionIssue] = []

        for period in payment.periods:
            points: list[PaymentResolvedPoint] = []
            for requirement in period.requirements:
                activity = progress.activities.get(requirement.activity_id)
                if activity is None:
                    issues.append(
                        PaymentPositionIssue(
                            period_id=period.period_id,
                            activity_id=requirement.activity_id,
                            code="activity_not_found",
                            message="Activity ID was not found in main Plan rows.",
                        )
                    )
                    continue
                if not activity.buckets:
                    issues.append(
                        PaymentPositionIssue(
                            period_id=period.period_id,
                            activity_id=requirement.activity_id,
                            code="no_plan_distribution",
                            message="Activity has no weekly Plan distribution to locate the requirement.",
                        )
                    )
                    continue

                required = requirement.required_fraction
                if required <= self.EPSILON:
                    bucket = activity.buckets[0]
                    edge = "left"
                elif required >= 1.0 - self.EPSILON:
                    bucket = activity.buckets[-1]
                    edge = "right"
                else:
                    cumulative = [bucket.cumulative_fraction for bucket in activity.buckets]
                    idx = min(bisect_left(cumulative, required), len(activity.buckets) - 1)
                    bucket = activity.buckets[idx]
                    edge = "right"

                points.append(
                    PaymentResolvedPoint(
                        period_id=period.period_id,
                        activity_id=requirement.activity_id,
                        required_fraction=required,
                        activity_row=activity.row_number,
                        timescale_column=bucket.column_index,
                        timescale_column_letter=bucket.column_letter,
                        boundary_edge=edge,
                        week_start=bucket.week_start,
                        reached_cumulative=bucket.cumulative_fraction,
                    )
                )

            points.sort(key=lambda point: point.activity_row)

            planned_eligible_date = None
            controlling_activity_ids: tuple[str, ...] = ()
            if points:
                latest_boundary = max(self._boundary(point) for point in points)
                controlling = tuple(
                    point for point in points
                    if self._boundary(point) == latest_boundary
                )
                planned_eligible_date = max(point.week_start for point in controlling)
                controlling_activity_ids = tuple(point.activity_id for point in controlling)

            periods.append(
                PaymentResolvedPeriod(
                    period_id=period.period_id,
                    payment_date=period.payment_date,
                    points=tuple(points),
                    planned_eligible_date=planned_eligible_date,
                    controlling_activity_ids=controlling_activity_ids,
                )
            )

        return PaymentPositionResult(
            periods=tuple(periods),
            issues=tuple(issues),
            requirement_count=payment.populated_requirements,
            resolved_count=sum(len(period.points) for period in periods),
        )
