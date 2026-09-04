from __future__ import annotations

from dataclasses import dataclass
from math import isclose
from typing import Iterable, Sequence


@dataclass(frozen=True)
class PaymentBreakdownItem:
    code: str
    amount: float
    period_progress: tuple[float, ...]

    @property
    def cumulative_progress(self) -> tuple[float, ...]:
        running = 0.0
        values: list[float] = []
        for value in self.period_progress:
            running += value
            values.append(running)
        return tuple(values)


@dataclass(frozen=True)
class PaymentBreakdownSourceActivity:
    """One source Activity row used by the generic PB-2 engine."""

    activity_id: str
    activity_name: str
    amount: float
    period_progress: tuple[float, ...]
    wbs: str | None = None

    @property
    def cumulative_progress(self) -> tuple[float, ...]:
        running = 0.0
        values: list[float] = []
        for value in self.period_progress:
            running += value
            values.append(running)
        return tuple(values)


@dataclass(frozen=True)
class PaymentBreakdownActivity:
    activity_name: str
    main_amount: float
    items: tuple[PaymentBreakdownItem, ...]
    period_progress: tuple[float, ...]
    cumulative_progress: tuple[float, ...]

    @property
    def breakdown_amount(self) -> float:
        return sum(item.amount for item in self.items)


@dataclass(frozen=True)
class DerivedPaymentActivity:
    """Exact-name aggregate of one or more source Activities.

    Source rows keep their own progress profiles. The combined progress is an
    amount-weighted aggregate. No higher-level keyword/group aggregation is
    performed here.
    """

    activity_name: str
    source_activities: tuple[PaymentBreakdownSourceActivity, ...]
    total_amount: float
    period_progress: tuple[float, ...]
    cumulative_progress: tuple[float, ...]

    @property
    def activity_ids(self) -> tuple[str, ...]:
        return tuple(activity.activity_id for activity in self.source_activities)


class PaymentBreakdownError(ValueError):
    pass


class PaymentBreakdownService:
    """Payment Breakdown calculation contracts.

    Fractions use 0.0..1.0 internally. The service intentionally knows nothing
    about worksheet rows/cells; Excel rendering belongs to infrastructure.

    PB-2 exact-name contract:
    - trim leading/trailing whitespace only;
    - no fuzzy/contains/case normalization;
    - source Activities with the same exact name remain individually visible;
    - their combined progress is amount-weighted;
    - different Activity Names are never re-grouped under a keyword heading.
    """

    EPSILON = 1e-9
    PROGRESS_TOLERANCE = 1e-6

    @staticmethod
    def _name(value: str) -> str:
        return value.strip()

    def exact_name_match(self, main_activity_name: str, breakdown_activity_name: str) -> bool:
        return self._name(main_activity_name) == self._name(breakdown_activity_name)

    def build_activity(
        self,
        *,
        main_activity_name: str,
        main_amount: float,
        breakdown_activity_name: str,
        items: Iterable[PaymentBreakdownItem],
    ) -> PaymentBreakdownActivity | None:
        """PB-1 compatibility contract retained unchanged."""
        if not self.exact_name_match(main_activity_name, breakdown_activity_name):
            return None

        item_list = tuple(items)
        if not item_list:
            raise PaymentBreakdownError("Matched activity has no Payment Breakdown items.")
        if main_amount <= self.EPSILON:
            raise PaymentBreakdownError("Matched activity amount must be greater than zero.")

        period_count = max((len(item.period_progress) for item in item_list), default=0)
        if period_count == 0:
            raise PaymentBreakdownError("Matched activity has no Payment Breakdown periods.")

        for item in item_list:
            if item.amount < -self.EPSILON:
                raise PaymentBreakdownError(f"Negative amount for Payment Breakdown item {item.code!r}.")
            if any(value < -self.EPSILON for value in item.period_progress):
                raise PaymentBreakdownError(f"Negative progress for Payment Breakdown item {item.code!r}.")
            total_progress = sum(item.period_progress)
            if not isclose(total_progress, 1.0, abs_tol=self.PROGRESS_TOLERANCE):
                raise PaymentBreakdownError(
                    f"Payment Breakdown item {item.code!r} must total 100%; got {total_progress:.6%}."
                )

        breakdown_amount = sum(item.amount for item in item_list)
        if not isclose(breakdown_amount, main_amount, abs_tol=0.01):
            raise PaymentBreakdownError(
                "Payment Breakdown amount does not balance to main Activity Amount: "
                f"main={main_amount:.2f}, breakdown={breakdown_amount:.2f}."
            )

        weighted = self._weighted_progress(
            amounts=(item.amount for item in item_list),
            profiles=(item.period_progress for item in item_list),
        )
        cumulative = self._cumulative(weighted)

        return PaymentBreakdownActivity(
            activity_name=self._name(main_activity_name),
            main_amount=main_amount,
            items=item_list,
            period_progress=weighted,
            cumulative_progress=cumulative,
        )

    def derive_activity(
        self,
        source_activities: Iterable[PaymentBreakdownSourceActivity],
    ) -> DerivedPaymentActivity:
        """Aggregate source rows only when their Activity Names match exactly."""
        activities = tuple(source_activities)
        if not activities:
            raise PaymentBreakdownError("Cannot derive Payment Breakdown from zero source activities.")

        canonical_name = self._name(activities[0].activity_name)
        if not canonical_name:
            raise PaymentBreakdownError("Activity Name must not be blank.")

        seen_ids: set[str] = set()
        for activity in activities:
            if self._name(activity.activity_name) != canonical_name:
                raise PaymentBreakdownError(
                    "Cannot combine different Activity Names in one Payment Breakdown activity."
                )
            if not activity.activity_id:
                raise PaymentBreakdownError("Source Activity ID must not be blank.")
            if activity.activity_id in seen_ids:
                raise PaymentBreakdownError(
                    f"Duplicate source Activity ID {activity.activity_id!r}."
                )
            seen_ids.add(activity.activity_id)
            self._validate_source_activity(activity)

        total_amount = sum(activity.amount for activity in activities)
        if total_amount <= self.EPSILON:
            raise PaymentBreakdownError("Derived Activity amount must be greater than zero.")

        weighted = self._weighted_progress(
            amounts=(activity.amount for activity in activities),
            profiles=(activity.period_progress for activity in activities),
        )
        cumulative = self._cumulative(weighted)

        return DerivedPaymentActivity(
            activity_name=canonical_name,
            source_activities=activities,
            total_amount=total_amount,
            period_progress=weighted,
            cumulative_progress=cumulative,
        )

    def derive_exact_name_groups(
        self,
        source_activities: Iterable[PaymentBreakdownSourceActivity],
        *,
        min_occurrences: int = 2,
    ) -> tuple[DerivedPaymentActivity, ...]:
        """Discover generic exact-name candidates from source Activities.

        ``min_occurrences=2`` models the current Payment-Breakdown derived
        Activity contract: only repeated exact Activity Names are candidates.
        A caller may use ``1`` when it explicitly wants singleton exact-name
        activities too (for diagnostics/prototyping such as the U-Glass review).
        """
        if min_occurrences < 1:
            raise ValueError("min_occurrences must be at least 1.")

        grouped: dict[str, list[PaymentBreakdownSourceActivity]] = {}
        order: list[str] = []
        for activity in source_activities:
            name = self._name(activity.activity_name)
            if not name:
                continue
            if name not in grouped:
                grouped[name] = []
                order.append(name)
            grouped[name].append(activity)

        results: list[DerivedPaymentActivity] = []
        for name in order:
            group = grouped[name]
            if len(group) < min_occurrences:
                continue
            results.append(self.derive_activity(group))
        return tuple(results)

    def _validate_source_activity(self, activity: PaymentBreakdownSourceActivity) -> None:
        if activity.amount <= self.EPSILON:
            raise PaymentBreakdownError(
                f"Source Activity {activity.activity_id!r} amount must be greater than zero."
            )
        if not activity.period_progress:
            raise PaymentBreakdownError(
                f"Source Activity {activity.activity_id!r} has no progress periods."
            )
        if any(value < -self.EPSILON for value in activity.period_progress):
            raise PaymentBreakdownError(
                f"Negative progress for source Activity {activity.activity_id!r}."
            )
        total_progress = sum(activity.period_progress)
        if not isclose(total_progress, 1.0, abs_tol=self.PROGRESS_TOLERANCE):
            raise PaymentBreakdownError(
                f"Source Activity {activity.activity_id!r} must total 100%; "
                f"got {total_progress:.6%}."
            )

    def _weighted_progress(
        self,
        *,
        amounts: Iterable[float],
        profiles: Iterable[Sequence[float]],
    ) -> tuple[float, ...]:
        amount_values = tuple(amounts)
        profile_values = tuple(tuple(profile) for profile in profiles)
        if len(amount_values) != len(profile_values):
            raise PaymentBreakdownError("Amounts/progress profile counts do not match.")

        total_amount = sum(amount_values)
        if total_amount <= self.EPSILON:
            raise PaymentBreakdownError("Weighted progress amount must be greater than zero.")

        period_count = max((len(profile) for profile in profile_values), default=0)
        if period_count == 0:
            raise PaymentBreakdownError("Weighted progress has no periods.")

        weighted: list[float] = []
        for period_index in range(period_count):
            amount_progress = 0.0
            for amount, profile in zip(amount_values, profile_values):
                fraction = profile[period_index] if period_index < len(profile) else 0.0
                amount_progress += amount * fraction
            weighted.append(amount_progress / total_amount)
        return tuple(weighted)

    @staticmethod
    def _cumulative(period_progress: Sequence[float]) -> tuple[float, ...]:
        running = 0.0
        cumulative: list[float] = []
        for value in period_progress:
            running += value
            cumulative.append(running)
        if cumulative:
            cumulative[-1] = 1.0
        return tuple(cumulative)
