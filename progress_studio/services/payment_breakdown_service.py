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
class PaymentBreakdownActivity:
    activity_name: str
    main_amount: float
    items: tuple[PaymentBreakdownItem, ...]
    period_progress: tuple[float, ...]
    cumulative_progress: tuple[float, ...]

    @property
    def breakdown_amount(self) -> float:
        return sum(item.amount for item in self.items)


class PaymentBreakdownError(ValueError):
    pass


class PaymentBreakdownService:
    """PB-1 calculation contract for exact-name Payment Breakdown activities.

    Fractions use 0.0..1.0 internally.  The service intentionally knows nothing
    about worksheet rows/cells; Excel rendering belongs to infrastructure.
    """

    EPSILON = 1e-9

    @staticmethod
    def _name(value: str) -> str:
        # PB-1 allows harmless leading/trailing whitespace only.
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
            if not isclose(total_progress, 1.0, abs_tol=1e-6):
                raise PaymentBreakdownError(
                    f"Payment Breakdown item {item.code!r} must total 100%; got {total_progress:.6%}."
                )

        breakdown_amount = sum(item.amount for item in item_list)
        if not isclose(breakdown_amount, main_amount, abs_tol=0.01):
            raise PaymentBreakdownError(
                "Payment Breakdown amount does not balance to main Activity Amount: "
                f"main={main_amount:.2f}, breakdown={breakdown_amount:.2f}."
            )

        weighted: list[float] = []
        for period_index in range(period_count):
            amount_progress = 0.0
            for item in item_list:
                fraction = (
                    item.period_progress[period_index]
                    if period_index < len(item.period_progress)
                    else 0.0
                )
                amount_progress += item.amount * fraction
            weighted.append(amount_progress / breakdown_amount)

        cumulative: list[float] = []
        running = 0.0
        for value in weighted:
            running += value
            cumulative.append(running)
        # Avoid harmless floating-point residue at the contract boundary.
        cumulative[-1] = 1.0

        return PaymentBreakdownActivity(
            activity_name=self._name(main_activity_name),
            main_amount=main_amount,
            items=item_list,
            period_progress=tuple(weighted),
            cumulative_progress=tuple(cumulative),
        )
