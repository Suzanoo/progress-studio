from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Callable


WeightFunction = Callable[[int], list[float]]


@dataclass(frozen=True)
class DistributionSpec:
    key: str
    name: str
    description: str
    preview: str
    suitable_for: tuple[str, ...]
    generator: WeightFunction


def _normalize(weights: list[float]) -> list[float]:
    if not weights:
        return []

    cleaned = [max(0.0, float(value)) for value in weights]
    total = sum(cleaned)

    if total <= 0:
        return [1.0 / len(cleaned)] * len(cleaned)

    normalized = [value / total for value in cleaned]

    # Floating-point correction so the final value closes exactly at 100%.
    if len(normalized) > 1:
        normalized[-1] = max(0.0, 1.0 - sum(normalized[:-1]))
    else:
        normalized[0] = 1.0

    return normalized


def flat_weights(count: int) -> list[float]:
    if count <= 0:
        return []
    return _normalize([1.0] * count)


def front_loaded_weights(count: int) -> list[float]:
    if count <= 0:
        return []

    # Linear decay from high initial effort to lower final effort.
    # Minimum stays above zero so every active period receives progress.
    weights = [count - index for index in range(count)]
    return _normalize(weights)


def back_loaded_weights(count: int) -> list[float]:
    if count <= 0:
        return []

    weights = [index + 1 for index in range(count)]
    return _normalize(weights)


def bell_curve_weights(count: int) -> list[float]:
    if count <= 0:
        return []

    if count == 1:
        return [1.0]

    center = (count - 1) / 2.0
    sigma = max(count / 5.0, 0.75)

    weights = [
        math.exp(-0.5 * ((index - center) / sigma) ** 2)
        for index in range(count)
    ]
    return _normalize(weights)


DISTRIBUTIONS: dict[str, DistributionSpec] = {
    "flat": DistributionSpec(
        key="flat",
        name="Flat Rate",
        description="Equal progress across the activity duration",
        preview="████████████",
        suitable_for=(
            "General activities",
            "Activities with a steady production rate",
        ),
        generator=flat_weights,
    ),
    "front": DistributionSpec(
        key="front",
        name="Front Loaded",
        description="Higher progress at the beginning, then gradually decreases",
        preview="████████▇▆▅▄▃▂",
        suitable_for=(
            "Earthwork",
            "Foundation",
            "Procurement",
            "Activities accelerated at the beginning",
        ),
        generator=front_loaded_weights,
    ),
    "back": DistributionSpec(
        key="back",
        name="Back Loaded",
        description="Lower progress at the beginning, increasing toward completion",
        preview="▂▃▄▅▆▇████████",
        suitable_for=(
            "Testing",
            "Commissioning",
            "Finishing",
            "Activities with most output near completion",
        ),
        generator=back_loaded_weights,
    ),
    "bell": DistributionSpec(
        key="bell",
        name="Bell Curve",
        description="Starts slowly, peaks in the middle, then decreases",
        preview="▁▂▄▆████▆▄▂▁",
        suitable_for=(
            "Concrete works",
            "Superstructure",
            "General building works",
            "Construction activities with a middle-period production peak",
        ),
        generator=bell_curve_weights,
    ),
}


def get_distribution(key: str) -> DistributionSpec:
    normalized = key.strip().lower()
    aliases = {
        "1": "flat",
        "2": "front",
        "3": "back",
        "4": "bell",
        "front_loaded": "front",
        "back_loaded": "back",
        "bell_curve": "bell",
    }
    normalized = aliases.get(normalized, normalized)

    if normalized not in DISTRIBUTIONS:
        available = ", ".join(DISTRIBUTIONS)
        raise KeyError(
            f"Unknown distribution '{key}'. Available: {available}"
        )

    return DISTRIBUTIONS[normalized]


def list_distributions() -> list[DistributionSpec]:
    return [
        DISTRIBUTIONS["flat"],
        DISTRIBUTIONS["front"],
        DISTRIBUTIONS["back"],
        DISTRIBUTIONS["bell"],
    ]
