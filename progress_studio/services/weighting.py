"""Source-neutral Create weighting; values are non-monetary calculation units."""
from __future__ import annotations

import math
from progress_studio.domain import Activity


def validate_weight_basis(basis: str) -> str:
    basis = basis.strip().lower()
    if basis not in {"equal", "duration", "amount"}:
        raise ValueError("Weight basis must be Equal, Duration or Amount.")
    return basis


def assign_dummy_weights(rows: list[Activity], basis: str) -> None:
    basis = validate_weight_basis(basis)
    if basis == "amount":
        raise ValueError("Amount requires an explicitly selected XML field; dummy weighting cannot be used.")
    assignments = []
    errors = []
    for row in rows:
        if row.is_summary:
            continue
        if row.is_milestone:
            weight = 0.0
        elif basis == "equal":
            weight = 1.0
        else:
            weight = row.duration_hours
            if weight is None or not math.isfinite(weight) or weight <= 0:
                errors.append(f"{row.activity_id} ({row.name}): source duration must be finite and greater than zero for an ordinary activity.")
                continue
        assignments.append((row, weight))
    if errors:
        raise ValueError("Invalid Duration weighting:\n" + "\n".join(errors))
    total = sum(weight for _, weight in assignments)
    if not math.isfinite(total) or total <= 0:
        raise ValueError("Weight total must be finite and greater than zero; no eligible positive-weight activities.")
    for row, weight in assignments:
        row.amount = weight
