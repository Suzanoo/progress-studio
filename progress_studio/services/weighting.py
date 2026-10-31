"""Source-neutral Create weighting; values are non-monetary calculation units."""
from __future__ import annotations

import math
from progress_studio.domain import Activity

XML_FIELD = "xml_field"
MANUAL_PLACEHOLDER = "manual_placeholder"


def validate_amount_initialization(basis: str | None, initialization: str,
                                   amount_field: str | None = None) -> str:
    """Validate explicit creation intent; never fall back from a bad XML field."""
    if initialization not in {XML_FIELD, MANUAL_PLACEHOLDER}:
        raise ValueError("Unknown Amount initialization source.")
    if initialization == MANUAL_PLACEHOLDER:
        if basis != "amount":
            raise ValueError("Enter in Excel later requires Amount weighting.")
        if amount_field:
            raise ValueError("Enter in Excel later cannot also select an XML Amount field.")
    return initialization


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
