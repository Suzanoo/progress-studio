"""Validate a chosen custom field atomically, without Excel or UI dependencies."""
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, localcontext
import math
import re

from progress_studio.domain.amount_field import AmountField


@dataclass(frozen=True)
class AmountValuePreview:
    activity_id: str
    name: str
    raw_value: str | None
    value: Decimal | None
    status: str


@dataclass(frozen=True)
class AmountPreview:
    field: AmountField
    rows: tuple[AmountValuePreview, ...]
    ordinary_count: int
    positive_count: int
    zero_count: int  # ordinary activities only; milestones reported separately
    milestone_count: int
    total: Decimal
    errors: tuple[str, ...]

    @property
    def valid(self):
        return not self.errors


def preview_amounts(rows, fields, selected_field: str | None) -> AmountPreview:
    if not selected_field:
        raise ValueError("Select an XML Amount field before Create.")
    matches = [field for field in fields if field.identity == selected_field]
    if len(matches) != 1:
        raise ValueError(f"Amount field {selected_field!r} is missing, not numeric, or has ambiguous definitions in this XML.")
    field = matches[0]
    previews, errors, ordinary_values = [], [], []
    ordinary = milestones = positive = zero = 0
    for row in rows:
        if row.is_summary:
            continue
        values = [value for key, value in row.amount_field_values if key == field.identity]
        raw = values[0] if len(values) == 1 else None
        if row.is_milestone:
            milestones += 1
            previews.append(AmountValuePreview(row.activity_id, row.name, raw, Decimal(0), "Milestone: zero weight; source value ignored"))
            continue
        ordinary += 1
        value = None
        issue = None
        if len(values) > 1:
            issue = "duplicate field values"
        elif raw is None or not raw.strip():
            issue = "missing Amount"
        else:
            try:
                value = Decimal(raw.strip())
                if not value.is_finite():
                    issue = "non-finite Amount"
                elif not re.fullmatch(r"[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?", raw.strip()):
                    issue = "non-numeric Amount"
                elif value < 0:
                    issue = "negative Amount"
                elif not math.isfinite(float(value)) or (value > 0 and float(value) == 0):
                    issue = "Amount outside supported workbook numeric range"
            except (InvalidOperation, ValueError, OverflowError):
                issue = "non-numeric Amount"
        if issue:
            errors.append(f"{row.activity_id} ({row.name}) — {field.label}: {issue}.")
            previews.append(AmountValuePreview(row.activity_id, row.name, raw, None, issue))
        else:
            ordinary_values.append(value)
            positive += value > 0
            zero += value == 0
            previews.append(AmountValuePreview(row.activity_id, row.name, raw, value, "Zero Amount" if value == 0 else "OK"))
    with localcontext() as ctx:
        # Exact preview sum for all values representable by the existing float engine.
        ctx.prec = max(28, max((v.adjusted() for v in ordinary_values), default=0)
                       - min((v.as_tuple().exponent for v in ordinary_values), default=0)
                       + len(str(len(ordinary_values))) + 4)
        total = sum(ordinary_values, Decimal(0))
    if not errors and (not total.is_finite() or total <= 0 or not math.isfinite(float(total))):
        errors.append(f"{field.label}: ordinary activity Amount total must be positive and within supported workbook numeric range.")
    return AmountPreview(field, tuple(previews), ordinary, positive, zero, milestones, total, tuple(errors))


def assign_xml_amounts(rows, fields, selected_field):
    preview = preview_amounts(rows, fields, selected_field)
    if preview.errors:
        raise ValueError("Invalid XML Amount weighting:\n" + "\n".join(preview.errors))
    weights = {row.activity_id: float(row.value) for row in preview.rows}
    for row in rows:
        if not row.is_summary:
            row.amount = weights[row.activity_id]
    return preview
