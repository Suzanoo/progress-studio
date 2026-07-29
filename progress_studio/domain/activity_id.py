from __future__ import annotations

import re
from collections.abc import Collection

_SANITIZE = re.compile(r"[^A-Za-z0-9]+")


def ensure_activity_id(
    source_id: str | None,
    *,
    task_id: int | None,
    uid: int | None,
    source_order: int,
    existing_ids: Collection[str] = (),
) -> str:
    """Return a unique source Activity ID or a deterministic generated fallback."""
    value = str(source_id or "").strip()
    if value and value not in existing_ids:
        return value

    if not value:
        identity = task_id if task_id is not None else uid
        if identity is None:
            identity = source_order + 1
        candidate = f"ACT-{int(identity):06d}"
    else:
        candidate = value

    if candidate not in existing_ids:
        return candidate

    suffix = 2
    while f"{candidate}__{suffix}" in existing_ids:
        suffix += 1
    return f"{candidate}__{suffix}"


def normalize_activity_id(value: object) -> str:
    """Normalize an Activity ID for reliable matching."""
    raw = str(value or "").strip()
    return _SANITIZE.sub("-", raw).strip("-")
