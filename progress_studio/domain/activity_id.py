from __future__ import annotations

import re

_SANITIZE = re.compile(r"[^A-Za-z0-9]+")


def ensure_activity_id(
    source_id: str | None,
    *,
    task_id: int | None,
    uid: int | None,
    source_order: int,
) -> str:
    """Return the source Activity ID or a deterministic generated fallback."""
    value = str(source_id or "").strip()
    if value:
        return value

    identity = task_id if task_id is not None else uid
    if identity is None:
        identity = source_order + 1
    return f"GEN-{int(identity):05d}"


def normalize_activity_id(value: object) -> str:
    """Normalize an Activity ID for reliable matching."""
    raw = str(value or "").strip()
    return _SANITIZE.sub("-", raw).strip("-")
