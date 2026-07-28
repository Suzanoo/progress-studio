from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .curves import get_distribution


@dataclass(frozen=True)
class AutoDecision:
    distribution: str
    source: str
    matched_rule: str
    reason: str
    confidence: int


def _clean(value: object) -> str:
    return " ".join(str(value or "").strip().lower().split())


def load_rules(path: Path | None = None) -> dict[str, Any]:
    rules_path = path or Path(__file__).with_name("distribution_rules.json")
    with rules_path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)

    default_method = str(data.get("default_distribution", "flat")).lower()
    get_distribution(default_method)  # Validate method now.

    return data


def _match_keyword(text: str, keyword: str, match_type: str) -> bool:
    keyword = _clean(keyword)
    if not keyword:
        return False

    if match_type == "exact":
        return text == keyword
    if match_type == "starts_with":
        return text.startswith(keyword)
    return keyword in text


def decide_distribution(
    *,
    activity_code: object,
    wbs: object,
    activity_name: object,
    rules: dict[str, Any],
) -> AutoDecision:
    values = {
        "activity_code": _clean(activity_code),
        "wbs": _clean(wbs),
        "activity_name": _clean(activity_name),
    }

    priority = rules.get(
        "priority",
        ["activity_code", "wbs", "activity_name"],
    )
    rule_groups = rules.get("rules", {})

    confidence_by_source = {
        "activity_code": 100,
        "wbs": 90,
        "activity_name": 80,
    }

    for source in priority:
        text = values.get(source, "")
        if not text:
            continue

        for rule in rule_groups.get(source, []):
            method = str(rule.get("distribution", "")).lower()
            get_distribution(method)

            match_type = str(rule.get("match", "contains")).lower()
            for keyword in rule.get("keywords", []):
                if _match_keyword(text, str(keyword), match_type):
                    return AutoDecision(
                        distribution=method,
                        source=source,
                        matched_rule=str(keyword),
                        reason=str(rule.get("reason", "Matched configured rule")),
                        confidence=confidence_by_source.get(source, 75),
                    )

    default_method = str(
        rules.get("default_distribution", "flat")
    ).lower()
    get_distribution(default_method)

    return AutoDecision(
        distribution=default_method,
        source="default",
        matched_rule="",
        reason="No matching rule; default distribution applied",
        confidence=0,
    )
