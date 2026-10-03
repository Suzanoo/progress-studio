from __future__ import annotations

# MS9.1 starts with English as the canonical UI language.  All new shell text
# is centralized here so later locales do not require edits to view classes.
TEXT = {
    "app_name": "Progress Studio",
    "project_default": "Project: Local Workspace",
    "mapping_workspace": "Mapping Workspace",
    "overview": "Overview",
    "batch_mapping": "Batch Mapping",
    "mapping_rules": "Mapping Rules",
    "mapping_memory": "Mapping Memory",
    "boq_data": "BOQ Data",
    "progress_activities": "Progress Activities",
    "project_settings": "Project Settings",
    "ai_settings": "Model & AI Settings",
    "preferences": "Preferences",
    "generator": "Workbook Generator",
    "activity_log": "Activity Log",
    "ready": "Ready",
}


from .translations import PAIRS, GUIDE

# Existing English message IDs remain stable during incremental migration.
# New views use namespaced keys. Never apply tr() to user data or service enums.
EN = {**TEXT, **{key: key for key in PAIRS}, **{key: value[0] for key, value in GUIDE.items()}}
TH = {**TEXT, **PAIRS, **{key: value[1] for key, value in GUIDE.items()}}
TH.update({"ready": "พร้อมใช้งาน", "mapping_workspace": "จับคู่ BOQ กับกิจกรรม"})
RESOURCES = {"en": EN, "th": TH}
_language = "en"


def set_language(language: str) -> None:
    """Called once at launch, not when the persisted preference changes."""
    global _language
    _language = language if language in RESOURCES else "en"


def tr(key: str, *, locale: str | None = None, **values) -> str:
    template = RESOURCES.get(locale or _language, EN).get(key, EN.get(key, key))
    return template.format(**values) if values else template
