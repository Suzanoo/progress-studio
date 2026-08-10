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


def tr(key: str) -> str:
    return TEXT.get(key, key)
