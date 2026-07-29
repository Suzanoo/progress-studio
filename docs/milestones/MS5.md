# Progress Studio V3 — MS-5 Persistent Mapping Session

## Goal

Allow users to save incomplete mapping work, close the application, and safely continue later.

## Completed features

- `Save Session...` writes a JSON sidecar.
- `Load Session...` restores the Progress workbook, BOQ workbook, selected worksheet, and all percentage allocations.
- `Recent Sessions...` lists up to ten existing session files, newest first.
- After the first save, Map, Unmap, Undo, and Clear All automatically save the session.
- Session status is visible in the toolbar: not saved, unsaved changes, saved, loaded, auto-saved, or failed.
- Session files are written atomically using a temporary file followed by `os.replace`.
- Progress and BOQ workbooks are fingerprinted with SHA-256.
- Loading verifies both workbooks with SHA-256.
- If a workbook was moved or renamed, the user can relink an identical copy.
- Changed workbook content is rejected; Progress Studio never merges it automatically.
- Version 1 session files are migrated to the current version when loaded.
- Allocation restoration validates BOQ keys, Activity IDs, duplicate pairs, numeric shares, and the 100% maximum.
- `Clear all` requires confirmation and remains recoverable with one Undo command.

## Session file

Default suggested name:

```text
<progress-workbook-name>.mapping.json
```

The JSON stores references and allocations, not complete Excel rows:

```json
{
  "format": "progress-studio-mapping-session",
  "version": 2,
  "saved_at": "2026-07-29T00:00:00+00:00",
  "progress": {
    "path": "C:/Project/progress.xlsx",
    "filename": "progress.xlsx",
    "size": 843577,
    "modified_ns": 0,
    "sha256": "..."
  },
  "boq": {
    "path": "C:/Project/BOQ.xlsx",
    "filename": "BOQ.xlsx",
    "size": 418938,
    "modified_ns": 0,
    "sha256": "..."
  },
  "boq_sheet": "NKC2",
  "allocations": [
    {
      "boq_key": "Project|25|25",
      "activity_id": "A1020",
      "share_percent": 40.0
    }
  ]
}
```

## User workflow

```text
Load Progress workbook
→ Load BOQ workbook and worksheet
→ Map some items
→ Save Session...
→ continue mapping (auto-save is active)
→ close application
→ Load Session... or Recent Sessions...
→ continue from the restored state
```

## Deliberately deferred

- Final workbook export and reconciliation — MS-6.
- On-demand S-Curve generation — MS-7.
- Final packaging and production polish — MS-8.


## Moving or renaming workbooks

The session stores the saved absolute path, filename, and SHA-256 fingerprint. If the original path no longer exists, Progress Studio offers a **Relink workbook** flow. The selected workbook is accepted only when its SHA-256 matches the saved fingerprint. Renaming or moving an unchanged file is safe; editing the workbook is not.

```text
Workbook missing at saved path
→ Browse for moved/renamed workbook
→ Verify SHA-256
→ Continue only when identical
```

## Session schema migration

The current schema is version 2. `MappingSessionRepository.load()` runs ordered migration functions for older supported versions. Version 1 is migrated by adding explicit workbook filenames. Future versions must add a migration function and tests rather than silently changing the JSON contract.

## Undo after loading

Loading a session establishes a new working baseline, so the Undo stack is intentionally empty. Undo applies only to mapping actions performed after the session was loaded.
