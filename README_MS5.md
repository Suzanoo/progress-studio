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
- Loading is rejected if either workbook is missing or has changed since the session was saved.
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
  "version": 1,
  "saved_at": "2026-07-29T00:00:00+00:00",
  "progress": {
    "path": "C:/Project/progress.xlsx",
    "size": 843577,
    "modified_ns": 0,
    "sha256": "..."
  },
  "boq": {
    "path": "C:/Project/BOQ.xlsx",
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
