# Progress Studio V3 — MS-3 Mapping Engine

## Completed

- One selected Activity can receive one or more selected BOQ items.
- Each BOQ item is allocated at 100% in MS-3.
- Remapping a BOQ item moves its amount from the old Activity to the new Activity.
- Activity mapped amount updates immediately.
- BOQ table shows Amount, Allocated, Remaining, Status, and Mapped To.
- Status model supports Unmapped, Partial, and Full; MS-3 produces Unmapped or Full.
- Undo restores the exact previous assignments for the last command.
- Unmap removes selected BOQ assignments.
- Summary updates mapped amount, remaining amount, and mapped item count.
- GUI refreshes only visible BOQ and Activity rows affected by a command.
- `Clear all` is intentionally excluded until persistent sessions and recovery exist.

## Out of scope

- Percentage share allocation (MS-4).
- Save/load mapping sessions (MS-5).
- Final export redesign (MS-6).
- S-Curve redesign (MS-7).

## Run

```powershell
python desktop.py
```
