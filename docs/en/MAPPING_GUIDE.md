# BOQ Mapping Guide

## Understand the workspace

### Left: Progress Activities

- Tick an Activity row to choose the mapping target.
- WBS rows are shown for context and grouping.
- Use Search to find an Activity ID, WBS, or description.
- Use Previous and Next for large schedules.

### Right: BOQ Items

- Tick one or more BOQ items.
- Filter by WBS-2 and WBS-3.
- Use Search to find descriptions or codes.
- Review Amount, Allocated, Remaining %, Status, and Mapped To.

The divider between the two tables can be dragged. Progress Studio remembers the position.

## Full allocation

Example: a BOQ item worth 100,000 is entirely assigned to Activity A1000.

1. Tick Activity A1000.
2. Tick the BOQ item.
3. Enter `100` in Share.
4. Click **Map**.

Result:

```text
Allocated = 100,000
Remaining = 0%
Status = Full
```

## Partial allocation

Example: split one BOQ item between two activities.

For Activity A1000:

1. Select A1000.
2. Select the BOQ item.
3. Enter `60`.
4. Click **Map**.

For Activity A1010:

1. Select A1010.
2. Select the same BOQ item.
3. Enter `40`.
4. Click **Map**.

The combined share must not exceed 100%.

## Mapping several BOQ items at once

When several BOQ items are selected, the Share value applies to every selected item.

Example:

```text
Selected BOQ items: 5
Share: 25%
```

Each selected item receives a 25% allocation to the selected Activity.

## Correcting mistakes

- **Undo** reverses the latest mapping action.
- **Unmap** removes the selected mapping allocation.
- **Clear all** removes every allocation after confirmation. Use this only when restarting the mapping.

## Heavy-workload tips

- Use WBS filters before text search.
- Map one logical work package at a time.
- Save the session regularly.
- Collapse **Workbook Inputs** after files are loaded.
- Click **Focus Mapping** to hide the generator and maximize the mapping area.
- Full items remain visible so previous decisions can be reviewed or changed.
