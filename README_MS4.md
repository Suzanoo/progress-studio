# Progress Studio V3 — MS-4 Share Allocation

## Goal

Allow one BOQ item to be distributed across multiple Activities by percentage while preserving fast, memory-backed mapping.

## Completed

- Added a `Share %` input to Amount Mapping; default is 100%.
- One selected Activity may receive the entered share from one or more selected BOQ items.
- Mapping the same BOQ/Activity pair again replaces that pair's share.
- A BOQ item's combined shares cannot exceed 100%.
- BOQ rows show live Allocated, Remaining, Partial/Full status, and every mapped Activity with its share.
- Progress Activity Amount is calculated from allocated amounts, not full BOQ amounts.
- Unmap removes only the selected Activity's share from selected BOQ items.
- Undo restores the exact previous share values.
- Export writes `Share %` and `Allocated Amount` to `BOQ Activity Mapping`.

## Mapping rule

```text
Allocated Amount = BOQ Amount × Share %
```

Example:

```text
BOQ item 100,000
  → A1020 40% = 40,000
  → A1030 60% = 60,000
```

## Deliberately deferred

- Save / Load / Auto-save Mapping Session — MS-5
- Clear All with recovery — MS-5
- Final workbook workflow — MS-6
- On-demand S-Curve generation — MS-7

## UI refinement

The BOQ table keeps monetary values for **Amount** and **Allocated**, while **Remaining** is displayed as an integer percentage. This makes the remaining allocatable share immediately visible without adding GUI progress bars.

Example: `Amount 189,225.39 | Allocated 132,457.77 | Remaining 30% | Partial`.
