# LW-0 — Rebuild / Export UX Contract

LW-0 changes presentation ownership only. It does not introduce the Live Workbook engine.

## Contract

- Remove the standalone **Export** workspace from the sidebar and Tools menu.
- Keep initial mapped-workbook export inside the **Create Progress / Mapping** flow.
- Rebuild has two independent selections:
  - **Output Mode**: Snapshot Workbook / Live Workbook.
  - **Rebuild Scope**: Progress / Payment.
- **Snapshot Workbook** continues to route to the existing production rebuild engine.
- **Live Workbook** is visible as the future mode, but LW-0 blocks execution so it cannot silently route through the Snapshot engine.

## Test freeze

Pre-LW UI assertions that require a standalone Export workspace or the old RB6 combined mode layout are kept as frozen historical tests and explicitly skipped where the contract is obsolete. Core rebuild/payment behavior remains tested.
