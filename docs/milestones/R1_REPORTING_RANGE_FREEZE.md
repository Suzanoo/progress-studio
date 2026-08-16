# R1 — Reporting Range Freeze

R1 freezes the boundary between the visible workbook timescale and reporting data.

## Contract

- `main` / `main_monthly` may keep the visible pre/post timescale margin.
- `progress` may keep the complete visible weekly history used by the live workbook contract.
- `Dashboard_Data`, dashboard KPI sources, and chart reporting lists must contain project reporting periods only.
- A weekly reporting period is real when its seven-day interval overlaps Project Start..Project Finish.
- The final overlapping week is retained even when its cutoff falls after Project Finish.
- Snapshot Progress rebuild and Live Progress rebuild use the same reporting-period overlap rule.
- Snapshot/Live Payment rebuild must preserve `Dashboard_Data`; Payment does not own reporting-range generation.

## 009 regression

`009-MSP_test_progress_live(1).xlsx` contains 68 visible weekly periods (`W1..W68`) including margin.
The project reporting range is 60 periods, `W5` (17-Apr-2026) through `W64` (04-Jun-2027).
After R1 Live Progress rebuild, `Dashboard_Data` contains only those 60 Weekly links and 15 Monthly reporting points.

R1 does not change Normalizer, S-Curve distribution, Payment calculations, final workbook policy, or Rebuild ownership.
