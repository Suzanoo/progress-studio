# MS-3 Acceptance — Timescale and Excel Infrastructure

MS-3 passes only when all of the following are true:

1. `TimescaleStep` is part of the application pipeline.
2. Stage 03 is executed by `TimescaleService`, not by subprocess.
3. Timescale workbook logic lives under `infrastructure/excel`.
4. The legacy continuation begins at stage 04.
5. V2 stage-03 workbook matches the V1 baseline for sheets, dimensions, values, formulas, number formats, merges, freeze panes, filters, and outline levels.
6. No Thai text exists in application source.
7. All automated tests pass.

Out of scope: BOQ mapping, progress calculation, distribution, OKD sheets, and removal of scripts 04-07.
