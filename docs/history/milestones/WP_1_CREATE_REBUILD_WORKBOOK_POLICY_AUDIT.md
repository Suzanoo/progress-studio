# WP-1 — Create / Rebuild Workbook Policy Audit

Status: **Audit complete — no production behavior changed in WP-1.**

Baseline: `normalizer-v1-stable` (`320e49d`)

## Purpose

Progress Studio currently creates a usable workbook through more than one pipeline. The XML Normalizer work (N1–N8) standardized schedule input, but it intentionally did not standardize the final workbook policy. WP-1 identifies where Create Progress, Mapping Export, Snapshot Rebuild, and Live Rebuild already agree, where they diverge, and which behavior should become shared in the next workbook-policy milestone.

The audit keeps three concerns separate:

1. **Calculation/rendering engines** — how weekly/monthly/progress/dashboard data are derived.
2. **Final workbook policy** — guide, sheet visibility, cell/sheet protection, and recalculation mode.
3. **Desktop workspace UI** — which buttons belong in each workspace.

WP-1 does **not** change any of these behaviors. It establishes ownership before refactoring.

## Current pipelines

### A. Create Progress from XML

`WorkbookGenerationService.generate()` currently runs:

```text
Normalized Schedule
  -> ImportWorkbookWriter
  -> ScheduleWorkbookService
  -> TimescaleService
  -> AmountService          (fake / supplied amount behavior remains downstream)
  -> ProgressService
  -> DistributionService
  -> OkdService
  -> MonthlyMainService
  -> copy final workbook
```

Important current behavior:

- `ProgressService` creates `progress`, `progress_table`, `Dashboard`, and `Dashboard_Data`.
- `ProgressService` applies the Activity Data hierarchy theme and incremental recalculation policy.
- `MonthlyMainService` creates `main_monthly`.
- The final Create Progress copy step does **not** call the shared final visibility policy.
- The final Create Progress copy step does **not** call the shared final protection policy.
- The final Create Progress copy step does **not** create the workbook `README` guide.
- The final Create Progress copy step does **not** call `build_traditional_overlays()`.

This means the workbook created directly from XML is not finalized by the same policy that Rebuild/Mapping Export already use.

### B. Mapping Export

`MappedWorkbookExporter` already performs a finalization pass after mapping:

- rebuild/refresh relevant dashboard data,
- configure incremental recalculation,
- `apply_final_sheet_visibility()`,
- `apply_final_sheet_protection()`.

Therefore a workbook can change its protection/visibility contract simply by passing through Mapping Export, even when the underlying schedule is the same.

### C. Snapshot Rebuild

`WorkbookRebuildEngine.rebuild_progress()` treats `main` as authoritative and regenerates:

```text
main_monthly
progress
progress_table
Dashboard_Data
Dashboard
```

It then applies:

- incremental recalculation,
- `apply_final_sheet_visibility()`,
- `apply_final_sheet_protection()`.

Snapshot Rebuild does **not** currently create the workbook `README` guide and does not call the traditional overlay renderer in this path.

### D. Live Rebuild

`WorkbookRebuildEngine.rebuild_live_progress()` uses the newer live path:

```text
main -> MainDataset
     -> monthly cache / main_monthly
     -> progress + Dashboard_Data + Dashboard
     -> traditional overlays
```

It then applies:

- `build_workbook_guide()`,
- `apply_final_sheet_visibility()`,
- `apply_final_sheet_protection()`,
- live save recalculation policy.

The Live Progress result deliberately does not own `progress_table`.

## Existing final visibility contract

The existing shared policy in `workbook_visibility.py` is:

| Sheet group | State |
|---|---|
| `README`, `main`, `main_monthly`, `Payment Input`, `Payment`, `Dashboard` | Visible |
| `progress`, `progress_table`, `Dashboard_Data` | Hidden |
| Every other support/internal sheet | VeryHidden |

This policy is currently applied by Rebuild, Mapping Export, and Payment workflows, but not by the final Create Progress stage.

### Audit note

The policy makes **all unknown/support sheets VeryHidden**, including sheets such as `Info`, `Timescale Info`, `Amount Mapping`, and `Distribution Report`. This is stronger than a normal Hidden policy and should be an explicit product decision in WP-2 rather than an accidental consequence of whichever pipeline last saved the workbook.

## Existing final protection contract

The shared policy in `workbook_protection.py` currently protects every sheet with the internal workbook sheet password while leaving selected inputs editable.

### `main`

Editable non-formula Activity cells include the configured Plan identity/schedule/value fields, Actual fields, and Activity timescale cells. Project/WBS rollups and formulas remain locked. Row insert/delete remains available on `main`.

### `main_monthly`

Protected/read-only except the local cutoff control in column M.

### `Dashboard`

Protected, with interactive controls unlocked:

- `G5` — Weekly / Monthly view,
- `K5` — Cutoff Date,
- `P37` — Activity Status Focus.

### `Payment Input`

Hierarchy/identity cells are locked; payment percentages are unlocked for ACT rows only.

### Workbook structure

Workbook structure is intentionally **not** protected (`lockStructure=False`).

### Audit finding

This is already a coherent shared policy. The main problem is **application timing**: Create Progress does not apply it at final output, while Rebuild/Mapping/Payment do.

## Recalculation policy audit

There are currently two intentional recalculation modes:

- **Incremental/manual editing policy** for snapshot-style workbooks.
- **Live save recalculation policy** for the Live Workbook path.

WP-2 should preserve this distinction. Recalculation mode belongs to the selected workbook mode, not to XML source type.

## S-Curve / margin / cutoff audit

These behaviors are calculation/rendering policy, not protection/visibility policy.

N7.1/N7.2 fixed the Create Progress dashboard boundary so display margin is not treated as reporting data and Actual is bounded by cutoff. Those changes live in the shared dashboard/progress code used by the current Create path.

However, calculation ownership is still duplicated across:

1. Create Progress (`prepare_progress_and_scurve` + distribution + monthly build),
2. Snapshot Rebuild (`build_progress_views_from_source` + snapshot monthly + dashboard),
3. Live Rebuild (`MainDataset` + live monthly/dashboard + overlays).

WP-1 does **not** merge those calculation engines. Doing so together with protection/visibility would create an unnecessarily large refactor. Any remaining curve divergence should be fixed as a dedicated calculation milestone, not hidden inside workbook policy work.

## `progress_table` audit

Current ownership is intentionally inconsistent by mode:

| Pipeline | `progress_table` |
|---|---|
| Create Progress | Generated snapshot |
| Mapping Export | Preserved/refreshed snapshot behavior |
| Snapshot Rebuild | Regenerated snapshot |
| Live Rebuild | Removed / not generated |

Decision: **defer**. The project has already explored removing and restoring this snapshot. WP-2 must not decide this implicitly.

## Desktop workspace UI audit

The desktop command bar is global. It currently displays Mapping commands while the Create Progress workspace is active:

```text
Undo | Map | Unmap | Export Mapped Workbook
```

The Create Progress heading also exposes `Export Mapped Workbook`, producing duplicated Mapping ownership in the Create workspace.

This is a presentation-layer issue, not workbook policy. It should be handled in a separate workspace UI milestone after WP-2.

Recommended Create Progress workspace ownership:

```text
Create Progress
  Schedule XML
  Weekly cutoff
  Fallback amount
  Plan distribution
  Create Progress Workbook
  Open output workbook
  Open output folder
  Go to Mapping
```

Mapping-only actions should appear only when Mapping is active.

## Ownership matrix

| Capability | Create XML | Mapping Export | Snapshot Rebuild | Live Rebuild | WP-1 decision |
|---|---:|---:|---:|---:|---|
| Normalized MSP/P6 input | Yes | N/A | N/A | N/A | Frozen at N8 |
| Weekly `main` | Create | Preserve/update amount | Preserve source-of-truth | Preserve source-of-truth | Keep |
| `main_monthly` | Create | Preserve | Rebuild snapshot | Rebuild live/cache | Calculation concern; do not unify in WP-2 |
| Dashboard | Create | Refresh | Rebuild | Rebuild live | Keep mode-specific builders |
| Actual cutoff | Yes after N7.2 | inherits workbook | Yes | Yes | Shared behavioral contract; test, do not rewrite in WP-2 |
| ± display margin | Yes | inherits workbook | mode-specific | mode-specific | Calculation concern |
| Traditional overlay | Not finalized here | inherited if present | No explicit render | Yes | Audit only; decide in renderer milestone |
| `progress_table` | Snapshot | snapshot behavior | Snapshot | None | Deferred |
| Workbook guide `README` | No | No explicit final creation | No | Yes | Candidate shared final policy |
| Final visibility policy | **No** | Yes | Yes | Yes | **Move to shared finalization** |
| Final sheet protection | **No** | Yes | Yes | Yes | **Move to shared finalization** |
| Workbook structure lock | No | No | No | No | Keep unlocked for now |
| Recalc mode | Incremental | Incremental | Incremental | Live save | Keep mode-specific |
| Workspace command ownership | Mixed | Mapping | Rebuild | Rebuild | Separate UI milestone |

## WP-2 target boundary

WP-2 should introduce one shared **final workbook policy** entry point, without merging calculation engines.

Conceptually:

```text
Create Progress -----------\
Mapping Export -------------+--> Final Workbook Policy
Snapshot Rebuild -----------+      - guide policy
Live Rebuild ---------------+      - visibility policy
Payment workflows ----------/      - protection policy
                                    - mode-specific recalc policy
```

The recommended boundary is a small orchestration function/service such as:

```text
finalize_workbook(workbook, mode=...)
```

It should call existing proven helpers rather than reimplement their rules.

### WP-2 must not include

- rewriting MSP/P6 normalization,
- changing fake amount behavior,
- merging Snapshot and Live calculation engines,
- deciding `progress_table` ownership,
- redesigning Create/Mapping workspace UI,
- changing S-Curve mathematics unless a separate regression demonstrates a calculation bug.

## WP-1 conclusion

The main architectural drift is not that Rebuild has an entirely separate protection system. The reusable protection and visibility helpers already exist. The drift occurs because **Create Progress never performs the same final workbook-policy pass**.

Therefore the lowest-risk next refactor is:

1. keep calculation engines intact,
2. centralize final workbook policy,
3. route Create Progress through that policy,
4. preserve mode-specific recalculation,
5. leave `progress_table` and workspace UI for later milestones.
