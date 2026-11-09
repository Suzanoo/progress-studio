# Progress Studio — Roadmap and milestone status

**Documentation synchronization:** 2026-11-09. **Reference:** accepted `main @ 15d8568` (BP-4 UI refinement). This is a source/documentation status record, **not** a claim that a new installer or public release was published. Accepted PO evidence and milestone tags remain authoritative; the current HEAD must be verified before integration.

[← Product overview](../README.md) · [Documentation index](README.md) · [Open investigations](NEXT_STEPS.md)

## Accepted and closed

| Area | Milestone | Accepted evidence / reference |
| --- | --- | --- |
| Product foundation | V1, V1.1 | `progress-studio-v1-stable`, `progress-studio-v1.1-stable` |
| Progress reporting | Weekly Compact V1 and Final Presentation | `weekly-compact-v1-stable`, `weekly-compact-final-presentation-stable` |
| Amount workflow | A1 Amount-First; WU-1 Cutoff Placement; WU-2 Duration Columns | `amount-first-a1-stable`, `wu1-cutoff-placement-stable`, `wu2-duration-columns-stable` |
| BOQ Update | BU-1 Amount-only BOQ Update | `bu1-boq-amount-update-stable` |
| P6 XML | XW-1 Engine and XW-2 Desktop Integration | `xw1-p6-xml-writeback-stable`; XW-2 accepted as the baseline for subsequent PI-1 work |
| Payment Input | PI-1 Activity Amount / WBS Rollup | `pi1-payment-input-rollup-stable` |
| BOQ Payment | BP-1 Projection | `bp1-boq-payment-projection-stable` |
| BOQ Payment | BP-2 Separate Workbook Export | Accepted before BP-3 work; see [BOQ Payment Export](BOQ_PAYMENT_EXPORT.md) |
| BOQ Payment | BP-3.2 Amount Completion | `bp32-amount-completion-stable` |
| BOQ Payment | BP-4 Manual BOQ Requirements | `bp4-manual-boq-requirements-stable`; UI refinement `main @ 15d8568` (PO Native UI PASS) |

The accepted BP-4 workflow is **Prepare BOQ Requirements → BOQ Requirement Input → Save → Export BOQ Payment Workbook**. Activity-mapped and manual known planned requirements are composed in the report with separate origin and coverage evidence. BOQ projections are **not certified payment**.

## Open, deferred, or separately authorized

| Topic | Status / boundary |
| --- | --- |
| **XW-DP — XML Decimal Precision** | PO approved two-decimal XML output direction; implementation and regression acceptance **not completed**. Do not claim shipped. |
| **Weekly Compact scrolling performance** | Investigation/proof requested; no accepted production change established in this status record. Preserve accepted visuals and live behavior. |
| **NS-1 — Rebuild weight-basis conversion** | Open investigation question; no automatic conversion authorized. |
| **NS-2 — Actual progress preservation on Rebuild** | Open investigation question; no conflict policy authorized. |
| **XW-3 / BR-0 structural BOQ work** | No production implementation authorization. |
| **DOC-1 — GitHub presentation and documentation** | Documentation-only refresh in preparation; no code/release changes. |

## Explicit exclusions

- **MS-3 Mini Gantt:** cancelled by PO; do not list as an upcoming feature.
- Finance/AI navigation: hidden in the accepted V1.1 desktop product; retained Finance implementation is not advertised as an active UI feature.
- Native P6 import may apply its own precision behavior; acceptance of XW-1/XW-2 does not establish the separately proposed XW-DP two-decimal rule.

## Historical evidence

Historical accepted test counts are retained in [Changelog](../CHANGELOG.md) and feature contracts. They are **not** a claim of a current full-regression run. The [pre-production roadmap](history/PRE_PRODUCTION_ROADMAP.md) is archival and does not authorize old milestones.
