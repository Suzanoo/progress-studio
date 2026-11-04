# BOQ Update and XML Amount Writeback — engineering plan

Date: 2026-11-01. Reviewed baseline: `main @ 78675f524933b691b0d61485fc347b2bdfe9854c`
(`wu2-duration-columns-stable`).

## Authority and status

The Product Owner accepted BU-1 automated, Windows Desktop and product acceptance,
and authorized **XW-1 only** on 2026-11-04 against `main @ de125d4`, tag
`bu1-boq-amount-update-stable`. X1–X5 and the XML saved-formula policy below are
accepted. XW-1 is implemented for engineering handoff, not PO-accepted or released.
**XW-2 and BR-0 remain unauthorized for implementation.**

See [P6 Amount writeback engine](XML_AMOUNT_WRITEBACK.md) for the implemented
boundary, reproducible proof and limitations. The Engineer has not performed Git
writes or native P6 acceptance. Coordinator owns separately authorized integration.

This file owns this initiative's accepted direction, proposed plan and decision
register. [Roadmap](ROADMAP.md) records status; [Next Steps](NEXT_STEPS.md) links
here rather than duplicating decisions. Existing [Architecture](ARCHITECTURE.md),
[Weight Selection](WEIGHT_SELECTION.md), [EV Monetary Sources](EV_MONETARY_SOURCES.md)
and role contracts remain authoritative for existing behavior.

## Accepted product direction

- BOQ Stage 1: update Amounts while preserving valid Mapping allocations;
  show changes and protect existing project data against unintended loss.
- BOQ Stage 2: subsequently address structural item/organization revisions;
  investigate identity, reconciliation and safety before implementation boundaries.
- XML: P6 first; one Progress workbook and one target XML per operation; review
  proposed changes before export; preserve the original XML. Transfer selected
  Workbook Activity Amounts to matching P6 Activities, including a valid subset
  of a combined project.
- Engineer independently recommends order, dependencies, scope and gates.
  Product semantics are not established by implementation convenience.
- XW-1 alone is currently authorized for implementation. Engineer delivers;
  Coordinator integrates only after PO review and explicit Git authorization.

## Completed evidence and relevant delta

Completed external reports, both investigated at `5bb7f37`:

- `XML_AMOUNT_WRITEBACK_REPOSITORY_INVESTIGATION.txt`
- `BOQ_UPDATE_RECOVERY_AND_INVESTIGATION.txt`

Their case-specific file operations and direct assertions are evidence, not a
shipped feature, native acceptance, general engine or binding technical design.
The BOQ recovery preserved 2,415 pairs over 1,050 items for nine Amount changes.
The XML investigation used 239 matching source IDs and reported formula-valued
Amounts; the example UDF ObjectId is not a reusable identity constant.

The comparison `5bb7f37...78675f5` contains two commits and 13 changed files.
WU-2 adds duration rendering, changes Create/Mapping/Monthly/Compact render paths
and working-tree schedule metadata, and adds duration regression coverage.
The session repository, MappingStore, BOQ reader, main dataset reader, P6 adapter
and XML Amount validation are unchanged. Earlier findings remain applicable.

WU-2 matters to both features: use header discovery with old and new layouts,
preserve paired Actual dates and explicit milestone identity, and retain duration
columns/formulas through Mapping export. `duration_workbook.py` stores explicit
milestone markers in Info; `working_tree_schedule_source.py` uses them with a
legacy monetary-input fallback. A zero Amount or same-day date is not sufficient
milestone evidence. Conflicting/missing source and target classification needs
validation, not a silent interpretation.

## Remaining work and reuse boundaries

### BOQ

`amount_mapping.py` currently falls back to the embedded snapshot when source
validation fails. Browsing/loading another BOQ calls `MappingStore.load_boq`,
which clears allocations. These are source restoration/loading behaviors, not
a revision reconciliation flow.

`mapping_reader.py` keys include source/current row positions; its stable ID also
contains source row and descriptive fields. Neither guarantees identity across
reordering. The reader drops non-positive Amount rows and does not carry every
identity-relevant source column (such as unit/quantity) into BOQRow. Comparing
only currently parsed positive rows is insufficient to prove an Amount-only edit.

Reuse readers, share validation, MappingStore derivations, fingerprints,
snapshots and atomic session writing. Add the missing comparison, preview and
safe state transition at their appropriate existing boundaries. Legacy sessions
without an old snapshot require a verified original BOQ or a clear stop; do not
treat the proposed replacement as its own comparison baseline.

### XML

`RebuildWorkbookReader` and `MainDataset` are useful reading foundations, but
Amount is not required by the header finder, malformed numbers become None,
blank-ID rows disappear from the activities projection, and formula/cache evidence
is not retained there. Strict writeback needs validation before those losses.
Do not reuse amount readers that aggregate duplicate IDs or skip unusable rows.

`P6XmlAdapter` discovers typed Activity UDFs by identity and reads the first direct
Project. Its normalized output is not a lossless XML document. Writeback needs
target-document preservation, project-bound matching, typed-field resolution,
preview, a separate output and verification of every intended mutation.
Reuse AmountField and appropriate Decimal checks, but do not copy Create's
milestone-zero or positive-total policy into export automatically.

The workbook's main Amount is Progress Weight; independent EV Monetary Inputs
owns Activity Contract Value. Creation metadata describes origin, not current
monetary validity. Manual-origin workbooks may contain real money after editing.
X1 now resolves this boundary with explicit per-operation monetary confirmation;
existing EV source semantics remain unchanged.

## Recommended sequence and dependencies

Recommended delivery order: **BU-1, XW-1, XW-2, BR-0**, followed by separately
scoped structural implementation only after BR-0 review.

BU-1 and XW-1 have no hard code dependency on one another. BU-1 is recommended
first because it addresses observed data loss/staleness with a bounded real case
and fewer monetary-policy questions. XML work can use any qualifying workbook;
BOQ Mapping is not a prerequisite. In an end-to-end user workflow, an updated
mapping must be exported to a current workbook before that workbook feeds XML.
XW-2 depends on acceptance of XW-1. BR-0 is not a prerequisite for Amount-only
updates or XML; it precedes any structural BOQ implementation. This sequence is
not authorization, a fixed schedule or a requirement for new shared infrastructure.

### BU-1 — Amount-only BOQ Update (accepted and integrated)

Outcome: a user can review and accept a verified Amount-only revision, retain
allocation pairs and exact shares, save/reopen the project and export updated
mapped amounts through the existing workbook lifecycle.

Scope: comparison against the session's verified BOQ; changed-item and affected
Activity amounts; unchanged/unmapped counts and totals; explicit accept/cancel;
coherent fingerprint/snapshot/mapping persistence; failed-update recovery.
Recommend strict rejection of structural/ambiguous changes rather than best-effort
remapping. B1/B2 below are accepted for BU-1 only.
Detect source changes after preview and require fresh validation before applying.
An unchanged candidate should be a harmless no-op with truthful feedback.

Likely affected areas (not prescribed architecture): existing BOQ Excel adapter,
Mapping services/store, session persistence, Mapping presentation, focused tests
and user workflow documentation. Reuse atomic writing but prove the in-memory
and on-disk update behave coherently on failure; atomic JSON alone does not prove it.

Engineering gates: real nine-change fixture; exact pair/share and working-tree
preservation; correct Activity totals; no-op/cancel; failed validation/write;
snapshot fallback and missing sources; legacy comparison-source availability;
Save/Close/Reopen; source changed after preview; duplicate/reordered/added/deleted
rows, changed unit/quantity/hierarchy and excluded zero rows must not evade checks.
Reconcile new exports and preserve EV inputs, Actual dates, WU-2 durations and
unrelated workbook content. Verify a previous session remains recoverable.

Native/PO gate: user-visible preview/update/reopen/export on the supported desktop,
Excel first-open/F9/Save/Reopen of the mapped output, then PO acceptance. Engineer
performs all feasible automation and failure injection first; state unavailable
native checks explicitly rather than assigning routine QA to the PO.

Deferred: structural reconciliation, split/merge, automatic XML updates, changes
to standalone Rebuild, Payment prerequisites, EV source selection or historic BAC.

### XW-1 — P6 writeback engine proof (engineering handoff)

Outcome: a repeatable proof generates and validates a separate P6 XML for reviewed
selected workbook Activity amounts; engine inputs/preview/results are testable
without a desktop screen. X1-X5 are accepted below.
This proof is not a released user-facing workflow or general P6 round-trip claim.

Scope: strict source evidence, project/ID matching, typed UDF resolution, before/
after preview, preservation of unrelated XML, safe separate output and diagnostics.
Recommend exact case-preserved matching within an explicit project boundary,
one unique target per selected ID, replace rather than accumulate, and unchanged
target-only Activities. An unmatched selected ID blocks export, not silent omission.
The XW-1 engine validates every Activity/P row in the selected workbook. That
workbook may be a subset of the target combined Project; arbitrary row filtering
is not exposed by this proof.
Revalidate input identity after preview. Do not infer selected IDs from row order.

Likely affected areas: workbook-reading boundaries, existing amount/field domain
validation, P6 XML infrastructure, orchestration and focused fixtures/tests.
No need for a generic revision framework shared with BOQ by default.

Engineering gates: subset of combined project; blank/duplicate/case-colliding IDs;
old/WU-2 columns; literal/formula/error/missing/nonfinite/negative/zero values;
milestone conflicts; duplicate UDF declarations/values, type mismatches, nil values
and missing field creation; namespaces, unrelated records and target-only values;
large/precise values; repeated-export idempotence; cancel/failure/source mutation.
Prove only intended UDF declarations/values change, not dates, relationships,
calendars, resource/cost records or workbook bytes. Re-read values with existing
adapters where applicable, and reconcile the supported Create import path without
mistaking its intentional milestone-zero weights for failed monetary preservation.

Gate: engineering proof and artifact review, then PO acceptance. Native P6 import
and export evidence is required before claiming P6 round-trip compatibility; use
copies/a test project and state what is unverified. This can be completed in the
proof when available, but remains a release gate for XW-2 if unavailable.
Deferred: GUI release, MSP, batch targets, automatic source merging, renumbering,
schedule editing and writing arbitrary native cost/resource fields.

### XW-2 — Reviewed desktop export (proposed)

Outcome: one workbook/one target workflow exposes validated selection, source/
field/project identity, before/after values and totals, exclusions and blockers,
and exports only after review. Depends on accepted XW-1; menu placement is open.

Likely areas: presentation, XW-1 orchestration integration, translations, workflow
documentation and integration tests. Preserve ENG/THA conventions. No second
implementation of the engine inside the GUI.

Gates: invalid input feedback; stale-preview invalidation; cancel and output
collision handling; no overwrite of either input; end-to-end desktop export;
unchanged source workbook/XML bytes; P6 compatibility gate above; existing
Create/Mapping/Rebuild/Payment/EV and WU-2 regression; PO acceptance.
Deferred: batch/MSP and broader monetary-source policies not explicitly approved.

### BR-0 — Structural BOQ identity and reconciliation design (proposed later work)

Outcome: independently reviewed identity/reconciliation design with a change
matrix and representative fixtures for inserted/deleted/reordered/renamed,
duplicate, split and merged items. Assess BOQ Studio export identity and legacy
input compatibility; do not assume durable IDs exist or require a cross-repo
change without separate authorization.

Likely areas: design/decision documentation and test evidence, not a shipped
structural updater. Gate: demonstrate unambiguous vs ambiguous matches, define
recoverable handling of displaced allocations, and obtain PO decisions for
remapping outcomes. Then propose implementation milestones based on evidence.
Deferred: production structural reconciliation. Stage 2 direction alone does
not authorize BR-0 execution or any subsequent implementation.

## Decision register — B1/B2 and X1–X5 accepted

| ID | Accepted decision | Milestone |
| --- | --- | --- |
| B1 | Zero transitions: initially support changes between positive amounts on existing eligible items; unchanged zero/non-item rows stay unchanged. Block newly eligible rows or positive-to-zero transitions explicitly until a zero-item retention policy is approved. Accepted for BU-1 by the implementation authorization. | BU-1 |
| B2 | Formula Amounts: accepted for BU-1: saved numeric results with visible formula count and a recalculate/save prerequisite; reject missing/error results. Cached-result freshness cannot be independently guaranteed. XW-1 also accepts saved numeric formula results, requires explicit recalculation/Save confirmation, rejects missing/invalid/error caches and warns that freshness cannot be guaranteed. | BU-1 / XW-1 respectively |
| X1 | Source meaning: use selected current main Activity/P Amounts with explicit per-operation monetary intent; do not infer money from origin/positive values and do not silently merge EV Monetary Inputs or Mapping Summary. Creation origin must not permanently disqualify later real money. | XW-1 |
| X2 | Milestones: preserve target monetary values and exclude milestone writes initially; detect unknown/conflicting identity and block affected selection. Future milestone money export needs a deliberately chosen monetary source. | XW-1 |
| X3 | Target field: use an explicit typed Activity Cost UDF, reusing the intended existing field only after identity/type validation, or offering creation of Progress Studio Amount with a fresh unused ID. Never hardcode the case-study ID or silently change native cost fields. | XW-1 |
| X4 | Project scope: one direct P6 Project initially; a combined building subset within it is supported. Multiple Project elements should block with explanation unless explicit project selection is approved. A combined project is not necessarily multi-project XML. | XW-1 |
| X5 | Numeric policy: individual zero and all-zero selected subsets be valid; blank is not zero; reject negative/nonfinite/invalid values; preserve supported stored precision without deliberate cents rounding. Values replace existing amounts; they are not added. | XW-1 |

These are product-visible scope/meaning choices, not requests for the PO to choose
classes, algorithms or test tooling. Exact diagnostics, implementation structure,
adapter reuse and verification methods remain engineering responsibilities.

## Preservation and acceptance boundaries

- BOQ update changes the mapping project's BOQ revision; old exported workbooks
  and XML remain unchanged. Use existing export/refresh ownership explicitly.
- XML export reads the selected current workbook; it does not rebuild progress,
  alter creation basis, populate independent EV inputs or refresh existing EV.
- Milestone Progress Weight stays zero; monetary values and BAC remain distinct.
- Standard Payment still requires no real Activity Amounts or BOQ. Payment
  Breakdown and EV retain their separate readiness/identity contracts.
- Rebuild remains workbook-based and independent of original XML/session/BOQ.
  Preserve WU-1 cutoff, WU-2 durations/Actual dates and accepted worksheet behavior.
- Failed/cancelled operations must preserve prior usable state. Verify output
  collision handling and original recovery; never quietly discard unmatched data.

BU-1 adds strict selected-sheet evidence comparison, staged MappingStore updates,
review, a sibling recovery project and atomic session publication. Reviewed BOQ
fingerprints use binary identity: semantic identity omits formula caches, so an
externally resaved file must fall back to the embedded reviewed snapshot.
No session format migration or XML/EV/Payment policy change is introduced.
See the delivered BU-1 implementation report for actual engineering results.
Every implementation milestone needs its own engineering results, applicable
native evidence, PO acceptance and separately authorized Git integration.
