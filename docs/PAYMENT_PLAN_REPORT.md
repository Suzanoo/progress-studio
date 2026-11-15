# Payment Plan Report

PR-0 design accepted. PR-1 provides the read-only Python foundation for review. One project-wide period schedule and one selection feed the future period and cumulative Word documents. Desktop preview and selection (PR-2), Word export and dynamic pagination (PR-3) are pending. WF-1 and WF-2 are deferred.

## Source

Call `progress_studio.infrastructure.excel.payment_plan_source.read_payment_plan_source(path, expected_revision=None)`. The adapter reads a single byte snapshot, fingerprints it with SHA-256 and uses embedded Payment Input, main and Info from that revision. Reuse the returned source for subsequent previews. Compare the expected digest on reopening. Selection keys are scoped to that project/revision. The caller must explicitly rebind and review a selection when the revision changes.

Payment Input owns the locked Activity Amount, sparse entries, row names and schedule order. Header row 6 discovers Type, WBS, Activity ID, Activity Name, Amount and P-number columns. Preserve nonconsecutive period IDs in ascending numeric order. Project root and matching main Plan outline levels establish hierarchy, including levels beyond Excel's outline limit. Validate every reporting row against main Plan order, identity, code and description. Compare locked and current Activity Amounts as stored numbers before aggregation. Duplicate identities, malformed requirements, missing headers, changed hierarchy and stale snapshots raise actionable errors. If both source Amounts are unavailable, retain the row, diagnostic and available percentages. Missing numeric input never supplies a zero Amount.

Supported sources carry XML Amount creation evidence or the existing BOQ Activity Mapping workflow. A standalone legacy input, BOQ report snapshot, missing project metadata, manual placeholders and unsupported weights require a prepared monetary source. PR-1 does not infer currency. Formula-based Amounts and requirement inputs are unsupported, including those with saved caches. Use a verified literal source. A saved workbook can establish internal consistency, but cannot certify external commercial revisions.

## Dates

Use the unchanged Payment sparse reader, weekly Plan index and Payment position resolver across the entire project. The marker comment and label are saved renderings. The resolver supplies the planned eligible date. Row 7 is a reference date only. A missing period input, unresolved controlling point, or formula/unverified weekly distribution in a required activity yields a missing date with issues. A resolved date earlier than a prior validated period is flagged and withheld while amounts remain available. Duplicate dates are valid. Selection never recomputes dates.

## Calculation and selection API

A ReportSource provides project, filename, revision, monetary origin, ordered periods, rows, source row numbers, hierarchy, diagnostics and leaves. Each Leaf includes the raw Amount and a tuple of shared RequirementState values with explicit, before-first, carried or unknown origin and source coordinates. WBS key segments encode the ancestor WBS path. Activity keys carry the exact Activity ID. The project key is PROJECT. Keep these returned keys opaque.

Call `source.select(keys, revision=source.revision)`. Returned rows stay in schedule order; included_in points to the nearest selected ancestor. The union of distinct descendant Activity IDs forms selected scope. Its complement and project scope are available on the same result. Duplicate requested keys count once. An empty selection covers zero activities. Full-project coverage compares all identities, including zero-Amount work.

Activity cumulative money is raw Amount × cumulative fraction. Period values are signed differences from previous cumulative values with a zero opening balance. WBS and scope fractions use all descendant Amounts as a fixed denominator. Full amounts and percentages remain unavailable for incomplete positive-weight work. Known subtotals, known monetary basis, known-leaf count and total scope identities expose partial coverage. Zero basis yields no weighted percentage. A single zero-Amount activity retains its own valid requirement percentage. Calculations use Decimal source numbers without display quantization (80 significant working digits for aggregation and weighted division). Report values describe planned requirements. Eligibility, physical progress, certification and cash transactions retain their own meanings.

## Visibility

Compare the unrounded cumulative fraction to 0.9999, inclusively. At or above it, show cumulative 100% and keep the threshold-crossing row visible. Repeat unchanged completed detail is suppressed. Any genuine subsequent raw movement, including above-threshold adjustments in either direction, remains visible with its true signed delta and amount. Dropping below the threshold reopens the row. A later crossing is recompletion. Unknown or invalid values are unresolved. Display rounding never decides completion. Consumers must use the returned transition and display fraction; they must not format below-threshold values with a less-than qualifier. Period percentage is always the raw period delta.

Per-period Value supplies raw fraction, delta, money, movement, display_fraction, transition, correction and both visibility flags. Known zero movement can be omitted in the period view. Selection.retained_money includes suppressed independent reporting components. visible_known_money gives the remaining known cumulative contribution. Their sum reconciles to the selected known money when available. Scope totals are computed before suppression. Overlapping explanatory detail contributes no second amount.

## Verification

Run the payment unit and integration suites. The separate PR-1 engineering evidence runner uses the six identified NKC1 files. It checks 239 × 28 Activity values, every PR-0 CSV row, full and partial selection, verified dates, unrounded money and boundary-specific visibility. Synthetic tests own unavailable states, stale and invalid inputs, sparse zero, threshold boundaries, adjustments, reopening and recompletion.
