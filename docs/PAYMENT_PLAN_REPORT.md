# Payment Plan Report

PR-0, PR-1 and PR-2 are accepted and integrated on the feature branch. PR-3 paired Word export is submitted for acceptance. WF-1 and WF-2 remain deferred.

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

## Desktop selection and preview (PR-2)

Open **Payment → Payment Plan Report → Open payment workbook**. Review the project, file, monetary basis, period count and Source details. The dialog uses one read-only PR-1 snapshot. Opening a new file happens in the background. Failed opens preserve the previous source, reporting rows and preview. Opening a different revision with an active selection asks before switching.

In **Select work**, expand a branch using its arrow. Click its checkbox or focus a row and press Space. Select project row replaces the selection with the project summary. Selected detail inside another selected row is labelled with its parent. Other represented rows show where their amount is included. Search keeps matching rows and their ancestors. Clearing search restores the prior expansion state. **Review selected rows** lists every explicit row regardless of filtering, collapse or preview suppression. The footer shows unique Activity count and fixed monetary basis against the whole project.

In **Preview**, choose the period and either Period or Cumulative. The project-wide eligible date, flags, signed movements, suppression, retained cumulative money, known coverage and totals come from PR-1. A period with unresolved data remains accessible. Source details lists the diagnostic messages and the workbook fingerprint. Workbook descriptions, monetary provenance and underlying source diagnostics retain their source wording in either interface language.

**Save selection** writes a separate versioned .paymentselection.json containing the explicit row keys, project, filename and source SHA-256. Open its matching workbook before **Restore selection**. Renaming or moving an identical workbook is supported. A changed revision, wrong project, malformed file or unknown key leaves the active selection intact. To use a revised workbook, review the new source and select its rows again. Selection files contain no copied payment amounts.


## Export both Word reports (PR-3)

From Payment → Payment Plan Report, open the saved workbook, select or restore reporting rows, then choose **Export both Word reports**. Choose an existing destination directory and enter a new folder name. The app creates that folder with two editable .docx files using its name as a prefix. For example, folder NKC1 contains NKC1_PERIOD_PAYMENT_PLAN_REPORT.docx and NKC1_CUMULATIVE_PAYMENT_PLAN_REPORT.docx. Both use the entire validated period sequence regardless of the preview tab or period.

The exporter captures one PR-1 selection and source revision. It checks the workbook digest before generation and again before publication. Empty selection, a changed source, an invalid destination, or an existing folder stops publication. Choose another name to keep prior outputs. A staged sibling directory holds both documents until they pass ZIP integrity checks; a single directory rename publishes the pair. Cancel during generation discards the staged pair. Final publication is a short non-interruptible filesystem operation. Use a writable directory and leave room for the pair. Install the declared python-docx dependency in the application environment. Microsoft Word is needed only for native acceptance, editing or printing.

Reports use English titles and labels with original project descriptions, including Thai. They show a scope register, source revision and basis, each planned eligible date, explanatory overlapping rows, signed movements, period issues and the independent selected total. Cumulative retention is shown even when all completed detail is suppressed. Missing final figures are labelled Unavailable; available known cumulative money and coverage are identified separately. Percentage/monetary formatting is for display. A4 portrait pages use dynamic Word table flow, repeated column headers, page numbers and kept totals. The first period begins after the register. No fixed number of periods per page is assumed.

Verify source integrity, complete chronological sections, preview-to-document agreement in both modes, repeated/completed rows and retained amounts, signed adjustments, unknowns, cancellation and collisions. Render real NKC1 and a long multipage selection. PO acceptance covers the Desktop export, practical editing of both documents, pagination and print preview in Microsoft Word.

### Product Owner approved deferrals

- PR-I01 — Revision-aware selection restoration. The exact-revision policy remains active.
- PR-I02 — Amount-First compatibility with workbook-entered Activity Amounts remains unverified.
- PR-I03 — Compatibility with formula-based Amounts and requirements remains unverified beyond the current source restrictions.
