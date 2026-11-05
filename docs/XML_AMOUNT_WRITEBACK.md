# P6 XML Amount writeback engine — XW-1

Status: XW-1 accepted, integrated and closed at `main @ c928d59`, tag
`xw1-p6-xml-writeback-stable`. Product Owner evidence records Windows automated
verification, real NKC1 proof and native Primavera P6 compatibility acceptance.
XW-2 Desktop integration is implemented for engineering handoff; its native
Windows usability, Product Owner Acceptance and Git integration remain pending.
X1–X5 are accepted in [the initiative plan](BOQ_UPDATE_XML_WRITEBACK_PLAN.md).

## Desktop operation — XW-2

Open **Tools > Export Amounts to P6 XML...** from any workspace. One resizable
operation dialog contains file selection, confirmations, review and export.
No new sidebar/Home/Rebuild route or Project Session requirement is added.

1. Explicitly choose a saved XLSX/XLSM Workbook and a target P6 XML. Nothing is
   inferred from Mapping, Project Session, an embedded snapshot or recent exports.
   After Mapping edits, first use Export Mapped Workbook and select that output.
2. Field discovery reads through XW-1 assessment. It is never export approval.
   Select a named Activity / Cost field (even if only one) or explicitly request
   creation of Progress Studio Amount. Field numbers distinguish duplicate labels;
   Field details exposes technical identities as secondary evidence. Review a
   candidate to inspect its existing values; the engine determines compatibility.
3. Confirm main Amounts represent money. If formulas are detected, recalculate
   all relevant formulas in Excel, Save, and confirm this prerequisite. Missing or
   invalid caches remain blockers. Cache freshness cannot independently be proven.
4. Review Amounts: inspect Project ID, field, Activity IDs/names, exact stored
   before/proposed values, replacements, unchanged ordinary rows, preserved
   milestones, ordinary zeros and target-only IDs. Missing is distinct from zero.
   Target-only values/names are explicitly not assessed by the engine; these
   records remain unchanged. Select a row to read full values in the details area.
   Totals cover matched ordinary Activities, not Project BAC or Contract Value.
   Invalid assessments display diagnostic subtotals only and never allow export.
5. Export XML and select a NEW destination in an existing writable directory.
   This explicit action and Save selection approve only the displayed assessment.
   Existing destinations and input aliases are refused, never silently renamed or
   overwritten. Zero replacements may be exported as an explicitly requested copy;
   a reviewed new field is still created if requested. Success shows the actual
   path; importing into Primavera remains a separate operation.

Changing inputs, field or confirmation clears the review. External edits are
checked by XW-1 before publication. Failure requires a fresh review; the UI never
silently reassesses and exports an unseen result. Original diagnostics are retained
under localized ENG/THA recovery guidance. Field labels and customer data are not
translated. Language preference continues to apply on next application launch.

Indeterminate work runs off the UI thread, with results consumed by Tk's event
loop. Controls and duplicate export are blocked while busy. Cancel/close requests
cancellation and keeps the dialog/application alive until the worker returns.
Assessment cannot interrupt the accepted engine mid-read; its result is discarded.
An export which already published is reported as success, even after a late cancel.
Close again after the outcome is shown. No fabricated percentages or P6 import
success claims are shown. Force termination/power loss remains outside UI control.

## Operation and ownership

`XmlWritebackService.assess(WritebackRequest(...))` returns a frozen assessment:
input paths/hashes, Project, selected and available typed fields, source rows,
per-Activity before/after/status, target-only IDs, zero/formula/replacement/unchanged/
milestone counts, source/write/known-before totals, missing-before count, errors
and warnings. Totals on an invalid assessment are diagnostic subtotals, never
approval of valid rows. There is no partially accepted export.

`export(assessment, new_path, reviewed=True, cancelled=callback)` revalidates the
assessment and inputs, renders/verifies the XML and publishes a separate output.
Without explicit monetary confirmation, a valid assessment and review confirmation,
export fails. Cancellation is checked before validation, after rendering and before
publication. Callers must display the assessment and warnings before approving it.

All Activity/P rows in main form the selected source; A rows and WBS summaries
are not Amount sources. No BOQ/Mapping session is required and no creation/EV
metadata certifies that current amounts are money. Every source Activity must be
valid and exactly matched, even if only one row is invalid. A workbook can contain
only NKC1 while its target Project also contains NKC2. No arbitrary source-row
filter, fuzzy match, renumbering or monetary-source fallback is exposed.

## Supported evidence and numbers

- XLSX/XLSM sparse OOXML reading; header discovery in the first 30 rows, supporting
  old and shifted WU-2 columns. Amount/identity headers and cells must be unambiguous.
- Numeric literals and numeric saved formula results only. Text, booleans, errors,
  blank, negative and nonfinite Amounts block export, including milestone rows.
- Formulas require explicit `recalculated_and_saved=True`. Excel must recalculate
  and Save first. The engine does not calculate formulas, verify external links,
  or independently prove cache freshness. A valid cache is not a freshness claim.
- Stored digits are parsed directly as Decimal, never through float and never
  rounded to cents. Zero and all-zero sources are valid. Nonzero individual values
  use the existing Excel monetary range 2.2250738585072014E-308 through
  9.99999999999999E307. Numeric lexical length is limited to 1,024 characters and
  exponent syntax to four digits. Values outside support are rejected explicitly.
  Exact aggregate totals may exceed the individual Excel range; they are review
  totals, not written as additional XML values. Excel may already have rounded a
  stored number; writeback cannot recover precision Excel did not save.
- Activity IDs are trimmed and matched case-preservingly. Blank, duplicate and
  case-colliding source/target IDs are blocked. No joins by name/WBS/row position.
- Source milestone evidence comes from explicit EV Monetary Inputs flags and
  WU-2 Info markers. Only classification is read from EV, never Contract Value.
  Info markers do not certify unmarked Activities as ordinary. Missing, duplicate,
  invalid or conflicting flags block the operation; legacy workbooks without
  sufficient explicit evidence need corrected classification before export.
  Dates, duration and Amount zero are never milestone evidence.
- Target Start/Finish Milestone types must agree with workbook classification.
  Their existing UDF bytes (including missing values) are preserved. They are
  excluded from write/known-before totals; source_total still includes their main
  numeric inputs. Unknown target types are rejected.

## XML field and preservation boundary

UTF-8 Primavera P6/P6Professional versioned APIBusinessObjects XML with exactly
one direct Project is supported. DTD/entities, other encodings/dialects and multiple
Projects fail closed. Namespace prefixes, Unicode, comments, processing
instructions and foreign records are preserved outside the selected value spans.

Select `p6:udf:<ObjectId>` explicitly. Its one direct root UDFType must be Activity
/ Cost, and its selected Activity values must not have duplicate or conflicting
typed representations. Native P6 Cost fields are never candidates. Existing
ordinary amounts must be valid or missing; malformed existing values block review.
Missing/blank/nil ordinary values are shown as missing and may be replaced.

Alternatively request `create_progress_studio_amount=True` with no existing field
identity. Creation is refused if any UDF declaration already has that title,
including an incompatible declaration. The proposed Cost field identity is one
greater than all numeric ObjectId/TypeObjectId values/references in the document,
bounded by signed 32-bit range. Its exact identity is shown for review; no customer
ID is hardcoded. Creation does not populate milestones or target-only Activities.

The writer uses namespace-aware Expat element spans and applies only selected
CostValue edits, missing UDF insertions and an optional UDFType insertion. Unchanged
equal numeric values retain their original lexical bytes. A newly generated XML
is independently parsed and its selected field/values verified before writing.
NormalizedSchedule remains an import model, not the export representation.

## Output safety

Both original inputs are read-only. Fingerprints detect changes since assessment;
the engine reassesses before export and checks hashes again immediately before
publication. It uses the exact reviewed XML snapshot. There is no cross-process
lock on Excel/P6: another process can change an input after the final check, but
cannot change the snapshot used for this output.

Destination must be new, in an existing directory. Existing files, symlinks,
input aliases and concurrent destination creation are refused. A same-directory
temporary file is written, flushed/fsynced and byte-verified; atomic hard-link
publication fails instead of overwriting. Filesystems without hard-link support
fail safely; there is no unsafe overwrite/copy fallback. Cancellation or errors
before publication leave no destination and attempt temporary cleanup. A cleanup
failure after publication returns success with a cleanup warning. Power-loss
durability and arbitrary network filesystem guarantees are not established.

## Reproducible proof

Install the project's existing development dependencies, then run:

```bash
python -m pytest tests/unit/xml/test_xml_amount_writeback.py tests/integration/xml/test_xml_writeback_safety.py -q
python -m pytest -q
```

The fixture module creates deterministic P6 and raw OOXML inputs, including saved
formula caches. Tests cover review, precision, subsets, milestone/UDF conflicts,
creation, XML preservation, idempotence, source mutation, cancellation, collisions,
write/fsync/publication failures and the existing P6/Create import boundary.

For a supplied workbook/XML pair, the engineering harness first prints review JSON:

```bash
python scripts/proofs/xml_amount_writeback.py --workbook SOURCE.xlsx --target-xml TARGET.xml --field p6:udf:47 --monetary-confirmed --recalculated-and-saved
```

Inspect the assessment and its `Review digest`, then repeat the same command with
`--review-digest DIGEST --output NEW.xml`. For absent field creation replace
`--field ...` with `--create-field`. Any input/option change invalidates the review
digest. This harness is engineering tooling, not a production CLI/GUI feature.
Customer files are separate proof artifacts, not repository test fixtures.

## Acceptance evidence and remaining Desktop gates

XW-1 native P6 compatibility is accepted by the Product Owner; this handoff does
not request re-running the engine acceptance suite. Automated parsing/preservation
and Progress Studio re-import still do not prove every native P6 version/import
option. P6 can remap database identities and numeric precision.

For XW-2, Windows PO acceptance concerns the new dialog (including ENG/THA,
resizing and file pickers) and one saved NKC1 Workbook -> reviewed new XML -> P6
workflow. Use a copy/test project and check the selected field/representative
amounts, milestone preservation and usability. Record actual P6 import outcomes
separately; no new native Windows/P6 test is claimed by headless verification.

Focused Desktop verification:

```bash
python -m pytest tests/integration/desktop/test_xml_writeback_workflow.py tests/integration/desktop/test_product_experience.py tests/unit/xml/test_xml_amount_writeback.py tests/integration/xml/test_xml_writeback_safety.py -q
```

MSP/batch, multi-Project selection, XML merging, structural BOQ revision, XW-3 and
unrelated work remain outside XW-2. Coordinator review, PO acceptance and Git
integration require their separate gates and authority.
