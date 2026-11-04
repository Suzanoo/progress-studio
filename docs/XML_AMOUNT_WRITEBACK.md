# P6 XML Amount writeback engine — XW-1

Status: engineering implementation handoff. Authorized 2026-11-04 against
`main @ de125d4310cabce504e236e3e090192bf2418aa2` / `bu1-boq-amount-update-stable`.
X1–X5 are accepted in [the initiative plan](BOQ_UPDATE_XML_WRITEBACK_PLAN.md).
No Desktop integration, PO acceptance, native P6 acceptance or release is claimed.

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

## Remaining native compatibility gates

Automated XML parsing, exact byte preservation and Progress Studio re-import do
not establish Primavera import compatibility. Before claiming native round-trip
support, import copies into an isolated P6 test project for (1) an existing Cost
UDF and (2) a newly created Cost UDF. Check P6's actual field identity/type/currency
and selected Activity values; export again and compare milestone amounts,
target-only Activities and schedule/resource/relationship data. Record P6 version,
import options and any diagnostics. Native P6 can remap database ObjectIds and
numeric precision; this engine does not promise preservation across P6's importer.

XW-2 Desktop integration, MSP/batch, multi-Project selection and structural BOQ
revision remain outside this implementation. Coordinator review and PO acceptance
are separate from engineering proof and Git integration requires its own authority.
