# Progress Studio User Workflow

This is the current high-level workflow. Historical v2.3 guides are archived under `docs/history/user-guides-v2.3/` and should not be treated as the current contract.

## 1. Create Progress

Input:

- Microsoft Project XML, or
- Primavera P6 XML.

Progress Studio normalizes the schedule and creates the initial Excel workbook.

The initial workbook includes weekly `main`, monthly `main_monthly`, Dashboard and internal/helper data required by the selected workflow.

Initial schedule Amount uses Progress Studio's fallback/fake amount behavior. Real BOQ cost ownership belongs to Mapping.

## 2. Mapping (optional)

Load the Progress workbook and BOQ workbook, then allocate BOQ amounts to Activities.

Activity identity is Activity ID. Mapping owns BOQ -> Activity allocation; it does not own Payment or Progress rebuild behavior.

## 3. Payment (optional)

Payment Input stores persistent payment requirements. Payment rendering uses the current workbook schedule/progress structure to position Payment output.

Payment must not rebuild Progress-owned Dashboard/overlay data.

## 4. Work in Excel

After initial generation, `main` is the workbook source of truth.

Typical edits include Actual progress and other approved editable inputs. Structural edits such as adding/deleting Activities, WBS or timescale periods require Rebuild to regenerate Python-owned derived outputs.

### Timescale labels

```text
X  X | W1 W2 ... Wn | X X
X    | M1 M2 ... Mn | X
```

- `X` = display margin.
- `Wn` / `Mn` = reporting-period labels.

## 5. Recalculation in Excel

Progress Studio workbooks use a user-driven Excel formula calculation policy.

- **F9** recalculates Excel formulas.
- **Save** asks Excel to calculate formulas on save.
- Neither action runs Python code.

If a generated snapshot/cache/view must be regenerated after structural changes, use Progress Studio Rebuild.

## 6. Rebuild

Rebuild uses the selected workbook itself as the source.

The workspace exposes Snapshot/Live workbook modes and Progress/Payment scope.

```text
                    Progress    Payment
Snapshot               ✓           ✓
Live                   ✓           ✓
```

### Progress Rebuild

- preserves `main` as the edited source;
- rebuilds Progress-owned derived outputs for the selected mode;
- preserves Payment/user-owned data according to the contract;
- does not renumber the weekly W/X labels in `main`.

### Payment Rebuild

- rebuilds Payment only from current `main + Payment Input`;
- preserves Progress-generated views.

## When to Rebuild

Use Rebuild when the workbook structure or Python-owned generated data needs regeneration, for example after:

- adding/deleting Activities;
- changing WBS structure;
- extending/shrinking the project timescale;
- major Plan/rebaseline changes;
- changing Payment Input that requires Payment to be regenerated.
