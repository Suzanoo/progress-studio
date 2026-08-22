# BOQ Workbook Requirements

The BOQ workbook may contain several worksheets. Progress Studio lets the user choose the worksheet manually.

## Expected BOQ columns

The current reader is designed for these fields:

```text
WBS-1
WBS-2
WBS-3
WBS-4
Description
Unit
Qty
Material
Labor
Amount
```

The most important mapping fields are **Description** and **Amount**. WBS columns are used for filtering and organization.

## Before loading the BOQ

1. Keep one BOQ table per worksheet.
2. Use one header row.
3. Remove merged cells inside the data table where possible.
4. Make sure Amount values are numeric.
5. Avoid hidden subtotal rows that repeat the same amount.
6. Keep descriptions meaningful enough for the user to identify the work item.
7. Save the file as `.xlsx` or `.xlsm`.

## Multiple worksheets

After selecting the BOQ file:

1. Review the **BOQ worksheet** list.
2. Choose the sheet containing the actual BOQ data.
3. Click **Load selected sheet**.

Progress Studio does not guess which worksheet is correct.

## Amount behavior

- `Amount` is the full BOQ item value.
- `Allocated` is the amount already assigned to activities.
- `Remaining %` shows the unallocated percentage.
- `Status` is normally Unmapped, Partial, or Full.
- One BOQ item can be split across several activities, but total allocation cannot exceed 100%.
