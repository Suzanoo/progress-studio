# Troubleshooting

| Message or symptom | Cause | What to do |
|---|---|---|
| No activities were found | XML does not contain recognizable activity/task nodes | Export the schedule again or use a supported XML structure |
| Missing Activity Name | At least one activity has no name | Add a task/activity name in the source schedule and export XML again |
| Missing or invalid Plan Start | Start is blank or unreadable | Correct the planned start date and export again |
| Missing or invalid Plan Finish | Finish is blank or unreadable | Correct the planned finish date and export again |
| Finish is earlier than Start | Schedule dates are inconsistent | Correct the activity dates in the source application |
| BOQ worksheet list is empty | Workbook cannot be read or has no usable worksheet | Open and resave the BOQ as `.xlsx`, then load again |
| BOQ data looks wrong | Wrong worksheet was selected | Choose the correct BOQ worksheet and reload it |
| Map button does nothing | No Activity or BOQ item is selected | Tick one Activity and at least one BOQ item |
| Share exceeds available amount | Combined allocation would be above 100% | Reduce the Share or unmap an earlier allocation |
| Nothing to undo | No reversible action remains | Select rows and use Unmap when appropriate |
| Session workbook cannot be verified | Workbook was moved, renamed, or modified | Browse to the original identical workbook; modified files are rejected |
| Required worksheet `main` was not found | Progress workbook contract was changed | Use a Progress Studio-generated workbook and do not rename `main` |
| Excel percentages do not update immediately | Formula results are cached until Excel recalculates | Open the export in Microsoft Excel, wait, and save |
| Excel repairs a workbook | An old or externally modified workbook may contain invalid table metadata | Use the current 2.3.0 export and avoid editing Excel table structures manually |
| Desktop window appears busy | A large XML or BOQ is being processed | Wait for completion; do not start the same operation twice |

## Before reporting a problem

Record these details:

```text
Progress Studio version
Windows version
Source application (P6, MS Project, other)
Exact error message
Number of activities
BOQ worksheet name
Steps performed before the error
```

Do not share confidential project files unless you are authorized to do so.
