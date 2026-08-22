# Schedule XML Requirements

Progress Studio accepts XML from Primavera P6, Microsoft Project, or another source when the schedule can be normalized to the required activity fields.

## Required for every activity

| Field | Rule |
|---|---|
| Activity Name | Must exist and must not be blank |
| Plan Start | Must be a valid date/time |
| Plan Finish | Must be a valid date/time and must not be earlier than Plan Start |

If any activity fails these rules, import stops and no workbook is created.

## Optional fields

| Field | Behavior when missing |
|---|---|
| Activity ID | A deterministic ID such as `ACT-000001` is generated |
| WBS / hierarchy | A flat project structure is created |
| Calendar | Not required for workbook generation |
| Relationships | Not required |
| Duration | Calculated from start and finish when needed |
| Actual dates | Optional |
| Percent complete | Optional |
| Resources and codes | Optional |

## Valid example

```xml
<Task>
  <Name>Excavation</Name>
  <Start>2026-03-01T08:00:00</Start>
  <Finish>2026-03-10T17:00:00</Finish>
</Task>
```

## Invalid examples

Missing activity name:

```xml
<Task>
  <Start>2026-03-01</Start>
  <Finish>2026-03-10</Finish>
</Task>
```

Finish earlier than start:

```xml
<Task>
  <Name>Excavation</Name>
  <Start>2026-03-10</Start>
  <Finish>2026-03-01</Finish>
</Task>
```

## Date guidance

ISO dates are safest:

```text
2026-03-01
2026-03-01T08:00:00
2026-03-01T08:00:00+07:00
```

Avoid ambiguous dates such as `03/04/2026` unless the source format has a defined locale.
