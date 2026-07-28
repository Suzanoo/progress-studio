# MS-4 Acceptance Criteria — BOQ and Amount

MS-4 passes only when all criteria below are satisfied:

1. `AmountStep` is part of the application pipeline.
2. Amount source decisions and results are represented by domain models.
3. Amount Mapping creation is owned by `AmountService` and Excel infrastructure.
4. Amount application and WBS/project rollups are owned by `AmountService` and Excel infrastructure.
5. `PipelineContext` exposes `amount_workbook`.
6. The legacy continuation starts at stage 05.
7. The legacy continuation does not require or execute `04_apply_boq.py`.
8. V2 Amount Mapping and stage-04 workbook match the V1 baseline.
9. No Thai text exists in application source.
10. All automated tests pass.
