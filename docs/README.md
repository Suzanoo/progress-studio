# Progress Studio Documentation

Active documentation is intentionally small. If two active documents
disagree, `ARCHITECTURE.md` owns technical architecture and the
implementation/tests own executable behavior.

## Collaboration roles

- [Coordinator Contract](COORDINATOR_CONTRACT.md) — product discussion, prompts,
  review, acceptance briefings and current agent assignments.
- [Engineer Contract](ENGINEER_CONTRACT.md) — engineering scope, ownership,
  acceptance, handoff and authorized Git execution boundaries.
- [Engineer Workflow](ENGINEER_WORKFLOW.md) — targeted investigation,
  implementation, verification and delivery procedure.

These role-based documents replace `CHATGPT_CONTRACT.md`, `WORK_CONTRACT.md`
and `WORK_SKILL.md`, respectively. Historical records may use the previous names
and agent assignments.

## Start here

-   [Repository README](../README.md) --- what Progress Studio is,
    inputs, outputs and product flow.
-   [User Workflow](USER_WORKFLOW.md) --- Create, Mapping, Payment,
    Excel and Rebuild workflow.
-   [Weight Selection](WEIGHT_SELECTION.md) --- Equal/Duration/Amount creation and monetary versus dummy-weight semantics.
-   [Payment Breakdown](PAYMENT_BREAKDOWN.md) --- exact-name grouping,
    calculation, workbook ownership and output contract.
-   [Architecture](../ARCHITECTURE.md) --- ownership boundaries and
    technical contracts.
-   [Development](DEVELOPMENT.md) --- developer setup and engineering
    rules.
-   [Testing](TESTING.md) --- current automated test tiers.
-   [Roadmap](../ROADMAP.md) --- pre-production milestones.
-   [Release Checklist](../RELEASE_CHECKLIST.md) --- RC/production gate.
-   [Changelog](../CHANGELOG.md) --- historical changes.

## Engineering reference

-   `regressions/` --- important regression investigations that remain
    useful for debugging.
-   `history/` --- milestone, acceptance, freeze and older user-guide
    records. These documents are **not** current product contracts.

## Historical user guides

The old v2.3 Thai/English manuals are archived at
`history/user-guides-v2.3/`. They contain obsolete contracts such as
generic Activity-ID/WBS fallback behavior and should not be used as the
current operational guide.

-   [WIN-1 Windows Portable Build](WIN1_WINDOWS_PORTABLE_BUILD.md) ---
    Windows portable packaging contract and acceptance.

- [Financial Forecast calculation contract](FINANCIAL_FORECAST_CONTRACT.md) — FF-1 source-neutral cash model, timing, reconciliation, defaults and limits.
- [Finance workbook](FF2_WORKBOOK.md) — FF-2 persistent inputs, live formulas, commands and Desktop Excel acceptance.
- [Financial Forecast desktop](FF3_DESKTOP.md) — normal application create/refresh workflow and FF-3 Product Owner acceptance procedure.
