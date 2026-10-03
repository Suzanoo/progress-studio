# Progress Studio documentation

## Current authority

- [Product README](../README.md) — product, installation and workflow.
- [User workflow](USER_WORKFLOW.md) — Create, Excel updates, optional Mapping/Payment and Rebuild.
- [Architecture](ARCHITECTURE.md) — technical ownership; executable behavior is protected by tests.
- [V1.1 product experience](V1_1_PRODUCT_EXPERIENCE.md) — approved scope, UI/localization and acceptance.
- [Weight selection](WEIGHT_SELECTION.md)
- [EV monetary sources](EV_MONETARY_SOURCES.md) — Activity Amount / BOQ BAC and milestone semantics.
- [EV live contract](EV_LIVE_CONTRACT.md) — live progress/date/refresh boundaries.
- [Payment Breakdown](PAYMENT_BREAKDOWN.md)
- [Roadmap/status](ROADMAP.md) — V1 stable, MS-3 cancelled, MS-4 completed, V1.1 pending PO acceptance.
- [Changelog](../CHANGELOG.md)

## Collaboration authority

- [Coordinator Contract](COORDINATOR_CONTRACT.md)
- [Engineer Contract](ENGINEER_CONTRACT.md)
- [Engineer Workflow](ENGINEER_WORKFLOW.md)

These replace CHATGPT_CONTRACT.md, WORK_CONTRACT.md and WORK_SKILL.md.
Historical agent assignments do not override the current role contracts.

## Engineering and distribution

- [Development](DEVELOPMENT.md)
- [Testing](TESTING.md)
- [Release checklist](RELEASE_CHECKLIST.md)
- [Windows build guide](../packaging/windows/WINDOWS_BUILD_GUIDE.md)
- [Windows portable contract](WIN1_WINDOWS_PORTABLE_BUILD.md)

## Retained Finance reference — hidden from normal V1.1 navigation

Underlying accepted Finance implementation and tests remain supported; no calculations
were redesigned for V1.1.

- [Financial Forecast calculation contract](FINANCIAL_FORECAST_CONTRACT.md)
- [Finance workbook](FF2_WORKBOOK.md)
- [Finance desktop reference](FF3_DESKTOP.md)

## History, not current product authority

- [Pre-production roadmap](history/PRE_PRODUCTION_ROADMAP.md)
- [EV-1 derivation notes](history/EV1_NOTES.md)
- `history/` contains milestone/freeze/acceptance records and obsolete v2.3 user guides.
- `regressions/` contains useful investigations; current contracts and tests take precedence.

Old roadmap schedules and BOQ-only EV assumptions are historical, not new work.
