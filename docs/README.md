# Progress Studio Documentation

Active documentation is intentionally small. If two active documents disagree, `ARCHITECTURE.md` owns technical architecture and the implementation/tests own executable behavior.

## Start here

- [Repository README](../README.md) — what Progress Studio is, inputs, outputs and product flow.
- [User Workflow](USER_WORKFLOW.md) — Create, Mapping, Payment, Excel and Rebuild workflow.
- [Architecture](../ARCHITECTURE.md) — ownership boundaries and technical contracts.
- [Development](DEVELOPMENT.md) — developer setup and engineering rules.
- [Testing](TESTING.md) — current automated test tiers.
- [Roadmap](../ROADMAP.md) — pre-production milestones.
- [Release Checklist](../RELEASE_CHECKLIST.md) — RC/production gate.
- [Changelog](../CHANGELOG.md) — historical changes.

## Engineering reference

- `regressions/` — important regression investigations that remain useful for debugging.
- `history/` — milestone, acceptance, freeze and older user-guide records. These documents are **not** current product contracts.

## Historical user guides

The old v2.3 Thai/English manuals are archived at `history/user-guides-v2.3/`. They contain obsolete contracts such as generic Activity-ID/WBS fallback behavior and should not be used as the current operational guide.

- [WIN-1 Windows Portable Build](WIN1_WINDOWS_PORTABLE_BUILD.md) — Windows portable packaging contract and acceptance.
