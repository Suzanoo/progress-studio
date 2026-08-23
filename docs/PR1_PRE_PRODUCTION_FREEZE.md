# PR-1 — Pre-Production Freeze

PR-1 freezes the current Progress Studio product behavior before platform-specific Windows and macOS builds begin.

## Scope

PR-1 is a verification and release-engineering milestone. It does **not** add product features and does **not** refactor the workbook engines.

Frozen product contracts include:

- MSP and P6 XML normalization into the shared normalized schedule model.
- Create Progress generation of Weekly/Monthly/Dashboard outputs.
- `X / Wn / Mn` reporting-period semantics.
- reporting range separated from display margins.
- Mapping ownership of BOQ-to-Activity amount allocation.
- Payment ownership of Payment outputs only.
- Rebuild Snapshot/Live Progress and Payment 2 x 2 ownership boundaries.
- workbook protection, visibility, Excel recalculation and chart/overlay contracts covered by regression tests.
- packaged runtime JSON/icon resources verified from a built wheel.

## Commercial readiness boundary

The application is prepared so licensing/entitlement can be added later at the application boundary. PR-1 intentionally does **not** implement:

- license keys or activation servers;
- machine binding;
- user accounts/subscriptions;
- payment gateways;
- feature locking.

Licensing must remain outside the Progress, Mapping, Payment and Rebuild engines.

## Release gate

Run:

```powershell
.\scripts\check-preproduction.ps1
```

The gate executes all current test directories (unit, regression, integration and each acceptance module), the smoke subset and a built-wheel package-resource check.

The acceptance modules are executed separately because some legacy end-to-end tests exhibit cumulative runtime slowdown when the entire suite shares one pytest process. This preserves full test coverage without changing product behavior.

## Exit criteria

PR-1 is complete when:

- Git working tree is clean;
- all unit, regression, integration and acceptance tests pass (skips must be understood);
- smoke passes;
- package artifact verification passes;
- no production source behavior is changed by the PR-1 milestone itself;
- documentation and release checklist reflect the pre-production state.

Platform-specific installer/signing work begins only after this gate is green.

## Windows release handoff

WIN-1 builds the one-folder Windows portable application. WIN-2 validates that
same artifact outside the repository and then on a clean Windows user/VM before
installer work begins. Neither milestone changes frozen Progress Studio core
behavior.
