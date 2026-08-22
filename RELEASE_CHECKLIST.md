# Progress Studio Release Checklist

This checklist is for Release Candidate / production builds. It intentionally does not carry an old fixed version number.

## Source / Git

- [ ] Working tree is clean.
- [ ] Intended release commit is tagged.
- [ ] `CHANGELOG.md` and `ROADMAP.md` are current.
- [ ] No `.venv`, pytest cache, build output, temporary workbooks or debug ZIPs are included in release source archives.

## Automated gates

- [ ] Smoke tests pass.
- [ ] Relevant regression tests pass.
- [ ] Rebuild 2 x 2 matrix passes.
- [ ] Normalizer MSP/P6 equivalence gate passes.
- [ ] Editable-workbook mutation suite passes.
- [ ] Full `release` test gate passes.

## Excel compatibility

- [ ] Representative Create Progress workbook opens in Microsoft Excel without repair/recovery dialog.
- [ ] Weekly and Monthly Plan/Actual overlays render correctly.
- [ ] Dashboard reporting range excludes display margins.
- [ ] F9 / Save formula recalculation behaves as documented.
- [ ] Protection/visibility policy is correct.
- [ ] Payment-only workflow preserves Progress-owned charts/data.

## Packaging

### Windows

- [ ] Build/installer produced from a clean environment.
- [ ] Tested on clean supported Windows machine.
- [ ] Launch, file dialogs, workbook generation and logs verified.

### macOS

- [ ] App bundle produced for the supported architecture(s).
- [ ] Signing/notarization policy completed as required.
- [ ] Tested on clean supported macOS machine.

## Manual acceptance

- [ ] Small synthetic schedule.
- [ ] Representative real MSP project.
- [ ] Representative P6 project.
- [ ] Create -> Mapping -> Payment -> Excel edit -> Rebuild workflow.
