# WIN-3 Acceptance Record — 23-Aug-2026

## Build identity

- Product: Progress Studio
- Version: 2.3.0
- Windows build environment: Windows 10 build 19045
- Packaging: PyInstaller one-folder payload + Inno Setup 6 installer

## Known-good portable payload

```text
SHA-256:
96F5D56348A8692659E3FFA0ED7238CE0E1632E4DB8E67AB25B6B09572A211A0
```

WIN-2 automated isolation probe: **PASS**

## Installer

```text
ProgressStudio-Setup-2.3.0.exe

SHA-256:
76728AF64EB4D772351B4985D7AD612E9689C1474479EF54824714025EAEC73A
```

Inno Setup compile: **PASS**

Windows Installed Apps registration: **PASS**

The installed entry was visible as **Progress Studio version 2.3.0** with an
Uninstall action.

## Scope boundary

This WIN-3 checkpoint does not add:

- code signing;
- auto-update;
- licensing / activation;
- account or payment services;
- commercial feature gating.

The installer wraps the validated portable application and does not alter
Progress Studio workbook or calculation behavior.
