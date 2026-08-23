# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller one-folder build for WIN-1.

Build via ``scripts/build-windows-portable.ps1`` rather than invoking this spec
by hand. WIN-1 intentionally uses a one-folder distribution because it is
faster to start and easier to inspect when diagnosing missing resources.
"""

from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files


# PyInstaller executes spec files with ``SPECPATH`` set to the directory that
# contains this file. Resolve every local path from that anchor so the build is
# independent of the caller's current working directory.
SPEC_DIR = Path(SPECPATH).resolve()
REPO_ROOT = SPEC_DIR.parents[1]
LAUNCHER = SPEC_DIR / "launcher.py"

if not LAUNCHER.is_file():
    raise FileNotFoundError(f"WIN-1 launcher not found: {LAUNCHER}")


datas = collect_data_files(
    "progress_studio",
    includes=[
        "config/*.json",
        "assets/dashboard/icons/*.png",
        "services/distribution/*.json",
    ],
)

a = Analysis(
    [str(LAUNCHER)],
    pathex=[str(REPO_ROOT)],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pytest"],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="ProgressStudio",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="ProgressStudio",
)
