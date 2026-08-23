# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller one-folder build for WIN-1.

Build via ``scripts/build-windows-portable.ps1`` rather than invoking this spec
by hand. WIN-1 intentionally uses a one-folder distribution because it is
faster to start and easier to inspect when diagnosing missing resources.
"""

from PyInstaller.utils.hooks import collect_data_files


datas = collect_data_files(
    "progress_studio",
    includes=[
        "config/*.json",
        "assets/dashboard/icons/*.png",
        "services/distribution/*.json",
    ],
)

a = Analysis(
    ["packaging/windows/launcher.py"],
    pathex=["."],
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
