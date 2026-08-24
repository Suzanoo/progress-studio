from __future__ import annotations

from pathlib import Path

import tomllib


ROOT = Path(__file__).resolve().parents[2]


def test_br12_brand_assets_are_packaged() -> None:
    data = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    package_data = data["tool"]["setuptools"]["package-data"]["progress_studio"]
    assert "assets/brand/*.png" in package_data
    assert "assets/brand/*.ico" in package_data


def test_br12_windows_build_uses_official_icon() -> None:
    spec = (ROOT / "packaging/windows/ProgressStudio.spec").read_text(encoding="utf-8")
    assert 'BRAND_ICON = REPO_ROOT / "progress_studio" / "assets" / "brand" / "progress_studio.ico"' in spec
    assert 'icon=str(BRAND_ICON)' in spec
    assert '"assets/brand/*.png"' in spec
    assert '"assets/brand/*.ico"' in spec


def test_br12_installer_uses_official_icon() -> None:
    iss = (ROOT / "packaging/windows/ProgressStudio.iss").read_text(encoding="utf-8")
    assert 'progress_studio\\assets\\brand\\progress_studio.ico' in iss
    assert "SetupIconFile={#MyAppIcon}" in iss


def test_br12_gui_sets_window_icon() -> None:
    app = (ROOT / "progress_studio/presentation/gui/app.py").read_text(encoding="utf-8")
    assert 'assets" / "brand" / "progress_studio.ico"' in app
    assert "self.iconbitmap" in app


def test_br12_portable_verifier_requires_brand_assets() -> None:
    verifier = (ROOT / "packaging/windows/verify_portable.py").read_text(encoding="utf-8")
    for name in ("progress_studio_brand.png", "progress_studio_icon.png", "progress_studio.ico"):
        assert name in verifier
