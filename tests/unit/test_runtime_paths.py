from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QSettings
from threep_commons.paths import (
    configure_qsettings,
    resolve_app_data_dir,
    resolve_config_root,
    resolve_data_root,
)

from arr_helper.constants import APP_IDENTITY, SETTINGS_APP_NAME, SETTINGS_ORG_NAME


def test_resolve_config_root_uses_config_dir_override(
    monkeypatch, tmp_path: Path
) -> None:
    target = tmp_path / "cfg"
    monkeypatch.setenv("CONFIG_DIR", str(target))

    root = resolve_config_root()

    assert root == target
    assert root.exists()


def test_resolve_data_root_uses_data_dir_override(monkeypatch, tmp_path: Path) -> None:
    target = tmp_path / "data"
    monkeypatch.setenv("DATA_DIR", str(target))

    root = resolve_data_root(APP_IDENTITY)

    assert root == target
    assert root.exists()


def test_resolve_app_data_dir_appends_app_name(monkeypatch, tmp_path: Path) -> None:
    target = tmp_path / "data"
    monkeypatch.setenv("DATA_DIR", str(target))

    app_dir = resolve_app_data_dir(APP_IDENTITY)

    assert app_dir == target / SETTINGS_APP_NAME
    assert app_dir.exists()


def test_configure_qsettings_sets_ini_user_path(monkeypatch, tmp_path: Path) -> None:
    config_dir = tmp_path / "cfg"
    monkeypatch.setenv("CONFIG_DIR", str(config_dir))

    configure_qsettings(APP_IDENTITY)
    settings = QSettings(
        QSettings.Format.IniFormat,
        QSettings.Scope.UserScope,
        SETTINGS_ORG_NAME,
        SETTINGS_APP_NAME,
    )
    settings.sync()
    settings_path = Path(settings.fileName())

    assert settings_path.parts[-2:] == (SETTINGS_ORG_NAME, f"{SETTINGS_APP_NAME}.ini")
    assert settings_path.parent.parent == config_dir
