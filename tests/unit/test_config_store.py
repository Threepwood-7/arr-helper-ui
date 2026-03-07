from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QSettings
from threep_commons.paths import configure_qsettings

from arr_helper.constants import APP_IDENTITY, SETTINGS_APP_NAME, SETTINGS_ORG_NAME
from arr_helper.core.paths import get_app_cache_dir
from arr_helper.media_checker.config import Config


def test_config_save_uses_config_namespace(monkeypatch, tmp_path: Path) -> None:
    config_dir = tmp_path / "cfg"
    monkeypatch.setenv("CONFIG_DIR", str(config_dir))
    configure_qsettings(APP_IDENTITY, config_dir_override=str(config_dir))

    cfg = Config()
    cfg.save()

    settings = QSettings(
        QSettings.Format.IniFormat,
        QSettings.Scope.UserScope,
        SETTINGS_ORG_NAME,
        SETTINGS_APP_NAME,
    )
    keys = settings.allKeys()

    assert keys
    assert all(key.startswith("config/") for key in keys)


def test_get_app_cache_dir_uses_data_dir_override(monkeypatch, tmp_path: Path) -> None:
    data_dir = tmp_path / "data"
    monkeypatch.setenv("DATA_DIR", str(data_dir))

    cache_dir = Path(get_app_cache_dir(str(tmp_path)))

    assert cache_dir == data_dir / SETTINGS_APP_NAME / "cache"
    assert cache_dir.exists()
