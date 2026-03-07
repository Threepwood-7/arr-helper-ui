from __future__ import annotations

import os
from pathlib import Path

from PySide6.QtCore import QSettings

SETTINGS_ORG_NAME = "ThreepSoftwz"
SETTINGS_APP_NAME = "arr_helper"


def resolve_config_root(override_dir: str | None = None) -> Path:
    override = str(override_dir or os.getenv("CONFIG_DIR", "")).strip()
    if override:
        root = Path(override).expanduser().resolve()
    elif os.name == "nt":
        fallback = Path(os.getenv("APPDATA", str(Path.home() / "AppData" / "Roaming")))
        root = fallback.expanduser().resolve()
    else:
        root = (Path.home() / ".config").expanduser().resolve()
    root.mkdir(parents=True, exist_ok=True)
    return root


def resolve_data_root(override_dir: str | None = None) -> Path:
    override = str(override_dir or os.getenv("DATA_DIR", "")).strip()
    if override:
        root = Path(override).expanduser().resolve()
    elif os.name == "nt":
        fallback = Path(os.getenv("LOCALAPPDATA", str(Path.home() / "AppData" / "Local")))
        root = fallback.expanduser().resolve() / SETTINGS_ORG_NAME
    else:
        root = ((Path.home() / ".local" / "share").expanduser().resolve()) / SETTINGS_ORG_NAME
    root.mkdir(parents=True, exist_ok=True)
    return root


def resolve_app_data_dir(override_dir: str | None = None) -> Path:
    app_dir = resolve_data_root(override_dir) / SETTINGS_APP_NAME
    app_dir.mkdir(parents=True, exist_ok=True)
    return app_dir


def configure_qsettings(config_dir_override: str | None = None) -> Path:
    config_root = resolve_config_root(config_dir_override)
    QSettings.setDefaultFormat(QSettings.Format.IniFormat)
    QSettings.setPath(
        QSettings.Format.IniFormat,
        QSettings.Scope.UserScope,
        str(config_root),
    )
    return config_root


def get_config_dir() -> Path:
    config_dir = resolve_config_root() / SETTINGS_ORG_NAME
    config_dir.mkdir(parents=True, exist_ok=True)
    return config_dir


def get_data_dir() -> Path:
    return resolve_app_data_dir()


def get_log_dir() -> Path:
    return get_data_dir() / "logs"


def get_settings_ini_path() -> Path:
    return get_config_dir() / f"{SETTINGS_APP_NAME}.ini"
