from __future__ import annotations

import os
from pathlib import Path

SETTINGS_ORG_NAME = "ThreepSoftwz"
SETTINGS_APP_NAME = "arr_helper"


def _resolve_base_dir(env_var: str, fallback: Path) -> Path:
    override = os.getenv(env_var, "").strip()
    if override:
        return Path(override).expanduser().resolve()
    return fallback.expanduser().resolve()


def get_config_dir() -> Path:
    if os.name == "nt":
        fallback = Path(os.getenv("APPDATA", str(Path.home() / "AppData" / "Roaming")))
    else:
        fallback = Path.home() / ".config"
    return _resolve_base_dir("CONFIG_DIR", fallback) / SETTINGS_ORG_NAME


def get_data_dir() -> Path:
    if os.name == "nt":
        fallback = Path(os.getenv("LOCALAPPDATA", str(Path.home() / "AppData" / "Local")))
    else:
        fallback = Path.home() / ".local" / "share"
    return _resolve_base_dir("DATA_DIR", fallback) / SETTINGS_ORG_NAME / SETTINGS_APP_NAME


def get_log_dir() -> Path:
    return get_data_dir() / "logs"


def get_settings_ini_path() -> Path:
    return get_config_dir() / f"{SETTINGS_APP_NAME}.ini"
