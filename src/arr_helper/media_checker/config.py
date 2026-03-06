"""Configuration and cache persistence for media quality checker."""

from __future__ import annotations

import copy
import json
import os
from pathlib import Path
from typing import Any

from PySide6.QtCore import QSettings

from ..core.locking import (
    acquire_lock_file as _core_acquire_lock_file,
)
from ..core.locking import (
    pid_is_running as _core_pid_is_running,
)
from ..core.locking import (
    release_lock_file as _core_release_lock_file,
)
from ..core.locking import (
    write_json_atomic_locked as _core_write_json_atomic_locked,
)
from ..core.paths import get_app_cache_dir
from ..runtime_paths import (
    SETTINGS_APP_NAME,
    SETTINGS_ORG_NAME,
    configure_qsettings,
)

APP_SLUG = "arr_helper"

SECRET_ENV_TO_KEYS = (
    ("ARR_HELPER_SONARR_API_KEY", ("sonarr", "api_key")),
    ("ARR_HELPER_RADARR_API_KEY", ("radarr", "api_key")),
    (
        "ARR_HELPER_SONARR_HTTP_BASIC_AUTH_PASSWORD",
        ("sonarr", "http_basic_auth_password"),
    ),
    (
        "ARR_HELPER_RADARR_HTTP_BASIC_AUTH_PASSWORD",
        ("radarr", "http_basic_auth_password"),
    ),
)

PLACEHOLDER_API_KEYS = {
    "your_sonarr_api_key_here",
    "your_radarr_api_key_here",
    "your-sonarr-api-key-here",
    "your-radarr-api-key-here",
    "",
}

DEFAULT_CONFIG: dict[str, Any] = {
    "sonarr": {
        "url": "http://localhost:8989",
        "api_key": "your-sonarr-api-key-here",
        "enabled": True,
        "http_basic_auth_username": "",
        "http_basic_auth_password": "",
    },
    "radarr": {
        "url": "http://localhost:7878",
        "api_key": "your-radarr-api-key-here",
        "enabled": True,
        "http_basic_auth_username": "",
        "http_basic_auth_password": "",
    },
    "settings": {
        "dry_run": False,
        "interactive": True,
        "require_english_audio": True,
        "require_english_subs": True,
        "english_language_codes": ["eng", "en", "english"],
        "highlight_missing_subs": "",
    },
}

# (path, value_type, default)
CONFIG_SCHEMA: tuple[tuple[str, type, Any], ...] = (
    ("config/sonarr/url", str, DEFAULT_CONFIG["sonarr"]["url"]),
    ("config/sonarr/api_key", str, DEFAULT_CONFIG["sonarr"]["api_key"]),
    ("config/sonarr/enabled", bool, DEFAULT_CONFIG["sonarr"]["enabled"]),
    (
        "config/sonarr/http_basic_auth_username",
        str,
        DEFAULT_CONFIG["sonarr"]["http_basic_auth_username"],
    ),
    (
        "config/sonarr/http_basic_auth_password",
        str,
        DEFAULT_CONFIG["sonarr"]["http_basic_auth_password"],
    ),
    ("config/radarr/url", str, DEFAULT_CONFIG["radarr"]["url"]),
    ("config/radarr/api_key", str, DEFAULT_CONFIG["radarr"]["api_key"]),
    ("config/radarr/enabled", bool, DEFAULT_CONFIG["radarr"]["enabled"]),
    (
        "config/radarr/http_basic_auth_username",
        str,
        DEFAULT_CONFIG["radarr"]["http_basic_auth_username"],
    ),
    (
        "config/radarr/http_basic_auth_password",
        str,
        DEFAULT_CONFIG["radarr"]["http_basic_auth_password"],
    ),
    ("config/settings/dry_run", bool, DEFAULT_CONFIG["settings"]["dry_run"]),
    ("config/settings/interactive", bool, DEFAULT_CONFIG["settings"]["interactive"]),
    (
        "config/settings/require_english_audio",
        bool,
        DEFAULT_CONFIG["settings"]["require_english_audio"],
    ),
    (
        "config/settings/require_english_subs",
        bool,
        DEFAULT_CONFIG["settings"]["require_english_subs"],
    ),
    (
        "config/settings/english_language_codes",
        list,
        DEFAULT_CONFIG["settings"]["english_language_codes"],
    ),
    (
        "config/settings/highlight_missing_subs",
        str,
        DEFAULT_CONFIG["settings"]["highlight_missing_subs"],
    ),
)


def _pid_is_running(pid: int) -> bool:
    return _core_pid_is_running(pid)


def _acquire_lock_file(lock_path: str, timeout_s: float = 10.0) -> tuple[int, str]:
    return _core_acquire_lock_file(lock_path, timeout_s=timeout_s, pid_checker=_pid_is_running)


def _release_lock_file(lock_path: str, lock_fd: int, token: str) -> None:
    _core_release_lock_file(lock_path, lock_fd, token)


def _write_json_atomic_locked(path: str, payload: dict, indent: int = 2) -> None:
    _core_write_json_atomic_locked(path, payload, indent=indent)


def _deep_merge_dicts(base: dict[str, Any], overlay: dict[str, Any]) -> dict[str, Any]:
    merged: dict[str, Any] = dict(base)
    for key, value in overlay.items():
        base_value = merged.get(key)
        if isinstance(base_value, dict) and isinstance(value, dict):
            merged[key] = _deep_merge_dicts(base_value, value)
        else:
            merged[key] = value
    return merged


def _set_nested_value(target: dict[str, Any], key_path: tuple[str, ...], value: Any) -> None:
    current: dict[str, Any] = target
    for key in key_path[:-1]:
        next_value = current.get(key)
        if not isinstance(next_value, dict):
            next_value = {}
            current[key] = next_value
        current = next_value
    current[key_path[-1]] = value


def _schema_config_path(schema_key: str) -> tuple[str, ...]:
    parts = tuple(schema_key.split("/"))
    if parts and parts[0] == "config":
        return parts[1:]
    return parts


def _new_config_settings() -> QSettings:
    configure_qsettings()
    return QSettings(
        QSettings.Format.IniFormat,
        QSettings.Scope.UserScope,
        SETTINGS_ORG_NAME,
        SETTINGS_APP_NAME,
    )


def _coerce_bool(value: Any, default: bool) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    if isinstance(value, str):
        token = value.strip().lower()
        if token in {"1", "true", "yes", "on"}:
            return True
        if token in {"0", "false", "no", "off"}:
            return False
    return bool(default)


def _coerce_str_list(value: Any, default: list[str]) -> list[str]:
    if isinstance(value, list):
        return [str(v).strip() for v in value if str(v).strip()]
    if isinstance(value, tuple):
        return [str(v).strip() for v in value if str(v).strip()]
    if isinstance(value, str):
        tokens = [item.strip() for item in value.split(",")]
        return [item for item in tokens if item]
    return list(default)


def _coerce_value(value: Any, value_type: type, default: Any) -> Any:
    if value_type is bool:
        return _coerce_bool(value, bool(default))
    if value_type is int:
        try:
            return int(value)
        except (TypeError, ValueError, OverflowError):
            return int(default)
    if value_type is float:
        try:
            return float(value)
        except (TypeError, ValueError, OverflowError):
            return float(default)
    if value_type is list:
        return _coerce_str_list(value, list(default))
    if value is None:
        return default
    return str(value)


def _settings_file_path(settings: QSettings) -> str:
    settings.sync()
    file_name = str(settings.fileName() or "").strip()
    if file_name:
        return file_name
    return str(Path.cwd() / f"{SETTINGS_APP_NAME}.ini")


class Config:
    """Load and manage configuration from QSettings."""

    def __init__(self) -> None:
        self._settings = _new_config_settings()
        self.config_path = _settings_file_path(self._settings)
        self.config = self._load_config()

        cache_dir = get_app_cache_dir(str(Path.cwd()))
        self.user_cache_path = os.path.join(cache_dir, "z_user.cache")
        self.files_cache_path = os.path.join(cache_dir, "z_files.cache")

    def _load_config(self) -> dict[str, Any]:
        merged = copy.deepcopy(DEFAULT_CONFIG)
        for key, value_type, default in CONFIG_SCHEMA:
            raw = self._settings.value(key, default)
            value = _coerce_value(raw, value_type, default)
            _set_nested_value(merged, _schema_config_path(key), value)
        self._apply_secret_env_overrides(merged)
        return merged

    def reload(self) -> dict[str, Any]:
        self.config = self._load_config()
        return self.config

    def _apply_secret_env_overrides(self, config: dict[str, Any]) -> None:
        for env_name, key_path in SECRET_ENV_TO_KEYS:
            env_value = os.environ.get(env_name, "")
            if env_value:
                _set_nested_value(config, key_path, env_value)

    def save(self, new_config: dict[str, Any] | None = None) -> None:
        if new_config is not None:
            self.config = _deep_merge_dicts(copy.deepcopy(DEFAULT_CONFIG), new_config)
        for key, value_type, default in CONFIG_SCHEMA:
            path = _schema_config_path(key)
            current: Any = self.config
            for part in path:
                if not isinstance(current, dict):
                    current = default
                    break
                current = current.get(part, default)
            value = _coerce_value(current, value_type, default)
            self._settings.setValue(key, value)
        self._settings.sync()
        self.config_path = _settings_file_path(self._settings)
        self.reload()

    def get_missing_required(self, context: str = "media_checker") -> list[str]:
        errors: list[str] = []
        sonarr = self.config.get("sonarr", {}) if isinstance(self.config, dict) else {}
        radarr = self.config.get("radarr", {}) if isinstance(self.config, dict) else {}

        sonarr_enabled = bool(sonarr.get("enabled", True)) if isinstance(sonarr, dict) else False
        radarr_enabled = bool(radarr.get("enabled", True)) if isinstance(radarr, dict) else False

        if context == "sonarr_ui":
            if not sonarr_enabled:
                errors.append("[sonarr] must be enabled for Sonarr UI")
                return errors
            targets = [("sonarr", sonarr, True)]
        else:
            if not sonarr_enabled and not radarr_enabled:
                errors.append("At least one of [sonarr] or [radarr] must be enabled")
            targets = [
                ("sonarr", sonarr, sonarr_enabled),
                ("radarr", radarr, radarr_enabled),
            ]

        for name, section, enabled in targets:
            if not enabled:
                continue
            if not isinstance(section, dict):
                errors.append(f"[{name}] is not configured")
                continue
            url = str(section.get("url", "") or "").strip()
            api_key = str(section.get("api_key", "") or "").strip()
            if not url:
                errors.append(f"[{name}] 'url' is missing or empty")
            if api_key.lower() in PLACEHOLDER_API_KEYS:
                errors.append(f"[{name}] 'api_key' is missing or placeholder")
        return errors

    def validate(self, context: str = "media_checker") -> list[str]:
        return self.get_missing_required(context=context)

    def validation_help(self, context: str = "media_checker") -> str:
        details = "\n".join(f"  - {item}" for item in self.get_missing_required(context=context))
        return (
            "Configuration validation failed:\n"
            f"{details}\n\n"
            f"Open and edit QSettings INI file: {self.config_path}"
        )

    def load_user_cache(self) -> dict[str, Any]:
        if os.path.exists(self.user_cache_path):
            try:
                with open(self.user_cache_path, encoding="utf-8") as f:
                    loaded = json.load(f)
                return loaded if isinstance(loaded, dict) else {}
            except Exception as e:
                print(f"Warning: Could not load user cache: {e}")
                return {}
        return {}

    def save_user_cache(self, cache: dict[str, Any]) -> None:
        try:
            _write_json_atomic_locked(self.user_cache_path, cache, indent=2)
        except Exception as e:
            print(f"Warning: Could not save user cache: {e}")

    def load_files_cache(self) -> dict[str, Any]:
        if os.path.exists(self.files_cache_path):
            try:
                with open(self.files_cache_path, encoding="utf-8") as f:
                    loaded = json.load(f)
                return loaded if isinstance(loaded, dict) else {}
            except Exception as e:
                print(f"Warning: Could not load files cache: {e}")
                return {}
        return {}

    def save_files_cache(self, cache: dict[str, Any]) -> None:
        try:
            _write_json_atomic_locked(self.files_cache_path, cache, indent=2)
        except Exception as e:
            print(f"Warning: Could not save files cache: {e}")

    def get_sonarr_config(self) -> dict[str, Any] | None:
        sonarr = self.config.get("sonarr", {}) if isinstance(self.config, dict) else {}
        if not isinstance(sonarr, dict) or not sonarr.get("enabled", True):
            return None
        return sonarr

    def get_radarr_config(self) -> dict[str, Any] | None:
        radarr = self.config.get("radarr", {}) if isinstance(self.config, dict) else {}
        if not isinstance(radarr, dict) or not radarr.get("enabled", True):
            return None
        return radarr

    def get_settings(self) -> dict[str, Any]:
        settings = self.config.get("settings", {}) if isinstance(self.config, dict) else {}
        return settings if isinstance(settings, dict) else {}
