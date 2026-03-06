"""Configuration and cache persistence for media quality checker."""

import json
import os
import shutil
import sys
from pathlib import Path

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
from ..core.toml_compat import load_toml

APP_SLUG = "arr-helper-ui"
DEFAULT_CONFIG_REL = Path("config/app.defaults.toml")
EXAMPLE_CONFIG_REL = Path("config/app.example.toml")
LOCAL_CONFIG_REL = Path("config/app.local.toml")
APP_CONFIG_PATH_ENV = "APP_CONFIG_PATH"
APP_SECRETS_PATH_ENV = "APP_SECRETS_PATH"
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


def _pid_is_running(pid: int) -> bool:
    return _core_pid_is_running(pid)


def _acquire_lock_file(lock_path: str, timeout_s: float = 10.0) -> tuple[int, str]:
    return _core_acquire_lock_file(lock_path, timeout_s=timeout_s, pid_checker=_pid_is_running)


def _release_lock_file(lock_path: str, lock_fd: int, token: str):
    _core_release_lock_file(lock_path, lock_fd, token)


def _write_json_atomic_locked(path: str, payload: dict, indent: int = 2):
    _core_write_json_atomic_locked(path, payload, indent=indent)


def _deep_merge_dicts(base: dict, overlay: dict) -> dict:
    merged: dict = dict(base)
    for key, value in overlay.items():
        base_value = merged.get(key)
        if isinstance(base_value, dict) and isinstance(value, dict):
            merged[key] = _deep_merge_dicts(base_value, value)
        else:
            merged[key] = value
    return merged


def _set_nested_value(target: dict, key_path: tuple[str, ...], value: str) -> None:
    current = target
    for key in key_path[:-1]:
        next_value = current.get(key)
        if not isinstance(next_value, dict):
            next_value = {}
            current[key] = next_value
        current = next_value
    current[key_path[-1]] = value


class Config:
    """Load and manage configuration from TOML file."""

    def __init__(self, config_path: str | None = None):
        self._cwd = Path.cwd()
        self.defaults_path = str(self._cwd / DEFAULT_CONFIG_REL)
        self.example_path = str(self._cwd / EXAMPLE_CONFIG_REL)
        self.config_path = str(self._resolve_local_config_path(config_path))
        self.secrets_path = str(self._resolve_secrets_path())
        self.config = self._load_config()
        cache_dir = get_app_cache_dir(str(self._cwd))
        self.user_cache_path = os.path.join(cache_dir, "z_user.cache")
        self.files_cache_path = os.path.join(cache_dir, "z_files.cache")

    def _resolve_local_config_path(self, config_path: str | None) -> Path:
        explicit_path = (
            config_path
            if config_path is not None
            else os.environ.get(APP_CONFIG_PATH_ENV, "")
        )
        if explicit_path:
            path = Path(explicit_path).expanduser()
            return path if path.is_absolute() else self._cwd / path
        return self._cwd / LOCAL_CONFIG_REL

    def _resolve_secrets_path(self) -> Path:
        explicit_path = os.environ.get(APP_SECRETS_PATH_ENV, "")
        if explicit_path:
            path = Path(explicit_path).expanduser()
            return path if path.is_absolute() else self._cwd / path
        if os.name == "nt":
            appdata = os.environ.get("APPDATA")
            if appdata:
                return Path(appdata) / APP_SLUG / "secrets.toml"
            return Path.home() / "AppData" / "Roaming" / APP_SLUG / "secrets.toml"
        return Path.home() / ".config" / APP_SLUG / "secrets.toml"

    def _load_optional_toml(self, path: str) -> dict:
        if not os.path.exists(path):
            return {}
        loaded = load_toml(path)
        return loaded if isinstance(loaded, dict) else {}

    def _apply_secret_env_overrides(self, config: dict) -> None:
        for env_name, key_path in SECRET_ENV_TO_KEYS:
            env_value = os.environ.get(env_name, "")
            if env_value:
                _set_nested_value(config, key_path, env_value)

    def _load_config(self) -> dict:
        if not os.path.exists(self.defaults_path):
            print(f"Error: Config defaults file not found: {self.defaults_path}")
            sys.exit(1)

        if not os.path.exists(self.config_path):
            print(f"Warning: Local config file not found: {self.config_path}")
            print("Creating local config template from defaults/example...")
            self._create_local_config()

        defaults = self._load_optional_toml(self.defaults_path)
        local_config = self._load_optional_toml(self.config_path)
        secrets_config = self._load_optional_toml(self.secrets_path)
        merged = _deep_merge_dicts(defaults, local_config)
        merged = _deep_merge_dicts(merged, secrets_config)
        self._apply_secret_env_overrides(merged)
        return merged

    def load_user_cache(self) -> dict:
        if os.path.exists(self.user_cache_path):
            try:
                with open(self.user_cache_path) as f:
                    return json.load(f)
            except Exception as e:
                print(f"Warning: Could not load user cache: {e}")
                return {}
        return {}

    def save_user_cache(self, cache: dict):
        try:
            _write_json_atomic_locked(self.user_cache_path, cache, indent=2)
        except Exception as e:
            print(f"Warning: Could not save user cache: {e}")

    def load_files_cache(self) -> dict:
        if os.path.exists(self.files_cache_path):
            try:
                with open(self.files_cache_path) as f:
                    return json.load(f)
            except Exception as e:
                print(f"Warning: Could not load files cache: {e}")
                return {}
        return {}

    def save_files_cache(self, cache: dict):
        try:
            _write_json_atomic_locked(self.files_cache_path, cache, indent=2)
        except Exception as e:
            print(f"Warning: Could not save files cache: {e}")

    def _create_local_config(self):
        local_path = Path(self.config_path)
        local_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            if os.path.exists(self.example_path):
                shutil.copy2(self.example_path, self.config_path)
            elif os.path.exists(self.defaults_path):
                shutil.copy2(self.defaults_path, self.config_path)
            else:
                local_path.write_text("", encoding="utf-8")
            print(f"Created local config at: {self.config_path}")
            print("Please edit this file with your Sonarr/Radarr settings.")
        except Exception as e:
            print(f"Error creating local config: {e}")

    def validate(self):
        errors = []
        placeholders = {
            'your_sonarr_api_key_here',
            'your_radarr_api_key_here',
            'your-sonarr-api-key-here',
            'your-radarr-api-key-here',
            '',
        }

        sonarr = self.config.get("sonarr", {})
        radarr = self.config.get("radarr", {})
        sonarr_enabled = sonarr.get("enabled", True)
        radarr_enabled = radarr.get("enabled", True)

        for name, section, enabled in [("sonarr", sonarr, sonarr_enabled), ("radarr", radarr, radarr_enabled)]:
            if not enabled:
                continue
            if "url" not in section or not section["url"].strip():
                errors.append(f"[{name}] 'url' is missing or empty")
            if "api_key" not in section:
                errors.append(f"[{name}] 'api_key' is missing")
            elif section["api_key"].strip().lower() in placeholders:
                errors.append(f"[{name}] 'api_key' is still set to a placeholder value")

        if errors:
            print("Config validation failed:")
            for err in errors:
                print(f"  - {err}")
            print(f"\nPlease edit {self.config_path}")
            sys.exit(1)

    def get_sonarr_config(self) -> dict | None:
        sonarr = self.config.get("sonarr", {})
        if not sonarr.get("enabled", True):
            return None
        return sonarr

    def get_radarr_config(self) -> dict | None:
        radarr = self.config.get("radarr", {})
        if not radarr.get("enabled", True):
            return None
        return radarr

    def get_settings(self) -> dict:
        return self.config.get("settings", {})
