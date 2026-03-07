"""Shared path helpers."""

import os

from threep_commons.paths import resolve_app_data_dir

from ..constants import APP_IDENTITY


def get_app_cache_dir(fallback_dir: str) -> str:
    """Return the app cache directory, falling back if creation fails."""
    cache_dir = str(resolve_app_data_dir(APP_IDENTITY) / "cache")
    try:
        os.makedirs(cache_dir, exist_ok=True)
        return cache_dir
    except OSError:
        return fallback_dir
