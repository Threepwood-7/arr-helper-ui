"""Shared path helpers."""

import os

from ..runtime_paths import resolve_app_data_dir


def get_app_cache_dir(fallback_dir: str) -> str:
    """Return the app cache directory, falling back if creation fails."""
    cache_dir = str(resolve_app_data_dir() / "cache")
    try:
        os.makedirs(cache_dir, exist_ok=True)
        return cache_dir
    except OSError:
        return fallback_dir
