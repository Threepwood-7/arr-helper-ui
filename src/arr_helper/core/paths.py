"""Shared path helpers."""

import os
import tempfile


def get_app_cache_dir(fallback_dir: str) -> str:
    """Return the app cache directory, falling back if creation fails."""
    cache_dir = os.path.join(tempfile.gettempdir(), 'temp_arr_helper_ui')
    try:
        os.makedirs(cache_dir, exist_ok=True)
        return cache_dir
    except OSError:
        return fallback_dir
