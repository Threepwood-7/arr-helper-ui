"""TOML loading compatibility helpers."""

from pathlib import Path
from typing import Any

try:
    import tomllib as _toml
except ModuleNotFoundError:  # pragma: no cover - only on Python < 3.11
    import tomli as _toml


def load_toml(path: str | Path) -> dict[str, Any]:
    with open(path, 'rb') as f:
        return _toml.load(f)
