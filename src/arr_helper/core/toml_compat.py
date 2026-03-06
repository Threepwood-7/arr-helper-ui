"""TOML loading compatibility helpers."""

import tomllib as _toml
from pathlib import Path
from typing import Any


def load_toml(path: str | Path) -> dict[str, Any]:
    with open(path, "rb") as f:
        return _toml.load(f)
