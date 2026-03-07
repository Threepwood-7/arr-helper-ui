from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

TEST_DEPENDENCIES = [
    "pytest>=8.0",
    "pytest-cov>=5.0",
    "pytest-qt>=4.4",
    "PySide6>=6.10.2",
    "qbittorrent-api>=2025.11.1",
]


def ensure_venv(repo_root: Path) -> int:
    venv_dir = repo_root / ".venv"
    if venv_dir.exists():
        return 0

    setup_script = repo_root / "scripts" / "windows" / "setup_env.py"
    python_exe = shutil.which("python") or shutil.which("py")

    if not setup_script.exists():
        print(
            "ERROR: setup script missing at scripts\\windows\\setup_env.py.",
            file=sys.stderr,
        )
        return 1

    if python_exe is None:
        print("ERROR: python executable not found in PATH.", file=sys.stderr)
        print("Install Python and retry.", file=sys.stderr)
        return 1

    return subprocess.run([python_exe, str(setup_script)], cwd=repo_root, check=False).returncode


def ensure_test_dependencies(repo_root: Path) -> int:
    python_exe = get_python(repo_root)
    if not python_exe.exists():
        print("ERROR: local interpreter not found at .venv\\Scripts\\python.exe.", file=sys.stderr)
        print("Run: python scripts\\windows\\setup_env.py", file=sys.stderr)
        return 1

    has_pytest = subprocess.run(
        [str(python_exe), "-c", "import pytest"],
        cwd=repo_root,
        check=False,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    ).returncode == 0
    if has_pytest:
        return 0

    uv_exe = shutil.which("uv")
    if uv_exe is None:
        print("ERROR: uv executable not found in PATH.", file=sys.stderr)
        print("Install uv and retry.", file=sys.stderr)
        return 1

    print("Installing test dependencies into local environment with uv pip install")
    cmd = [uv_exe, "pip", "install", "--python", str(python_exe), *TEST_DEPENDENCIES]
    return subprocess.run(cmd, cwd=repo_root, check=False).returncode


def get_python(repo_root: Path) -> Path:
    return repo_root / ".venv" / "Scripts" / "python.exe"


def get_pythonw(repo_root: Path) -> Path:
    return repo_root / ".venv" / "Scripts" / "pythonw.exe"
