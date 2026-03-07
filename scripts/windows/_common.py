from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path


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


def get_python(repo_root: Path) -> Path:
    return repo_root / ".venv" / "Scripts" / "python.exe"


def get_pythonw(repo_root: Path) -> Path:
    return repo_root / ".venv" / "Scripts" / "pythonw.exe"
