from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from _common import ensure_test_dependencies, ensure_venv, get_python


def main() -> int:
    repo_root = Path(__file__).resolve().parents[2]
    rc = ensure_venv(repo_root)
    if rc != 0:
        return rc

    rc = ensure_test_dependencies(repo_root)
    if rc != 0:
        return rc

    python_exe = get_python(repo_root)
    if not python_exe.exists():
        print("ERROR: local interpreter not found at .venv\\Scripts\\python.exe.", file=sys.stderr)
        print("Run: python scripts\\windows\\setup_env.py", file=sys.stderr)
        return 1

    cmd = [str(python_exe), "-m", "pytest", *sys.argv[1:]]
    return subprocess.run(cmd, cwd=repo_root, check=False).returncode


if __name__ == "__main__":
    raise SystemExit(main())
