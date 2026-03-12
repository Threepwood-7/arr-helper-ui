from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path


def run(cmd: list[str], cwd: Path) -> int:
    return subprocess.run(cmd, cwd=cwd, check=False).returncode


def main() -> int:
    repo_root = Path(__file__).resolve().parents[2]
    pythonw_exe = repo_root / ".venv" / "Scripts" / "pythonw.exe"

    uv_exe = shutil.which("uv")
    if uv_exe is None:
        print("ERROR: uv executable not found in PATH.", file=sys.stderr)
        print("Install uv and retry.", file=sys.stderr)
        return 1

    print("Synchronizing project environment with: uv sync --locked")
    rc = run([uv_exe, "sync", "--locked"], repo_root)
    if rc != 0:
        print("Locked sync failed; retrying with: uv sync", file=sys.stderr)
        rc = run([uv_exe, "sync"], repo_root)
        if rc != 0:
            print("ERROR: failed to create .venv with uv sync.", file=sys.stderr)
            return rc

    if not pythonw_exe.exists():
        print(
            "ERROR: setup completed but .venv\\Scripts\\pythonw.exe is still missing.",
            file=sys.stderr,
        )
        return 1

    print(f"Environment ready: {pythonw_exe}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
