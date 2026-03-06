from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path


def main() -> int:
    repo_root = Path(__file__).resolve().parents[2]
    hatch_exe = shutil.which("hatch")

    if hatch_exe is None:
        print("ERROR: hatch executable not found in PATH.", file=sys.stderr)
        print("Install Hatch and retry.", file=sys.stderr)
        return 1

    cmd = [hatch_exe, "run", "test", *sys.argv[1:]]
    return subprocess.run(cmd, cwd=repo_root, check=False).returncode


if __name__ == "__main__":
    raise SystemExit(main())
