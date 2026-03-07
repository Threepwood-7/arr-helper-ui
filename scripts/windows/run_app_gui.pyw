from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from _common import ensure_venv, get_pythonw


def show_error(message: str) -> None:
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("run_app_gui", message)
        root.destroy()
    except Exception:
        print(message, file=sys.stderr)


def main() -> int:
    repo_root = Path(__file__).resolve().parents[2]
    rc = ensure_venv(repo_root)
    if rc != 0:
        show_error(
            "ERROR: failed to create local environment.\n"
            "Run: python scripts\\windows\\setup_env.py"
        )
        return rc

    pythonw_exe = get_pythonw(repo_root)
    if not pythonw_exe.exists():
        show_error(
            "ERROR: local GUI interpreter not found.\n"
            "Run: python scripts\\windows\\setup_env.py"
        )
        return 1

    cmd = [str(pythonw_exe), "-m", "arr_helper", *sys.argv[1:]]
    kwargs: dict[str, object] = {"cwd": repo_root}
    if sys.platform == "win32":
        kwargs["creationflags"] = (
            getattr(subprocess, "DETACHED_PROCESS", 0)
            | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
            | getattr(subprocess, "CREATE_NO_WINDOW", 0)
        )
    subprocess.Popen(cmd, **kwargs)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
