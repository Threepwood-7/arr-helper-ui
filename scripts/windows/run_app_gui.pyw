from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path


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


def ensure_local_pythonw(repo_root: Path) -> Path | None:
    pythonw_exe = repo_root / ".venv" / "Scripts" / "pythonw.exe"
    if pythonw_exe.exists():
        return pythonw_exe

    setup_script = repo_root / "scripts" / "windows" / "setup_env.py"
    python_exe = shutil.which("python") or shutil.which("py")
    if setup_script.exists() and python_exe:
        subprocess.run([python_exe, str(setup_script)], cwd=repo_root, check=False)

    if pythonw_exe.exists():
        return pythonw_exe
    return None


def main() -> int:
    repo_root = Path(__file__).resolve().parents[2]
    pythonw_exe = ensure_local_pythonw(repo_root)
    if pythonw_exe is None:
        show_error(
            "ERROR: local GUI interpreter not found.\n"
            "Run: python scripts\windows\setup_env.py"
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
