"""Locate ffprobe executable, checking PATH and common Windows install locations."""

import os
import platform
from pathlib import Path

from threep_commons.executables import find_first_available_executable


def find_ffprobe() -> str | None:
    """Return the full path to ffprobe, or None if not found.

    Checks PATH first, then common Windows installation directories.
    """
    candidates = _windows_candidates() if platform.system() == "Windows" else []
    found = find_first_available_executable(
        command_names=("ffprobe",),
        candidate_paths=candidates,
    )
    return str(found) if found is not None else None


def _windows_candidates() -> list[Path]:
    """Generate candidate ffprobe.exe paths for common Windows installs."""
    candidates: list[Path] = []

    # Direct / manual installs
    for base in (
        Path(r"C:\ffmpeg\bin"),
        Path(r"C:\Program Files\ffmpeg\bin"),
        Path(r"C:\Program Files (x86)\ffmpeg\bin"),
        Path(r"C:\tools\ffmpeg\bin"),
    ):
        candidates.append(base / "ffprobe.exe")

    # Chocolatey
    choco = Path(r"C:\ProgramData\chocolatey")
    env_choco = Path(str(os.environ.get("CHOCOLATEYINSTALL", str(choco))))
    candidates.append(env_choco / "bin" / "ffprobe.exe")

    # Scoop
    userprofile_text = os.environ.get("USERPROFILE", "")
    userprofile = Path(userprofile_text) if userprofile_text else None
    if userprofile is not None:
        candidates.append(userprofile / "scoop" / "shims" / "ffprobe.exe")

    # WinGet
    localappdata_text = os.environ.get("LOCALAPPDATA", "")
    if localappdata_text:
        candidates.append(
            Path(localappdata_text) / "Microsoft" / "WinGet" / "Links" / "ffprobe.exe"
        )

    # Scan C:\ and %USERPROFILE% for ffmpeg*/bin/ffprobe.exe (versioned extracts)
    for root_dir in (Path(r"C:\\"), userprofile):
        if root_dir is not None and root_dir.is_dir():
            try:
                for entry in root_dir.iterdir():
                    if entry.is_dir() and entry.name.lower().startswith("ffmpeg"):
                        candidate = entry / "bin" / "ffprobe.exe"
                        if candidate not in candidates:
                            candidates.append(candidate)
            except OSError:
                pass

    return candidates
