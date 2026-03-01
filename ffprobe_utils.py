"""Locate ffprobe executable, checking PATH and common Windows install locations."""

import os
import platform
import shutil
import subprocess


def find_ffprobe() -> str | None:
    """Return the full path to ffprobe, or None if not found.

    Checks PATH first, then common Windows installation directories.
    """
    found = shutil.which('ffprobe')
    if found:
        return found

    if platform.system() == 'Windows':
        for path in _windows_candidates():
            if os.path.isfile(path):
                return path

    return None


def ffprobe_subprocess_kwargs() -> dict:
    """Return subprocess kwargs to keep ffprobe hidden on Windows."""
    if platform.system() != 'Windows':
        return {}

    kwargs = {}
    try:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = getattr(subprocess, 'SW_HIDE', 0)
        kwargs['startupinfo'] = startupinfo
    except Exception:
        pass

    kwargs['creationflags'] = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
    return kwargs


def _windows_candidates() -> list[str]:
    """Generate candidate ffprobe.exe paths for common Windows installs."""
    candidates = []

    # Direct / manual installs
    for base in [
        r'C:\ffmpeg\bin',
        r'C:\Program Files\ffmpeg\bin',
        r'C:\Program Files (x86)\ffmpeg\bin',
        r'C:\tools\ffmpeg\bin',
    ]:
        candidates.append(os.path.join(base, 'ffprobe.exe'))

    # Chocolatey
    choco = os.environ.get('ChocolateyInstall', r'C:\ProgramData\chocolatey')
    candidates.append(os.path.join(choco, 'bin', 'ffprobe.exe'))

    # Scoop
    userprofile = os.environ.get('USERPROFILE', '')
    if userprofile:
        candidates.append(os.path.join(userprofile, 'scoop', 'shims', 'ffprobe.exe'))

    # WinGet
    localappdata = os.environ.get('LOCALAPPDATA', '')
    if localappdata:
        candidates.append(os.path.join(localappdata, 'Microsoft', 'WinGet', 'Links', 'ffprobe.exe'))

    # Scan C:\ and %USERPROFILE% for ffmpeg*/bin/ffprobe.exe (versioned extracts)
    for root_dir in ['C:\\', userprofile]:
        if root_dir and os.path.isdir(root_dir):
            try:
                for entry in os.scandir(root_dir):
                    if entry.is_dir() and entry.name.lower().startswith('ffmpeg'):
                        p = os.path.join(entry.path, 'bin', 'ffprobe.exe')
                        if p not in candidates:
                            candidates.append(p)
            except OSError:
                pass

    return candidates
