"""Utility helpers for Sonarr UI."""

import os
import platform
import subprocess


def _open_path(path: str):
    """Open a file or directory with the system default handler (cross-platform)."""
    system = platform.system()
    if system == 'Windows':
        os.startfile(path)
    elif system == 'Darwin':
        subprocess.Popen(['open', path])
    else:
        subprocess.Popen(['xdg-open', path])


def fmt_size(size_bytes: int) -> str:
    if size_bytes <= 0:
        return '0 GB'
    gb = size_bytes / (1024 ** 3)
    if gb >= 1:
        return f'{gb:.2f} GB'
    mb = size_bytes / (1024 ** 2)
    return f'{mb:.0f} MB'
