"""Utility helpers for Sonarr UI."""


def fmt_size(size_bytes: int) -> str:
    if size_bytes <= 0:
        return "0 GB"
    gb = size_bytes / (1024**3)
    if gb >= 1:
        return f"{gb:.2f} GB"
    mb = size_bytes / (1024**2)
    return f"{mb:.0f} MB"
