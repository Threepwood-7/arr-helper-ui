from __future__ import annotations

import pytest

from arr_helper.core import ffprobe, locking, paths
from arr_helper.sonarr_ui import helpers


def test_write_json_atomic_locked_writes_payload_and_cleans_lock(tmp_path) -> None:
    target = tmp_path / "cache.json"
    locking.write_json_atomic_locked(str(target), {"ok": True})

    assert target.exists()
    assert target.read_text(encoding="utf-8").strip().startswith("{")
    assert not (tmp_path / "cache.json.lock").exists()


def test_find_ffprobe_prefers_path_lookup(monkeypatch) -> None:
    monkeypatch.setattr(
        "threep_commons.executables.shutil.which",
        lambda _: r"C:\bin\ffprobe.exe",
    )

    assert ffprobe.find_ffprobe() == r"C:\bin\ffprobe.exe"


def test_find_ffprobe_uses_windows_candidates(monkeypatch, tmp_path) -> None:
    monkeypatch.setattr("threep_commons.executables.shutil.which", lambda _: None)
    monkeypatch.setattr(ffprobe.platform, "system", lambda: "Windows")
    candidate = tmp_path / "ffprobe.exe"
    candidate.write_text("", encoding="utf-8")
    monkeypatch.setattr(ffprobe, "_windows_candidates", lambda: [str(candidate)])

    assert ffprobe.find_ffprobe() == str(candidate)


def test_get_app_cache_dir_falls_back_when_creation_fails(monkeypatch, tmp_path) -> None:
    monkeypatch.setattr(paths.os, "makedirs", lambda *_a, **_k: (_ for _ in ()).throw(OSError("boom")))

    assert paths.get_app_cache_dir(str(tmp_path)) == str(tmp_path)


def test_fmt_size_handles_mb_and_gb() -> None:
    assert helpers.fmt_size(0) == "0 GB"
    assert helpers.fmt_size(50 * 1024 * 1024) == "50 MB"
    assert helpers.fmt_size(3 * 1024 * 1024 * 1024) == "3.00 GB"


def test_dialog_exports_are_available() -> None:
    pytest.importorskip("PySide6")
    from arr_helper.sonarr_ui.dialogs import AddShowDialog, ManualSearchDialog

    assert AddShowDialog.__name__ == "AddShowDialog"
    assert ManualSearchDialog.__name__ == "ManualSearchDialog"
