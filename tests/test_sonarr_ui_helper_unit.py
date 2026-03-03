import json
from pathlib import Path

import pytest

import sonarr_ui_helper as sui


class _Resp:
    def __init__(self, json_data=None, text="{}"):
        self._json_data = json_data if json_data is not None else {}
        self.text = text

    def raise_for_status(self):
        return None

    def json(self):
        return self._json_data


def test_sonarr_api_uses_configured_timeout(monkeypatch):
    seen = []

    def _mk(name):
        def _fn(url, headers=None, auth=None, timeout=None, **kwargs):
            seen.append((name, timeout, url, kwargs))
            if name == "delete":
                return _Resp(text="")
            return _Resp({})

        return _fn

    monkeypatch.setattr(sui.requests, "get", _mk("get"))
    monkeypatch.setattr(sui.requests, "post", _mk("post"))
    monkeypatch.setattr(sui.requests, "put", _mk("put"))
    monkeypatch.setattr(sui.requests, "delete", _mk("delete"))

    api = sui.SonarrAPI("http://x", "k", request_timeout=(1, 2))
    api.get_series()
    api.command({"name": "SeriesSearch"})
    api.update_series({"id": 1})
    api.delete_series(1)

    assert all(timeout == (1, 2) for _, timeout, _, _ in seen)


@pytest.fixture
def probe_cache_paths(tmp_path, monkeypatch):
    cache_path = tmp_path / "z_fprobe.cache"
    lock_path = tmp_path / "z_fprobe.cache.lock"
    monkeypatch.setattr(sui, "_PROBE_CACHE_PATH", str(cache_path))
    monkeypatch.setattr(sui, "_PROBE_CACHE_LOCK_PATH", str(lock_path))
    monkeypatch.setattr(sui, "_probe_cache", {})
    return cache_path, lock_path


def test_save_probe_cache_merges_existing_when_replace_false(probe_cache_paths, monkeypatch):
    cache_path, _ = probe_cache_paths
    cache_path.write_text(json.dumps({"old": {"size_bytes": 1}}), encoding="utf-8")
    monkeypatch.setattr(sui, "_probe_cache", {"new": {"size_bytes": 2}})

    sui._save_probe_cache(replace=False)

    saved = json.loads(cache_path.read_text(encoding="utf-8"))
    assert saved["old"]["size_bytes"] == 1
    assert saved["new"]["size_bytes"] == 2


def test_save_probe_cache_replaces_when_replace_true(probe_cache_paths, monkeypatch):
    cache_path, _ = probe_cache_paths
    cache_path.write_text(json.dumps({"old": {"size_bytes": 1}}), encoding="utf-8")
    monkeypatch.setattr(sui, "_probe_cache", {"new": {"size_bytes": 2}})

    sui._save_probe_cache(replace=True)

    saved = json.loads(cache_path.read_text(encoding="utf-8"))
    assert "old" not in saved
    assert saved["new"]["size_bytes"] == 2


def test_probe_file_uses_cache_without_subprocess(monkeypatch, tmp_path):
    media = tmp_path / "a.mkv"
    media.write_bytes(b"123456")
    cached = {
        "video_codec": "H264",
        "video_resolution": "1920x1080",
        "video_bitrate": "1000 kbps",
        "audio_codec": "AAC",
        "audio_bitrate": "128 kbps",
        "hdr": "SDR",
        "audio_langs": ["eng"],
        "sub_langs": ["eng"],
        "size_bytes": media.stat().st_size,
        "_probe_ok": True,
    }
    monkeypatch.setattr(sui, "_probe_cache", {str(media): cached})

    def _no_run(*args, **kwargs):
        raise AssertionError("subprocess.run should not be called for cache hits")

    monkeypatch.setattr(sui.subprocess, "run", _no_run)

    assert sui.probe_file(str(media)) == cached


def test_probe_file_parses_ffprobe_json(monkeypatch, tmp_path):
    media = tmp_path / "b.mkv"
    media.write_bytes(b"abcd")
    monkeypatch.setattr(sui, "_probe_cache", {})

    payload = {
        "streams": [
            {
                "codec_type": "video",
                "codec_name": "hevc",
                "width": 3840,
                "height": 2160,
                "bit_rate": "8000000",
                "color_transfer": "smpte2084",
            },
            {
                "codec_type": "audio",
                "codec_name": "aac",
                "bit_rate": "256000",
                "tags": {"language": "eng"},
            },
            {"codec_type": "subtitle", "tags": {"language": "eng"}},
        ],
        "format": {"bit_rate": "9000000"},
    }

    class _RunResult:
        returncode = 0
        stdout = json.dumps(payload)

    monkeypatch.setattr(sui.subprocess, "run", lambda *a, **k: _RunResult())

    result = sui.probe_file(str(media))
    assert result["video_codec"] == "HEVC"
    assert result["video_resolution"] == "3840x2160"
    assert result["audio_codec"] == "AAC"
    assert result["hdr"] == "HDR"
    assert result["audio_langs"] == ["eng"]
    assert result["sub_langs"] == ["eng"]


def test_probe_file_does_not_cache_failed_probe(monkeypatch, tmp_path):
    media = tmp_path / "c.mkv"
    media.write_bytes(b"abcdef")
    monkeypatch.setattr(sui, "_probe_cache", {})

    calls = {"n": 0}

    class _FailResult:
        returncode = 1
        stdout = ""

    class _OkResult:
        returncode = 0
        stdout = json.dumps({"streams": []})

    def _run(*_args, **_kwargs):
        calls["n"] += 1
        return _FailResult() if calls["n"] == 1 else _OkResult()

    monkeypatch.setattr(sui.subprocess, "run", _run)

    first = sui.probe_file(str(media))
    second = sui.probe_file(str(media))

    assert calls["n"] == 2
    assert first["video_codec"] == ""
    assert second.get("_probe_ok") is True


class _APIGetSeriesError:
    def get_series(self):
        raise RuntimeError("boom")


def test_loadworker_emits_error_and_empty_when_series_fetch_fails():
    worker = sui.LoadWorker(_APIGetSeriesError())
    progress = []
    ready = []
    worker.progress.connect(progress.append)
    worker.series_ready.connect(ready.append)

    worker.run()

    assert any(p.startswith("Error:") for p in progress)
    assert ready == [[]]


class _APIWithWarnings:
    def get_series(self):
        return [
            {"id": 1, "title": "Bad", "year": 2022, "path": "x", "seasons": []},
            {"id": 2, "title": "Good", "year": 2023, "path": "y", "seasons": [{"seasonNumber": 1}]},
        ]

    def get_episode_files(self, series_id):
        if series_id == 1:
            raise RuntimeError("epfiles fail")
        return [{"id": 20, "path": "good.mkv"}]

    def get_episodes(self, series_id):
        if series_id == 1:
            raise RuntimeError("episodes fail")
        return [{"id": 200, "episodeFileId": 20, "seasonNumber": 1, "episodeNumber": 1}]


def test_loadworker_reports_warnings_and_finishes(monkeypatch):
    worker = sui.LoadWorker(_APIWithWarnings())
    progress = []
    ready = []
    worker.progress.connect(progress.append)
    worker.series_ready.connect(ready.append)

    monkeypatch.setattr(
        sui,
        "probe_file",
        lambda _p: {
            "video_codec": "H264",
            "video_resolution": "1920x1080",
            "video_bitrate": "1000 kbps",
            "audio_codec": "AAC",
            "audio_bitrate": "128 kbps",
            "hdr": "SDR",
            "audio_langs": ["eng"],
            "sub_langs": ["eng"],
            "size_bytes": 123,
        },
    )
    monkeypatch.setattr(sui, "_save_probe_cache", lambda *a, **k: None)

    worker.run()

    assert any("Warning: failed episode files" in p for p in progress)
    assert any("Done with warnings:" in p for p in progress)
    assert len(ready) == 1
    assert len(ready[0]) == 1
    assert ready[0][0]["title"] == "Good"


class _APIBadShape:
    def get_series(self):
        return [None]


def test_loadworker_unexpected_exception_is_reported():
    worker = sui.LoadWorker(_APIBadShape())
    progress = []
    ready = []
    worker.progress.connect(progress.append)
    worker.series_ready.connect(ready.append)

    worker.run()

    assert any("unexpected loader failure" in p for p in progress)
    assert ready == [[]]
