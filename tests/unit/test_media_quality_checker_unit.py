import json
import os

import pytest

from arr_helper.core import locking as core_locking
from arr_helper.media_checker import checker as mqc


class _DummyConfig:
    def __init__(self):
        self._user = {}
        self._files = {}

    def load_user_cache(self):
        return self._user

    def load_files_cache(self):
        return self._files

    def save_user_cache(self, data):
        self._user = data

    def save_files_cache(self, data):
        self._files = data


class _Response:
    def __init__(self, json_data=None, text="{}", raise_exc=None):
        self._json_data = json_data
        self.text = text
        self._raise_exc = raise_exc

    def raise_for_status(self):
        if self._raise_exc:
            raise self._raise_exc

    def json(self):
        if isinstance(self._json_data, Exception):
            raise self._json_data
        return self._json_data


def _checker():
    return mqc.MediaQualityChecker(
        sonarr_url="http://sonarr.local",
        sonarr_api="sonarr-key",
        radarr_url="http://radarr.local",
        radarr_api="radarr-key",
        interactive=False,
        config=_DummyConfig(),
    )


def test_make_request_returns_none_on_non_json_success(monkeypatch):
    checker = _checker()

    def _fake_get(*args, **kwargs):
        return _Response(json_data=ValueError("bad"), text="OK")

    monkeypatch.setattr(mqc.requests, "get", _fake_get)

    result = checker._make_request(
        checker.sonarr_url, checker.sonarr_api, "series", auth=checker.sonarr_http_auth
    )
    assert result is None


def test_make_request_uses_explicit_auth(monkeypatch):
    checker = _checker()
    seen = {}

    def _fake_get(*args, **kwargs):
        seen["auth"] = kwargs.get("auth")
        return _Response(json_data={}, text="{}")

    monkeypatch.setattr(mqc.requests, "get", _fake_get)
    checker._make_request(
        checker.sonarr_url, checker.sonarr_api, "series", auth=("u", "p")
    )
    assert seen["auth"] == ("u", "p")


def test_download_release_treats_empty_json_as_success(monkeypatch):
    checker = _checker()
    monkeypatch.setattr(checker, "_make_request", lambda *a, **k: {})

    ok = checker.download_release(
        checker.sonarr_url,
        checker.sonarr_api,
        {"guid": "g1", "indexerId": 11},
        is_sonarr=True,
    )
    assert ok is True


@pytest.mark.parametrize(
    ("has_eng_audio", "has_eng_subs", "expected"),
    [
        (True, True, False),
        (False, True, True),
        (True, False, True),
    ],
)
def test_should_redownload_matrix(
    has_eng_audio: bool, has_eng_subs: bool, expected: bool
):
    checker = _checker()
    checker.require_audio = True
    checker.require_subs = True

    assert checker.should_redownload(has_eng_audio, has_eng_subs) is expected


def test_check_file_streams_parses_english_streams(monkeypatch, tmp_path):
    checker = _checker()
    media = tmp_path / "movie.mkv"
    media.write_bytes(b"x")

    payload = {
        "streams": [
            {"codec_type": "audio", "tags": {"language": "eng"}},
            {"codec_type": "subtitle", "tags": {"language": "en"}},
        ]
    }

    class _RunResult:
        returncode = 0
        stdout = json.dumps(payload)
        stderr = ""

    monkeypatch.setattr(mqc.subprocess, "run", lambda *a, **k: _RunResult())

    has_audio, has_subs = checker.check_file_streams(str(media))
    assert has_audio is True
    assert has_subs is True


def test_get_episode_releases_builds_quality_profile_endpoint(monkeypatch):
    checker = _checker()
    seen = {}

    def _fake_make(url, api_key, endpoint, **kwargs):
        seen["url"] = url
        seen["api_key"] = api_key
        seen["endpoint"] = endpoint
        return []

    monkeypatch.setattr(checker, "_make_request", _fake_make)
    checker.get_episode_releases(42, quality_profile_id=9)

    assert seen["url"] == checker.sonarr_url
    assert seen["api_key"] == checker.sonarr_api
    assert seen["endpoint"] == "release?episodeId=42&qualityProfileId=9"


def test_process_sonarr_noninteractive_delete_then_search(monkeypatch):
    checker = _checker()
    calls = []
    monkeypatch.setattr(checker, "check_file_streams", lambda _p: (False, False))

    def _fake_make(url, api_key, endpoint, method="GET", data=None, auth=None):
        calls.append((method, endpoint, data))
        if endpoint == "series":
            return [{"id": 1, "title": "Show", "qualityProfileId": 10}]
        if endpoint == "episodefile?seriesId=1":
            return [{"path": "show.mkv", "id": 11}]
        if endpoint == "episode?seriesId=1":
            return [{"id": 101, "episodeFileId": 11}]
        if endpoint == "episodefile/11":
            return {}
        if endpoint == "command":
            return {}
        return {}

    monkeypatch.setattr(checker, "_make_request", _fake_make)
    checker.process_sonarr(dry_run=False)

    endpoints = [ep for _, ep, _ in calls]
    assert endpoints.index("episode?seriesId=1") < endpoints.index("episodefile/11")
    assert ("POST", "command", {"name": "EpisodeSearch", "episodeIds": [101]}) in calls


def test_process_radarr_noninteractive_delete_and_search(monkeypatch):
    checker = _checker()
    calls = []
    monkeypatch.setattr(checker, "check_file_streams", lambda _p: (False, False))

    def _fake_make(url, api_key, endpoint, method="GET", data=None, auth=None):
        calls.append((method, endpoint, data))
        if endpoint == "movie":
            return [
                {
                    "id": 2,
                    "title": "Movie",
                    "hasFile": True,
                    "qualityProfileId": 20,
                    "movieFile": {"id": 22, "path": "movie.mkv"},
                }
            ]
        if endpoint == "moviefile/22":
            return {}
        if endpoint == "command":
            return {}
        return {}

    monkeypatch.setattr(checker, "_make_request", _fake_make)
    checker.process_radarr(dry_run=False)

    assert ("DELETE", "moviefile/22", None) in calls
    assert ("POST", "command", {"name": "MoviesSearch", "movieIds": [2]}) in calls


def test_cache_signature_invalidation_on_file_change(tmp_path):
    checker = _checker()
    media = tmp_path / "episode.mkv"
    media.write_bytes(b"a")

    sig = checker._file_signature(str(media))
    cache = {str(media): sig}
    assert checker._is_cached_match(cache, str(media)) is True

    media.write_bytes(b"ab")
    assert checker._is_cached_match(cache, str(media)) is False
    assert str(media) not in cache


def test_acquire_lock_file_reclaims_stale_dead_owner(tmp_path, monkeypatch):
    lock_path = tmp_path / "cache.lock"
    lock_path.write_text("999999:1:1", encoding="utf-8")
    os.utime(lock_path, (1, 1))

    fd, token = core_locking.acquire_lock_file(
        str(lock_path),
        timeout_s=0.5,
        pid_checker=lambda _pid: False,
    )
    try:
        assert token.startswith(f"{os.getpid()}:")
    finally:
        core_locking.release_lock_file(str(lock_path), fd, token)

    assert not lock_path.exists()


def test_acquire_lock_file_does_not_steal_live_owner_even_if_old(tmp_path, monkeypatch):
    lock_path = tmp_path / "cache.lock"
    # Old lock file, but with a live owner pid marker.
    lock_path.write_text(f"{os.getpid()}:1:1", encoding="utf-8")
    os.utime(lock_path, (1, 1))

    with pytest.raises(TimeoutError):
        core_locking.acquire_lock_file(
            str(lock_path),
            timeout_s=0.05,
            pid_checker=lambda _pid: True,
        )
