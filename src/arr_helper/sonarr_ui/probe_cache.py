"""ffprobe cache and probing helpers for Sonarr UI."""

import contextlib
import json
import os
import subprocess
import threading
import time
from typing import Any

from threep_commons.subprocess_helpers import windows_no_window_run_kwargs

from ..core.paths import get_app_cache_dir

_FFPROBE: str | None = None  # resolved at startup in main()


def _get_app_cache_dir() -> str:
    return get_app_cache_dir(os.path.dirname(os.path.abspath(__file__)))


_PROBE_CACHE_PATH = os.path.join(_get_app_cache_dir(), "z_fprobe.cache")
_PROBE_CACHE_LOCK_PATH = f"{_PROBE_CACHE_PATH}.lock"
_probe_cache: dict = {}
_probe_cache_lock = threading.RLock()


def _pid_is_running(pid: int) -> bool:
    if pid <= 0:
        return False
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except OSError:
        return True
    return True


def _acquire_probe_lock(timeout_s: float = 10.0) -> tuple[int, str]:
    token = f"{os.getpid()}:{threading.get_ident()}:{time.time_ns()}"
    deadline = time.time() + timeout_s
    while True:
        try:
            fd = os.open(_PROBE_CACHE_LOCK_PATH, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, token.encode("ascii", errors="ignore"))
            return fd, token
        except FileExistsError:
            stale = False
            try:
                age = time.time() - os.path.getmtime(_PROBE_CACHE_LOCK_PATH)
                if age > 5:
                    owner = ""
                    try:
                        with open(_PROBE_CACHE_LOCK_PATH) as f:
                            owner = f.read().strip()
                    except OSError:
                        owner = ""
                    try:
                        owner_pid = int(owner.split(":", 1)[0]) if owner else 0
                    except (TypeError, ValueError):
                        owner_pid = 0
                    if owner_pid and not _pid_is_running(owner_pid):
                        stale = True
                    elif not owner_pid and age > 3600:
                        # Legacy or corrupted lock format with no owner pid.
                        stale = True
            except OSError:
                pass
            if stale:
                try:
                    os.remove(_PROBE_CACHE_LOCK_PATH)
                    continue
                except OSError:
                    pass
            if time.time() >= deadline:
                raise TimeoutError(
                    f"Timeout acquiring cache lock: {_PROBE_CACHE_LOCK_PATH}"
                ) from None
            time.sleep(0.05)


def _release_probe_lock(lock_fd: int, token: str):
    try:
        os.close(lock_fd)
    finally:
        try:
            owner = ""
            with open(_PROBE_CACHE_LOCK_PATH) as f:
                owner = f.read().strip()
            if owner == token:
                os.remove(_PROBE_CACHE_LOCK_PATH)
        except OSError:
            pass


def _load_probe_cache():
    global _probe_cache
    if os.path.exists(_PROBE_CACHE_PATH):
        try:
            with open(_PROBE_CACHE_PATH) as f:
                loaded = json.load(f)
            with _probe_cache_lock:
                _probe_cache = loaded if isinstance(loaded, dict) else {}
        except Exception:
            with _probe_cache_lock:
                _probe_cache = {}


def _save_probe_cache(replace: bool = False):
    with _probe_cache_lock:
        snapshot = dict(_probe_cache)
    lock_fd = None
    lock_token = ""
    tmp_path = f"{_PROBE_CACHE_PATH}.tmp.{os.getpid()}.{threading.get_ident()}"
    try:
        lock_fd, lock_token = _acquire_probe_lock()
        if not replace and os.path.exists(_PROBE_CACHE_PATH):
            try:
                with open(_PROBE_CACHE_PATH) as f:
                    existing = json.load(f)
                if isinstance(existing, dict):
                    existing.update(snapshot)
                    snapshot = existing
            except Exception:
                pass
        with open(tmp_path, "w") as f:
            json.dump(snapshot, f, indent=1)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp_path, _PROBE_CACHE_PATH)
    except Exception:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except OSError:
            pass
    finally:
        if lock_fd is not None:
            _release_probe_lock(lock_fd, lock_token)


def _as_int(value: Any) -> int:
    """Parse ffprobe-style numeric fields safely (e.g. 'N/A' -> 0)."""
    try:
        if value is None:
            return 0
        if isinstance(value, bool):
            return int(value)
        if isinstance(value, (int, float)):
            return int(value)
        text = str(value).strip()
        if not text or text.lower() in {"n/a", "na", "none", "null"}:
            return 0
        return int(float(text))
    except (TypeError, ValueError):
        return 0


def probe_file(file_path: str) -> dict:  # noqa: C901 - ffprobe metadata normalization
    """Return dict with codecs, resolution, bitrates, HDR, languages, size."""
    info: dict = {
        "video_codec": "",
        "video_resolution": "",
        "video_bitrate": "",
        "audio_codec": "",
        "audio_bitrate": "",
        "hdr": "",
        "audio_langs": [],
        "sub_langs": [],
        "size_bytes": 0,
    }
    with contextlib.suppress(OSError):
        info["size_bytes"] = os.path.getsize(file_path)

    # check cache — keyed by path, invalidated if size changed or fields missing
    with _probe_cache_lock:
        cached = _probe_cache.get(file_path)
    if (
        cached
        and cached.get("size_bytes") == info["size_bytes"]
        and cached.get("_probe_ok") is True
    ):
        return cached

    try:
        cmd = [
            _FFPROBE or "ffprobe",
            "-v",
            "quiet",
            "-print_format",
            "json",
            "-show_streams",
            "-show_format",
            file_path,
        ]
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=30,
            **windows_no_window_run_kwargs(),
        )
        if result.returncode != 0:
            return info
        data = json.loads(result.stdout)

        parsed: dict = {
            "video_codec": "",
            "video_resolution": "",
            "video_bitrate": "",
            "audio_codec": "",
            "audio_bitrate": "",
            "hdr": "",
            "audio_langs": [],
            "sub_langs": [],
            "size_bytes": info["size_bytes"],
        }

        for s in data.get("streams", []):
            codec_type = s.get("codec_type", "")
            lang = s.get("tags", {}).get("language", "")
            if codec_type == "video" and not parsed["video_codec"]:
                parsed["video_codec"] = s.get("codec_name", "").upper()
                w = _as_int(s.get("width", 0))
                h = _as_int(s.get("height", 0))
                if w and h:
                    parsed["video_resolution"] = f"{w}x{h}"
                vbr = _as_int(s.get("bit_rate", 0))
                if vbr:
                    parsed["video_bitrate"] = f"{vbr // 1000} kbps"
                color_transfer = s.get("color_transfer", "")
                color_space = s.get("color_space", "")
                side_data = s.get("side_data_list", [])
                has_hdr_transfer = color_transfer in ("smpte2084", "arib-std-b67")
                has_hdr_space = color_space in ("bt2020nc", "bt2020c")
                has_dovi = (
                    any(
                        sd.get("side_data_type", "")
                        in ("DOVI configuration record", "Dolby Vision configuration")
                        for sd in side_data
                    )
                    if side_data
                    else False
                )
                if has_dovi:
                    parsed["hdr"] = "DV"
                elif has_hdr_transfer or has_hdr_space:
                    parsed["hdr"] = "HDR"
                else:
                    parsed["hdr"] = "SDR"
            elif codec_type == "audio":
                if not parsed["audio_codec"]:
                    parsed["audio_codec"] = s.get("codec_name", "").upper()
                    abr = _as_int(s.get("bit_rate", 0))
                    if abr:
                        parsed["audio_bitrate"] = f"{abr // 1000} kbps"
                if lang:
                    parsed["audio_langs"].append(lang)
            elif codec_type == "subtitle":
                if lang:
                    parsed["sub_langs"].append(lang)

        if not parsed["video_bitrate"]:
            fmt_br = _as_int(data.get("format", {}).get("bit_rate", 0))
            if fmt_br:
                parsed["video_bitrate"] = f"{fmt_br // 1000} kbps"
    except Exception:
        return info

    parsed["_probe_ok"] = True
    with _probe_cache_lock:
        _probe_cache[file_path] = parsed
    return parsed


def set_ffprobe_path(ffprobe_path: str | None):
    global _FFPROBE
    _FFPROBE = ffprobe_path


def clear_probe_cache():
    global _probe_cache
    with _probe_cache_lock:
        _probe_cache = {}
    _save_probe_cache(replace=True)


def get_probe_cache_path() -> str:
    return _PROBE_CACHE_PATH
