"""File locking helpers for cache writes."""

import json
import os
import threading
import time
from collections.abc import Callable
from typing import Any


def pid_is_running(pid: int) -> bool:
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


def acquire_lock_file(
    lock_path: str,
    timeout_s: float = 10.0,
    pid_checker: Callable[[int], bool] | None = None,
) -> tuple[int, str]:
    token = f"{os.getpid()}:{threading.get_ident()}:{time.time_ns()}"
    deadline = time.time() + timeout_s
    check_pid = pid_checker or pid_is_running
    while True:
        try:
            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, token.encode("ascii", errors="ignore"))
            return fd, token
        except FileExistsError:
            stale = False
            try:
                age = time.time() - os.path.getmtime(lock_path)
                if age > 5:
                    owner = ""
                    try:
                        with open(lock_path) as f:
                            owner = f.read().strip()
                    except OSError:
                        owner = ""
                    try:
                        owner_pid = int(owner.split(":", 1)[0]) if owner else 0
                    except (TypeError, ValueError):
                        owner_pid = 0
                    if (owner_pid and not check_pid(owner_pid)) or (
                        not owner_pid and age > 3600
                    ):
                        stale = True
            except OSError:
                pass
            if stale:
                try:
                    os.remove(lock_path)
                    continue
                except OSError:
                    pass
            if time.time() >= deadline:
                raise TimeoutError(f"Timeout acquiring lock: {lock_path}") from None
            time.sleep(0.05)


def release_lock_file(lock_path: str, lock_fd: int, token: str):
    try:
        os.close(lock_fd)
    finally:
        try:
            owner = ""
            with open(lock_path) as f:
                owner = f.read().strip()
            if owner == token:
                os.remove(lock_path)
        except OSError:
            pass


def write_json_atomic_locked(path: str, payload: dict[str, Any], indent: int = 2):
    lock_path = f"{path}.lock"
    lock_fd, lock_token = acquire_lock_file(lock_path)
    tmp_path = f"{path}.tmp.{os.getpid()}.{threading.get_ident()}"
    try:
        with open(tmp_path, "w") as f:
            json.dump(payload, f, indent=indent)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp_path, path)
    except Exception:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except OSError:
            pass
        raise
    finally:
        release_lock_file(lock_path, lock_fd, lock_token)
