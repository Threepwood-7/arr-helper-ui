import os
import sys
from pathlib import Path

import pytest

# Ensure Qt tests run headless in CI/local terminals without a display server.
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

# Ensure src-layout package modules are importable when tests are executed from tests/.
ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))


class _FixtureAPI:
    def get_quality_profiles(self):
        return []


@pytest.fixture(autouse=True)
def _isolate_runtime_dirs(monkeypatch, tmp_path):
    config_dir = tmp_path / "config"
    data_dir = tmp_path / "data"
    monkeypatch.setenv("CONFIG_DIR", str(config_dir))
    monkeypatch.setenv("DATA_DIR", str(data_dir))

    from arr_helper.runtime_paths import configure_qsettings

    configure_qsettings(str(config_dir))
    yield


@pytest.fixture
def window(qtbot, monkeypatch):
    from arr_helper.sonarr_ui.main_window import MainWindow

    monkeypatch.setattr(MainWindow, "_start_worker", lambda self: None)
    win = MainWindow(_FixtureAPI(), settings={})
    qtbot.addWidget(win)
    return win
