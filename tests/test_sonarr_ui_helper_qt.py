from pathlib import Path

from PySide6.QtGui import QCloseEvent
from PySide6.QtWidgets import QMessageBox

import sonarr_ui_helper as sui


class _MinimalAPI:
    def get_quality_profiles(self):
        return []


def test_mainwindow_uses_loader_api_for_worker(qtbot, monkeypatch):
    seen = {}

    class _Sig:
        def connect(self, _fn):
            return None

    class _DummyWorker:
        def __init__(self, api):
            seen["api"] = api
            self.progress = _Sig()
            self.series_ready = _Sig()

        def start(self):
            seen["started"] = True

        def isRunning(self):
            return False

    monkeypatch.setattr(sui, "LoadWorker", _DummyWorker)

    api = _MinimalAPI()
    loader_api = object()
    win = sui.MainWindow(api, loader_api=loader_api, settings={})
    qtbot.addWidget(win)

    assert seen["api"] is loader_api
    assert seen["started"] is True


def test_mainwindow_empty_data_prefers_last_error(qtbot, monkeypatch):
    monkeypatch.setattr(sui.MainWindow, "_start_worker", lambda self: None)

    win = sui.MainWindow(_MinimalAPI(), settings={})
    qtbot.addWidget(win)
    win._last_worker_error = "Error: boom"
    win._worker_warning_count = 3

    win._on_data_loaded([])

    assert win.status_label.text() == "Error: boom"


def test_close_event_force_close_blocks_if_worker_still_running(qtbot, monkeypatch):
    monkeypatch.setattr(sui.MainWindow, "_start_worker", lambda self: None)
    win = sui.MainWindow(_MinimalAPI(), settings={})
    qtbot.addWidget(win)

    class _StuckWorker:
        def __init__(self):
            self.terminate_called = False

        def isRunning(self):
            return True

        def terminate(self):
            self.terminate_called = True

        def wait(self, _timeout):
            return False

    stuck = _StuckWorker()
    win.worker = stuck
    monkeypatch.setattr(win, "_stop_worker", lambda _timeout: False)
    monkeypatch.setattr(QMessageBox, "question", lambda *a, **k: QMessageBox.Yes)

    critical_calls = []
    monkeypatch.setattr(QMessageBox, "critical", lambda *a, **k: critical_calls.append((a, k)))

    event = QCloseEvent()
    win.closeEvent(event)

    assert not event.isAccepted()
    assert stuck.terminate_called is True
    assert len(critical_calls) == 1


def test_mainwindow_uses_ini_qsettings_backend(qtbot, monkeypatch):
    monkeypatch.setattr(sui.MainWindow, "_start_worker", lambda self: None)

    win = sui.MainWindow(_MinimalAPI(), settings={})
    qtbot.addWidget(win)

    ini_path = Path(win.preferences_store.fileName())
    assert win.preferences_store.format() == sui.QSettings.IniFormat
    assert ini_path.suffix.lower() == ".ini"
    assert ini_path.stem == sui.SETTINGS_APP_NAME
    assert win._settings is win.preferences_store


def test_tools_menu_edit_ini_file_opens_settings_file(qtbot, monkeypatch):
    monkeypatch.setattr(sui.MainWindow, "_start_worker", lambda self: None)

    opened = []
    monkeypatch.setattr(sui, "_open_path", lambda path: opened.append(path))

    win = sui.MainWindow(_MinimalAPI(), settings={})
    qtbot.addWidget(win)
    ini_path = Path(win.preferences_store.fileName())

    tools_action = next(act for act in win.menuBar().actions() if act.text() == "&Tools")
    edit_action = next(act for act in tools_action.menu().actions() if act.text() == "Edit .ini file")
    edit_action.trigger()

    assert ini_path.exists()
    assert opened == [str(ini_path)]
