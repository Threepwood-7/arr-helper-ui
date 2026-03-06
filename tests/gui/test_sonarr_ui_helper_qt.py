from pathlib import Path

from PySide6.QtGui import QCloseEvent, QStandardItem
from PySide6.QtWidgets import QMessageBox

from arr_helper.sonarr_ui import main_window as sui


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


class _ManualSearchAPI(_MinimalAPI):
    def __init__(self):
        self.calls = []

    def get_release(self, episode_id):
        self.calls.append(("episode", episode_id))
        return []

    def get_release_by_series(self, series_id):
        self.calls.append(("series", series_id))
        return []

    def get_release_by_season(self, series_id, season_number):
        self.calls.append(("season", series_id, season_number))
        return []


def test_manual_search_season_queries_series_season_endpoint(qtbot, monkeypatch):
    monkeypatch.setattr(sui.MainWindow, "_start_worker", lambda self: None)
    monkeypatch.setattr(QMessageBox, "information", lambda *a, **k: None)
    api = _ManualSearchAPI()
    win = sui.MainWindow(api, settings={})
    qtbot.addWidget(win)

    monkeypatch.setattr(
        win,
        "_run_api_action",
        lambda _start, action, on_success, _err: (on_success(action()), True)[1],
    )

    item = QStandardItem("Season 1")
    item.setData("season", sui.ROLE_NODE_TYPE)
    item.setData(77, sui.ROLE_SERIES_ID)
    item.setData(1, sui.ROLE_SEASON_NUM)
    win._ctx_manual_search(item, "season")

    assert api.calls == [("season", 77, 1)]


def test_manual_search_series_queries_series_endpoint(qtbot, monkeypatch):
    monkeypatch.setattr(sui.MainWindow, "_start_worker", lambda self: None)
    monkeypatch.setattr(QMessageBox, "information", lambda *a, **k: None)
    api = _ManualSearchAPI()
    win = sui.MainWindow(api, settings={})
    qtbot.addWidget(win)

    monkeypatch.setattr(
        win,
        "_run_api_action",
        lambda _start, action, on_success, _err: (on_success(action()), True)[1],
    )

    item = QStandardItem("My Series")
    item.setData("series", sui.ROLE_NODE_TYPE)
    item.setData(88, sui.ROLE_SERIES_ID)
    win._ctx_manual_search(item, "series")

    assert api.calls == [("series", 88)]


class _DeleteAPI(_MinimalAPI):
    def __init__(self):
        self.deleted_file_ids = []

    def delete_episode_file(self, file_id):
        self.deleted_file_ids.append(file_id)


def test_delete_from_disk_episode_keeps_row_when_file_delete_fails(qtbot, monkeypatch, tmp_path):
    monkeypatch.setattr(sui.MainWindow, "_start_worker", lambda self: None)
    monkeypatch.setattr(QMessageBox, "question", lambda *a, **k: QMessageBox.Yes)
    monkeypatch.setattr(QMessageBox, "warning", lambda *a, **k: None)
    api = _DeleteAPI()
    win = sui.MainWindow(api, settings={})
    qtbot.addWidget(win)

    monkeypatch.setattr(
        win,
        "_run_api_action",
        lambda _start, action, on_success, _err: (on_success(action()), True)[1],
    )

    file_path = tmp_path / "locked.mkv"
    file_path.write_bytes(b"x")

    series_item = QStandardItem("Series")
    series_item.setData("series", sui.ROLE_NODE_TYPE)
    season_item = QStandardItem("Season 1")
    season_item.setData("season", sui.ROLE_NODE_TYPE)
    ep_item = QStandardItem("E01")
    ep_item.setData("episode", sui.ROLE_NODE_TYPE)
    ep_item.setData(123, sui.ROLE_FILE_ID)
    ep_item.setData(str(file_path), sui.ROLE_FILE_PATH)

    def _make_row(item):
        row = [QStandardItem("") for _ in range(len(win._columns))]
        row[0] = item
        return row

    season_item.appendRow(_make_row(ep_item))
    series_item.appendRow(_make_row(season_item))
    win.model.appendRow(_make_row(series_item))

    def _fail_remove(_path):
        raise PermissionError("locked")

    monkeypatch.setattr(sui.os, "remove", _fail_remove)

    win._ctx_delete_from_disk(ep_item, "episode")

    assert season_item.rowCount() == 1
    assert api.deleted_file_ids == [123]
    assert "disk delete failed" in win.status_label.text().lower()
