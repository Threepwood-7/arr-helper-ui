"""Main Qt window for Sonarr UI helper."""

from __future__ import annotations

import logging
import os
import shutil
from pathlib import Path
from typing import TYPE_CHECKING, cast

from PySide6.QtCore import QModelIndex, QPoint, Qt
from PySide6.QtGui import (
    QAction,
    QCloseEvent,
    QColor,
    QKeySequence,
    QShortcut,
    QStandardItem,
    QStandardItemModel,
)
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMainWindow,
    QMenu,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QStatusBar,
    QTreeView,
    QVBoxLayout,
    QWidget,
)
from threep_commons.desktop import open_path_in_default_app
from threep_commons.settings import QSettingsValueStore

from ..constants import APP_IDENTITY
from .dialogs.add_show import AddShowDialog
from .dialogs.manual_search import ManualSearchDialog
from .helpers import fmt_size
from .probe_cache import clear_probe_cache
from .roles import (
    ROLE_EPISODE_DATA,
    ROLE_FILE_ID,
    ROLE_FILE_PATH,
    ROLE_IS_MISSING,
    ROLE_NODE_TYPE,
    ROLE_SEASON_NUM,
    ROLE_SEASON_PATH,
    ROLE_SERIES_ID,
    ROLE_SERIES_PATH,
)
from .workers import ApiActionWorker, LoadWorker

if TYPE_CHECKING:
    from collections.abc import Callable

    from ..media_checker.config import ConfigMap
    from .api import JsonDict, SonarrAPI
    from .workers import EpisodeEntry, LoadedSeriesEntry

_CONTEXT_MENU_POLICY = Qt.ContextMenuPolicy.CustomContextMenu
_KEY_DELETE = Qt.Key.Key_Delete
_KEY_RETURN = Qt.Key.Key_Return
_WAIT_CURSOR = Qt.CursorShape.WaitCursor
_MSG_YES = QMessageBox.StandardButton.Yes
_MSG_NO = QMessageBox.StandardButton.No
_DIALOG_ACCEPTED = QDialog.DialogCode.Accepted

logger = logging.getLogger(__name__)


class MainWindow(QMainWindow):
    """Present Sonarr series data and batch actions in a tree-driven UI."""

    def _configure_window(self) -> None:
        """Apply top-level window settings and initialize preference stores."""
        self.preferences_store = QSettingsValueStore.from_identity(APP_IDENTITY)
        self._settings = self.preferences_store
        self._ui_settings = QSettingsValueStore(
            self.preferences_store.qsettings,
            namespace="ui/sonarr_ui",
        )
        self.setWindowTitle("Sonarr UI Helper")
        self.resize(1600, 800)
        self._last_worker_error = ""
        self._worker_warning_count = 0

    def _build_toolbar(self, layout: QVBoxLayout) -> None:
        """Build the primary toolbar above the series tree."""
        toolbar = QHBoxLayout()
        btn_expand_all = QPushButton("Expand &All")
        _ = btn_expand_all.clicked.connect(self._expand_all)
        btn_expand_series = QPushButton("Expand &Series")
        _ = btn_expand_series.clicked.connect(self._expand_series)
        btn_collapse_seasons = QPushButton("Collapse S&easons")
        _ = btn_collapse_seasons.clicked.connect(self._collapse_all_seasons)
        btn_collapse_series = QPushButton("&Collapse Series")
        _ = btn_collapse_series.clicked.connect(self._collapse_all_series)
        btn_add_show = QPushButton("A&dd Show")
        _ = btn_add_show.clicked.connect(self._add_show)
        toolbar.addWidget(btn_expand_all)
        toolbar.addWidget(btn_expand_series)
        toolbar.addWidget(btn_collapse_seasons)
        toolbar.addWidget(btn_collapse_series)
        btn_refresh = QPushButton("&Refresh")
        _ = btn_refresh.clicked.connect(self._refresh)
        self.chk_show_missing = QCheckBox("Show &Missing")
        self.chk_show_missing.setChecked(False)
        _ = self.chk_show_missing.toggled.connect(self._toggle_missing)
        toolbar.addWidget(self.chk_show_missing)
        toolbar.addStretch()
        toolbar.addWidget(btn_add_show)
        toolbar.addWidget(btn_refresh)
        layout.addLayout(toolbar)

    def _build_tree_area(self, layout: QVBoxLayout) -> None:
        """Create the tree view and its backing model."""
        self.tree = QTreeView()
        self.tree.setAlternatingRowColors(True)
        self.tree.setUniformRowHeights(True)
        self.tree.setAnimated(False)
        self.tree.setContextMenuPolicy(_CONTEXT_MENU_POLICY)
        _ = self.tree.customContextMenuRequested.connect(self._on_context_menu)
        _ = self.tree.doubleClicked.connect(self._on_double_click)

        self.model = QStandardItemModel()
        self._columns = [
            "Name",
            "Size",
            "Mon",
            "Quality Profile",
            "Resolution",
            "V.Bitrate",
            "V.Codec",
            "HDR",
            "A.Codec",
            "A.Bitrate",
            "Audio Lang",
            "Sub Lang",
        ]
        self.model.setHorizontalHeaderLabels(self._columns)
        self.tree.setModel(self.model)
        layout.addWidget(self.tree)

    def _build_status_bar(self) -> None:
        """Create the window status bar and progress indicator."""
        self.status_label = QLabel("Loading...")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setMaximumWidth(200)
        status_bar = QStatusBar()
        status_bar.addWidget(self.status_label, 1)
        status_bar.addPermanentWidget(self.progress_bar)
        self.setStatusBar(status_bar)

    def _configure_shortcuts(self) -> None:
        """Register keyboard shortcuts for the tree view."""
        del_shortcut = QShortcut(QKeySequence(_KEY_DELETE), self.tree)
        _ = del_shortcut.activated.connect(self._on_delete)

        enter_shortcut = QShortcut(QKeySequence(_KEY_RETURN), self.tree)
        _ = enter_shortcut.activated.connect(self._on_enter)

    def _load_quality_profiles(self) -> None:
        """Cache Sonarr quality profiles for quick label lookup."""
        self._quality_profiles = []
        self._qp_map = {}
        try:
            self._quality_profiles = self.api.get_quality_profiles()
            self._qp_map = {
                profile_id: profile_name
                for profile in self._quality_profiles
                if isinstance((profile_id := profile.get("id")), int)
                and isinstance((profile_name := profile.get("name")), str)
            }
        except Exception:
            logger.debug("Failed to cache Sonarr quality profiles", exc_info=True)

    def _build_main_layout(self) -> None:
        """Assemble the central widget layout for the main window."""
        central = QWidget()
        layout = QVBoxLayout(central)
        layout.setContentsMargins(4, 4, 4, 4)
        self._build_toolbar(layout)
        self._build_tree_area(layout)
        self.setCentralWidget(central)

    def __init__(
        self,
        api: SonarrAPI,
        loader_api: SonarrAPI | None = None,
        settings: ConfigMap | None = None,
    ) -> None:
        super().__init__()
        self.api = api
        self.loader_api = loader_api or api
        self.cfg: ConfigMap = settings or {}
        self._configure_window()

        # ── Menu bar ──────────────────────────────────────────
        self._build_menu_bar()

        self._build_main_layout()
        self._build_status_bar()
        self._configure_shortcuts()
        self._load_quality_profiles()

        # start loading
        self.worker: LoadWorker | None = None
        self.action_worker: ApiActionWorker | None = None
        self._start_worker()

    # ── menu bar ───────────────────────────────────────────────

    @staticmethod
    def _ui_key(name: str) -> str:
        return f"ui/sonarr_ui/{name}"

    @staticmethod
    def _item_role_str(item: QStandardItem, role: int) -> str:
        value = item.data(role)
        return str(value) if isinstance(value, str) else ""

    @staticmethod
    def _item_role_int(item: QStandardItem, role: int) -> int | None:
        value = item.data(role)
        return value if isinstance(value, int) else None

    @staticmethod
    def _item_role_bool(item: QStandardItem, role: int) -> bool:
        return bool(item.data(role))

    @staticmethod
    def _as_json_dict(value: object) -> JsonDict | None:
        if not isinstance(value, dict):
            return None
        return dict(cast("JsonDict", value))

    @staticmethod
    def _as_json_list(value: object) -> list[JsonDict]:
        if not isinstance(value, list):
            return []
        payloads: list[JsonDict] = []
        entries = cast("list[object]", value)
        for entry in entries:
            payload = MainWindow._as_json_dict(entry)
            if payload is not None:
                payloads.append(payload)
        return payloads

    @staticmethod
    def _item_role_payloads(item: QStandardItem, role: int) -> list[JsonDict]:
        return MainWindow._as_json_list(item.data(role))

    @staticmethod
    def _result_int(result: object, key: str) -> int:
        result_map = MainWindow._as_json_dict(result)
        if result_map is None:
            return 0
        value = result_map.get(key, 0)
        return value if isinstance(value, int) else 0

    @staticmethod
    def _result_bool(result: object, key: str) -> bool:
        result_map = MainWindow._as_json_dict(result)
        if result_map is None:
            return False
        return bool(result_map.get(key, False))

    @staticmethod
    def _result_str(result: object, key: str) -> str:
        result_map = MainWindow._as_json_dict(result)
        if result_map is None:
            return ""
        value = result_map.get(key, "")
        return value if isinstance(value, str) else ""

    def _build_menu_bar(self) -> None:
        mb = self.menuBar()

        # File menu
        file_menu = mb.addMenu("&File")
        act = file_menu.addAction("&Add Show")
        act.setShortcut(QKeySequence("Ctrl+N"))
        _ = act.triggered.connect(self._add_show)
        file_menu.addSeparator()
        act = file_menu.addAction("E&xit")
        act.setShortcuts([QKeySequence("Ctrl+Q"), QKeySequence("Alt+X")])
        _ = act.triggered.connect(self.close)

        # View menu
        view_menu = mb.addMenu("&View")
        act = view_menu.addAction("&Refresh")
        act.setShortcut(QKeySequence("F5"))
        _ = act.triggered.connect(self._refresh)
        act = view_menu.addAction("&Clear Cache && Refresh")
        act.setShortcut(QKeySequence("Ctrl+F5"))
        _ = act.triggered.connect(self._clear_cache_and_refresh)
        view_menu.addSeparator()
        act = view_menu.addAction("E&xpand All")
        act.setShortcut(QKeySequence("Ctrl+E"))
        _ = act.triggered.connect(self._expand_all)
        act = view_menu.addAction("Expand &Series")
        act.setShortcut(QKeySequence("Ctrl+Shift+E"))
        _ = act.triggered.connect(self._expand_series)
        act = view_menu.addAction("Collapse S&easons")
        act.setShortcut(QKeySequence("Ctrl+W"))
        _ = act.triggered.connect(self._collapse_all_seasons)
        act = view_menu.addAction("Co&llapse All")
        act.setShortcut(QKeySequence("Ctrl+Shift+W"))
        _ = act.triggered.connect(self._collapse_all_series)
        view_menu.addSeparator()
        self.act_show_missing = QAction("Show &Missing", self)
        self.act_show_missing.setCheckable(True)
        self.act_show_missing.setChecked(False)
        self.act_show_missing.setShortcut(QKeySequence("Ctrl+M"))
        _ = self.act_show_missing.toggled.connect(self._toggle_missing_from_menu)
        view_menu.addAction(self.act_show_missing)
        act = view_menu.addAction("&Fit Columns")
        _ = act.triggered.connect(self._fit_columns)
        act = view_menu.addAction("Reset &View")
        act.setShortcut(QKeySequence("Ctrl+Shift+R"))
        _ = act.triggered.connect(self._reset_view_settings)

        # Actions menu
        actions_menu = mb.addMenu("&Actions")
        act = actions_menu.addAction("&Monitor")
        act.setShortcut(QKeySequence("M"))
        _ = act.triggered.connect(lambda: self._ctx_on_selected("monitor"))
        act = actions_menu.addAction("&Auto Search")
        act.setShortcut(QKeySequence("S"))
        _ = act.triggered.connect(lambda: self._ctx_on_selected("auto_search"))
        act = actions_menu.addAction("Ma&nual Search")
        act.setShortcut(QKeySequence("N"))
        _ = act.triggered.connect(lambda: self._ctx_on_selected("manual_search"))
        actions_menu.addSeparator()
        act = actions_menu.addAction("Change &Quality Profile")
        act.setShortcut(QKeySequence("Q"))
        _ = act.triggered.connect(
            lambda: self._ctx_on_selected("change_quality_profile")
        )
        actions_menu.addSeparator()
        act = actions_menu.addAction("&Unmonitor")
        act.setShortcut(QKeySequence("U"))
        _ = act.triggered.connect(lambda: self._ctx_on_selected("unmonitor"))
        act = actions_menu.addAction("&Delete from Disk")
        act.setShortcut(QKeySequence("D"))
        _ = act.triggered.connect(lambda: self._ctx_on_selected("delete_from_disk"))
        act = actions_menu.addAction("Unmonitor && De&lete")
        act.setShortcut(QKeySequence("Ctrl+Delete"))
        _ = act.triggered.connect(lambda: self._ctx_on_selected("unmonitor_delete"))
        actions_menu.addSeparator()
        act = actions_menu.addAction("&Open in Explorer")
        act.setShortcut(QKeySequence("O"))
        _ = act.triggered.connect(self._on_enter)

        # Tools menu
        tools_menu = mb.addMenu("&Tools")
        act = tools_menu.addAction("Edit &.ini File")
        _ = act.triggered.connect(self._edit_ini_file)

        # Help menu
        help_menu = mb.addMenu("&Help")
        act = help_menu.addAction("&Help")
        act.setShortcut(QKeySequence("F1"))
        _ = act.triggered.connect(self._show_help)

    def _toggle_missing_from_menu(self, checked: bool) -> None:
        """Sync the menu checkbox with the toolbar checkbox."""
        self.chk_show_missing.setChecked(checked)

    def _reset_view_settings(self) -> None:
        reply = QMessageBox.question(
            self,
            "Reset View",
            "Reset all saved UI view settings to defaults?\n"
            "This clears saved column widths, splitter positions, and other "
            "stored view state.",
            _MSG_YES | _MSG_NO,
            _MSG_NO,
        )
        if reply != _MSG_YES:
            return

        self._ui_settings.clear_all()
        self._settings.sync()

        # Re-apply in-memory defaults immediately.
        self.chk_show_missing.setChecked(False)
        self._apply_default_column_widths()
        self.status_label.setText("View settings reset to defaults")

    def _edit_ini_file(self) -> None:
        self._settings.sync()
        ini_path = Path(self.preferences_store.file_name())
        try:
            ini_path.parent.mkdir(parents=True, exist_ok=True)
            ini_path.touch(exist_ok=True)
            open_path_in_default_app(str(ini_path))
            self.status_label.setText(f"Opened settings file: {ini_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open settings file:\n{e}")
            self.status_label.setText("Failed to open settings file")

    def _ctx_on_selected(self, action_name: str) -> None:
        """Dispatch a context-menu action on the currently selected tree item."""
        if not self._ensure_action_idle():
            return
        item = self._current_item()
        if not item:
            self.status_label.setText("No item selected")
            return
        node_type = self._item_role_str(item, ROLE_NODE_TYPE)
        if action_name == "monitor":
            self._ctx_monitor(item, node_type)
        elif action_name == "auto_search":
            self._ctx_auto_search(item, node_type)
        elif action_name == "manual_search":
            self._ctx_manual_search(item, node_type)
        elif action_name == "change_quality_profile":
            self._ctx_change_quality_profile(item)
        elif action_name == "unmonitor":
            self._ctx_unmonitor(item, node_type)
        elif action_name == "delete_from_disk":
            self._ctx_delete_from_disk(item, node_type)
        elif action_name == "unmonitor_delete":
            self._ctx_unmonitor_delete(item, node_type)

    def _show_help(self) -> None:
        help_text = (
            "<h2>Keyboard Shortcuts</h2>"
            "<table cellpadding='4' cellspacing='0'>"
            "<tr><td><b>General</b></td><td></td></tr>"
            "<tr><td><code>F1</code></td><td>Show this help</td></tr>"
            "<tr><td><code>F5</code></td><td>Refresh data from Sonarr</td></tr>"
            "<tr><td><code>Ctrl+F5</code></td><td>Clear cache &amp; refresh</td></tr>"
            "<tr><td><code>Ctrl+N</code></td><td>Add a new show</td></tr>"
            "<tr><td><code>Ctrl+Q</code></td><td>Quit</td></tr>"
            "<tr><td><code>Alt+X</code></td><td>Quit</td></tr>"
            "<tr><td></td><td></td></tr>"
            "<tr><td><b>Navigation</b></td><td></td></tr>"
            "<tr><td><code>Enter</code></td><td>Open file/folder in Explorer</td></tr>"
            "<tr><td><code>O</code></td><td>Open in Explorer</td></tr>"
            "<tr><td><code>Delete</code></td><td>Remove series/season from "
            "Sonarr (deletes files)</td></tr>"
            "<tr><td><code>Double-click</code></td><td>Open file/folder</td></tr>"
            "<tr><td></td><td></td></tr>"
            "<tr><td><b>View</b></td><td></td></tr>"
            "<tr><td><code>Ctrl+E</code></td><td>Expand all (except Specials)</td></tr>"
            "<tr><td><code>Ctrl+Shift+E</code></td><td>Expand series only</td></tr>"
            "<tr><td><code>Ctrl+W</code></td><td>Collapse seasons</td></tr>"
            "<tr><td><code>Ctrl+Shift+W</code></td><td>Collapse all</td></tr>"
            "<tr><td><code>Ctrl+M</code></td><td>Toggle show/hide missing</td></tr>"
            "<tr><td><code>Ctrl+Shift+R</code></td><td>Reset saved view state</td></tr>"
            "<tr><td></td><td></td></tr>"
            "<tr><td><b>Actions (on selected item)</b></td><td></td></tr>"
            "<tr><td><code>M</code></td><td>Monitor</td></tr>"
            "<tr><td><code>S</code></td><td>Auto Search</td></tr>"
            "<tr><td><code>N</code></td><td>Manual Search</td></tr>"
            "<tr><td><code>Q</code></td><td>Change Quality Profile (series "
            "only)</td></tr>"
            "<tr><td><code>U</code></td><td>Unmonitor</td></tr>"
            "<tr><td><code>D</code></td><td>Delete files from disk (keep in "
            "Sonarr)</td></tr>"
            "<tr><td><code>Ctrl+Delete</code></td><td>Unmonitor &amp; Delete "
            "from disk</td></tr>"
            "<tr><td></td><td></td></tr>"
            "<tr><td><b>Toolbar Mnemonics (Alt+key)</b></td><td></td></tr>"
            "<tr><td><code>Alt+A</code></td><td>Expand All</td></tr>"
            "<tr><td><code>Alt+S</code></td><td>Expand Series</td></tr>"
            "<tr><td><code>Alt+E</code></td><td>Collapse Seasons</td></tr>"
            "<tr><td><code>Alt+C</code></td><td>Collapse Series</td></tr>"
            "<tr><td><code>Alt+M</code></td><td>Show Missing checkbox</td></tr>"
            "<tr><td><code>Alt+D</code></td><td>Add Show</td></tr>"
            "<tr><td><code>Alt+R</code></td><td>Refresh</td></tr>"
            "</table>"
            "<br>"
            "<p>Right-click any item for the context menu with all actions.</p>"
        )
        QMessageBox.information(self, "Keyboard Shortcuts", help_text)

    # ── populate tree ───────────────────────────────────────────

    def _on_progress(self, text: str) -> None:
        self.status_label.setText(text)

    def _on_worker_progress(self, worker: LoadWorker, text: str) -> None:
        if worker is not self.worker:
            return
        if text.startswith("Error:"):
            self._last_worker_error = text
        elif text.startswith("Warning:"):
            self._worker_warning_count += 1
        self._on_progress(text)

    def _on_worker_series_ready(
        self,
        worker: LoadWorker,
        series_list: list[LoadedSeriesEntry],
    ) -> None:
        if worker is not self.worker:
            return
        self._on_data_loaded(series_list)

    def _start_worker(self) -> None:
        self._last_worker_error = ""
        self._worker_warning_count = 0
        worker = LoadWorker(self.loader_api)

        def _handle_progress(text: object, w: LoadWorker = worker) -> None:
            self._on_worker_progress(w, str(text))

        def _handle_series_ready(data: object, w: LoadWorker = worker) -> None:
            self._on_worker_series_ready(
                w,
                cast("list[LoadedSeriesEntry]", data),
            )

        _ = worker.progress.connect(_handle_progress)
        _ = worker.series_ready.connect(_handle_series_ready)
        self.worker = worker
        worker.start()

    def _stop_worker(self, timeout_ms: int = 5000) -> bool:
        worker = self.worker
        if not worker or not worker.isRunning():
            return True
        worker.requestInterruption()
        return worker.wait(timeout_ms)

    def _action_in_progress(self) -> bool:
        return bool(self.action_worker and self.action_worker.isRunning())

    def _ensure_action_idle(
        self, status_text: str = "Another action is still running"
    ) -> bool:
        if self._action_in_progress():
            self.status_label.setText(status_text)
            return False
        return True

    def _stop_action_worker(self, timeout_ms: int = 5000) -> bool:
        worker = self.action_worker
        if not worker or not worker.isRunning():
            return True
        return worker.wait(timeout_ms)

    def _run_api_action(
        self,
        start_text: str,
        action: Callable[[], object],
        on_success: Callable[[object], None],
        error_text: str,
    ) -> bool:
        if not self._ensure_action_idle():
            return False

        worker = ApiActionWorker(action)
        self.action_worker = worker
        self.status_label.setText(start_text)
        QApplication.setOverrideCursor(_WAIT_CURSOR)

        def _on_done(
            result: object, error: object, w: ApiActionWorker = worker
        ) -> None:
            if self.action_worker is not w:
                return
            self.action_worker = None
            QApplication.restoreOverrideCursor()
            try:
                if error is not None:
                    QMessageBox.critical(self, "Error", f"{error_text}:\n{error}")
                    self.status_label.setText(error_text)
                    return
                on_success(result)
            finally:
                w.deleteLater()

        _ = worker.finished_action.connect(_on_done)
        worker.start()
        return True

    @staticmethod
    def _make_row(cols: int) -> list[QStandardItem]:
        """Create a list of empty, non-editable QStandardItems."""
        row: list[QStandardItem] = []
        for _ in range(cols):
            item = QStandardItem("")
            item.setEditable(False)
            row.append(item)
        return row

    def _highlight_row(self, row: list[QStandardItem], sub_langs: str) -> None:
        """Apply light red background if configured sub language is missing.

        highlight_missing_subs is a label that maps to english_language_codes,
        e.g. highlight_missing_subs = "english" with
        english_language_codes = ["eng", "en", "english"] means any of those
        codes count as a match.
        """
        hl = str(self.cfg.get("highlight_missing_subs", "")).strip().lower()
        if not hl:
            return
        # expand via the language codes array
        raw_codes = self.cfg.get("english_language_codes", [hl])
        code_values = (
            cast("list[object]", raw_codes) if isinstance(raw_codes, list) else []
        )
        codes = (
            [str(code).lower() for code in code_values if isinstance(code, str)]
            if isinstance(raw_codes, list)
            else [hl]
        )
        if not codes:
            codes = [hl]
        langs = {lang.strip().lower() for lang in sub_langs.split(",") if lang.strip()}
        if not langs.intersection(codes):
            bg = QColor(255, 200, 200)
            for cell in row:
                cell.setBackground(bg)

    def _apply_default_column_widths(self):
        # Default layout: title column wide, data columns sized to content.
        self.tree.setColumnWidth(0, 600)
        for col in range(1, len(self._columns)):
            self.tree.resizeColumnToContents(col)

    def _fit_columns(self):
        """Resize all tree columns to fit their current contents."""
        for col in range(len(self._columns)):
            self.tree.resizeColumnToContents(col)
        self.status_label.setText("Columns fitted to contents")

    def _season_label(
        self,
        season_num: int,
        downloaded_eps: list[EpisodeEntry],
        missing_eps: list[JsonDict],
    ) -> str:
        """Build one season label with download and missing counts."""

        season_label = f"Season {season_num}" if season_num > 0 else "Specials"
        has_downloaded = len(downloaded_eps) > 0
        has_missing = len(missing_eps) > 0
        if has_downloaded and has_missing:
            season_label += (
                f" ({len(downloaded_eps)} downloaded, {len(missing_eps)} missing)"
            )
        elif not has_downloaded:
            season_label += f" ({len(missing_eps)} missing)"
        return season_label

    def _season_monitored_state(
        self,
        series_data: JsonDict,
        season_num: int,
    ) -> bool:
        """Read the monitored flag for one season from the series payload."""

        season_info_obj = series_data.get("seasons", [])
        season_info_list = (
            cast("list[JsonDict]", season_info_obj)
            if isinstance(season_info_obj, list)
            else []
        )
        for season_info in season_info_list:
            if season_info.get("seasonNumber") == season_num:
                return bool(season_info.get("monitored", True))
        return True

    def _append_downloaded_episode_rows(
        self,
        season_item: QStandardItem,
        *,
        downloaded_eps: list[EpisodeEntry],
        series_id: int,
        season_num: int,
    ) -> None:
        """Append downloaded episode rows for one season."""

        for episode in downloaded_eps:
            ep_label = f"E{episode['episode_number']:02d} - {episode['file_name']}"
            ep_item = QStandardItem(ep_label)
            ep_item.setEditable(False)
            ep_item.setData("episode", ROLE_NODE_TYPE)
            ep_item.setData(episode["file_path"], ROLE_FILE_PATH)
            ep_item.setData(episode["file_id"], ROLE_FILE_ID)
            ep_item.setData(episode["episode_data"], ROLE_EPISODE_DATA)
            ep_item.setData(series_id, ROLE_SERIES_ID)
            ep_item.setData(season_num, ROLE_SEASON_NUM)
            ep_item.setData(False, ROLE_IS_MISSING)

            ep_row = self._make_row(len(self._columns))
            ep_row[0] = ep_item
            ep_row[1].setText(fmt_size(episode["size_bytes"]))
            ep_monitored = any(
                bool(payload.get("monitored", False))
                for payload in episode["episode_data"]
            )
            ep_row[2].setText("Y" if ep_monitored else "N")
            ep_row[4].setText(episode["video_resolution"])
            ep_row[5].setText(episode["video_bitrate"])
            ep_row[6].setText(episode["video_codec"])
            ep_row[7].setText(episode["hdr"])
            ep_row[8].setText(episode["audio_codec"])
            ep_row[9].setText(episode["audio_bitrate"])
            ep_row[10].setText(episode["audio_langs"])
            ep_row[11].setText(episode["sub_langs"])
            self._highlight_row(ep_row, episode["sub_langs"])
            season_item.appendRow(ep_row)

    def _append_missing_episode_rows(
        self,
        season_item: QStandardItem,
        *,
        missing_eps: list[JsonDict],
        series_id: int,
        season_num: int,
        missing_color: QColor,
    ) -> None:
        """Append missing episode rows for one season."""

        for missing_episode in missing_eps:
            ep_num_obj = missing_episode.get("episodeNumber", 0)
            ep_num = ep_num_obj if isinstance(ep_num_obj, int) else 0
            ep_title = str(missing_episode.get("title", ""))
            monitored = bool(missing_episode.get("monitored", False))
            ep_item = QStandardItem(f"E{ep_num:02d} - {ep_title}")
            ep_item.setEditable(False)
            ep_item.setForeground(missing_color)
            ep_item.setData("episode", ROLE_NODE_TYPE)
            ep_item.setData(series_id, ROLE_SERIES_ID)
            ep_item.setData(season_num, ROLE_SEASON_NUM)
            ep_item.setData(True, ROLE_IS_MISSING)
            ep_item.setData([missing_episode], ROLE_EPISODE_DATA)

            ep_row = self._make_row(len(self._columns))
            ep_row[0] = ep_item
            ep_row[2].setText("Y" if monitored else "N")
            season_item.appendRow(ep_row)

    def _append_loaded_series_row(
        self,
        series_entry: LoadedSeriesEntry,
        *,
        missing_color: QColor,
    ) -> None:
        """Append one loaded series tree branch to the model."""

        series_data = series_entry["series_data"]
        series_label = (
            f"{series_entry['title']}, {series_entry['year']}"
            if series_entry["year"]
            else series_entry["title"]
        )
        series_item = QStandardItem(series_label)
        series_item.setEditable(False)
        series_item.setData("series", ROLE_NODE_TYPE)
        series_item.setData(series_entry["series_id"], ROLE_SERIES_ID)
        series_item.setData(series_entry["path"], ROLE_SERIES_PATH)
        font = series_item.font()
        font.setBold(True)
        series_item.setFont(font)

        series_row = self._make_row(len(self._columns))
        series_row[0] = series_item
        series_row[1].setText(fmt_size(series_entry["total_size"]))
        series_row[2].setText("Y" if bool(series_data.get("monitored", False)) else "N")
        qp_id_obj = series_data.get("qualityProfileId", 0)
        qp_id = qp_id_obj if isinstance(qp_id_obj, int) else 0
        series_row[3].setText(self._qp_map.get(qp_id, str(qp_id)))

        for season_num in series_entry["all_season_nums"]:
            downloaded_eps = series_entry["seasons"].get(season_num, [])
            missing_eps = series_entry["missing_seasons"].get(season_num, [])
            season_size = sum(entry["size_bytes"] for entry in downloaded_eps)
            season_path = ""
            if downloaded_eps and downloaded_eps[0]["file_path"]:
                season_path = str(Path(downloaded_eps[0]["file_path"]).parent)

            has_downloaded = len(downloaded_eps) > 0
            season_item = QStandardItem(
                self._season_label(season_num, downloaded_eps, missing_eps)
            )
            season_item.setEditable(False)
            season_item.setData("season", ROLE_NODE_TYPE)
            season_item.setData(series_entry["series_id"], ROLE_SERIES_ID)
            season_item.setData(season_num, ROLE_SEASON_NUM)
            season_item.setData(season_path, ROLE_SEASON_PATH)
            season_item.setData(not has_downloaded, ROLE_IS_MISSING)
            if not has_downloaded:
                season_item.setForeground(missing_color)

            season_row = self._make_row(len(self._columns))
            season_row[0] = season_item
            if season_size > 0:
                season_row[1].setText(fmt_size(season_size))
            season_row[2].setText(
                "Y" if self._season_monitored_state(series_data, season_num) else "N"
            )
            self._append_downloaded_episode_rows(
                season_item,
                downloaded_eps=downloaded_eps,
                series_id=series_entry["series_id"],
                season_num=season_num,
            )
            self._append_missing_episode_rows(
                season_item,
                missing_eps=missing_eps,
                series_id=series_entry["series_id"],
                season_num=season_num,
                missing_color=missing_color,
            )
            series_item.appendRow(season_row)

        self.model.appendRow(series_row)

    def _expand_loaded_tree(self) -> None:
        """Expand loaded rows down to season level while collapsing Specials."""

        root = self.model.invisibleRootItem()
        for row in range(root.rowCount()):
            series_idx = self.model.index(row, 0)
            self.tree.expand(series_idx)
            series_item = self.model.itemFromIndex(series_idx)
            for season_row in range(series_item.rowCount()):
                season_idx = self.model.index(season_row, 0, series_idx)
                season_item = self.model.itemFromIndex(season_idx)
                if season_item and season_item.data(ROLE_SEASON_NUM) == 0:
                    self.tree.collapse(season_idx)
                else:
                    self.tree.expand(season_idx)

    def _restore_loaded_column_widths(self) -> None:
        """Restore persisted column widths or fall back to defaults."""

        header = self.tree.header()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        saved = self._settings.value(self._ui_key("column_widths"))
        if saved and len(saved) == len(self._columns):
            for col, width in enumerate(saved):
                self.tree.setColumnWidth(col, int(width))
            return
        self._apply_default_column_widths()

    def _loaded_status_text(self, total_series: int) -> str:
        """Build the status-bar text for a successful load result."""

        status_text = f"{total_series} series loaded"
        if self._worker_warning_count:
            status_text += f" ({self._worker_warning_count} warnings)"
        return status_text

    def _on_data_loaded(self, series_list: list[LoadedSeriesEntry]) -> None:
        self.progress_bar.hide()
        self.model.removeRows(0, self.model.rowCount())

        if not series_list:
            if self._last_worker_error:
                self.status_label.setText(self._last_worker_error)
            elif self._worker_warning_count:
                self.status_label.setText(
                    f"No series loaded ({self._worker_warning_count} warnings)"
                )
            else:
                self.status_label.setText("No series with downloaded episodes found.")
            return

        total_series = len(series_list)
        missing_color = QColor(128, 128, 128)  # grey for missing episodes

        for series_entry in series_list:
            self._append_loaded_series_row(series_entry, missing_color=missing_color)

        self._expand_loaded_tree()
        self._restore_loaded_column_widths()
        self.status_label.setText(self._loaded_status_text(total_series))

        # apply initial missing visibility
        self._apply_missing_visibility(self.chk_show_missing.isChecked())

    # ── expand / collapse helpers ────────────────────────────────

    def _expand_all(self):
        """Expand everything except Specials (season 0)."""
        root = self.model.invisibleRootItem()
        for row in range(root.rowCount()):
            series_idx = self.model.index(row, 0)
            self.tree.expand(series_idx)
            series_item = self.model.itemFromIndex(series_idx)
            for s_row in range(series_item.rowCount()):
                season_idx = self.model.index(s_row, 0, series_idx)
                season_item = self.model.itemFromIndex(season_idx)
                if season_item and season_item.data(ROLE_SEASON_NUM) == 0:
                    self.tree.collapse(season_idx)
                else:
                    self.tree.expand(season_idx)

    def _expand_series(self):
        root = self.model.invisibleRootItem()
        for row in range(root.rowCount()):
            series_idx = self.model.index(row, 0)
            self.tree.expand(series_idx)
            series_item = self.model.itemFromIndex(series_idx)
            for s_row in range(series_item.rowCount()):
                season_idx = self.model.index(s_row, 0, series_idx)
                self.tree.collapse(season_idx)

    def _collapse_all_series(self):
        self.tree.collapseAll()

    def _collapse_all_seasons(self):
        root = self.model.invisibleRootItem()
        for row in range(root.rowCount()):
            series_idx = self.model.index(row, 0)
            self.tree.expand(series_idx)
            series_item = self.model.itemFromIndex(series_idx)
            for s_row in range(series_item.rowCount()):
                season_idx = self.model.index(s_row, 0, series_idx)
                self.tree.collapse(season_idx)

    # ── show/hide missing ─────────────────────────────────────

    def _toggle_missing(self, show: bool):
        # sync menu checkbox without re-triggering
        self.act_show_missing.blockSignals(True)
        self.act_show_missing.setChecked(show)
        self.act_show_missing.blockSignals(False)
        self._apply_missing_visibility(show)

    def _apply_missing_visibility(self, show: bool):
        """Show or hide rows marked as missing (seasons and episodes)."""
        root = self.model.invisibleRootItem()
        for s_row in range(root.rowCount()):
            series_item = root.child(s_row, 0)
            for sn_row in range(series_item.rowCount()):
                season_item = series_item.child(sn_row, 0)
                season_is_missing = season_item.data(ROLE_IS_MISSING)
                season_idx = self.model.indexFromItem(season_item)

                if season_is_missing:
                    self.tree.setRowHidden(sn_row, series_item.index(), not show)
                    continue

                # check individual episodes inside this season
                for ep_row in range(season_item.rowCount()):
                    ep_item = season_item.child(ep_row, 0)
                    if ep_item.data(ROLE_IS_MISSING):
                        self.tree.setRowHidden(ep_row, season_idx, not show)

    # ── context menu ───────────────────────────────────────────

    def _on_context_menu(self, pos: QPoint) -> None:
        if not self._ensure_action_idle():
            return
        index = self.tree.indexAt(pos)
        if not index.isValid():
            return
        if index.column() != 0:
            index = index.siblingAtColumn(0)
        item = self.model.itemFromIndex(index)
        if not item:
            return

        node_type = self._item_role_str(item, ROLE_NODE_TYPE)
        menu = QMenu(self)

        act_monitor = menu.addAction("Monitor")
        act_auto_search = menu.addAction("Auto Search")
        act_manual_search = menu.addAction("Manual Search")
        menu.addSeparator()
        act_change_qp = menu.addAction("Change Quality Profile")
        act_change_qp.setEnabled(node_type == "series")
        menu.addSeparator()
        act_unmonitor = menu.addAction("Unmonitor")
        act_delete_disk = menu.addAction("Delete From Disk")
        act_unmonitor_delete = menu.addAction("Unmonitor && Delete From Disk")

        action = menu.exec(self.tree.viewport().mapToGlobal(pos))
        if not action:
            return

        if action == act_monitor:
            self._ctx_monitor(item, node_type)
        elif action == act_auto_search:
            self._ctx_auto_search(item, node_type)
        elif action == act_manual_search:
            self._ctx_manual_search(item, node_type)
        elif action == act_change_qp:
            self._ctx_change_quality_profile(item)
        elif action == act_unmonitor:
            self._ctx_unmonitor(item, node_type)
        elif action == act_delete_disk:
            self._ctx_delete_from_disk(item, node_type)
        elif action == act_unmonitor_delete:
            self._ctx_unmonitor_delete(item, node_type)

    def _update_mon_column(self, item: QStandardItem, monitored: bool) -> None:
        """Update the 'Mon' column (col 2) for the row containing item."""
        parent = item.parent() or self.model.invisibleRootItem()
        mon_item = parent.child(item.row(), 2)
        if mon_item:
            mon_item.setText("Y" if monitored else "N")

    @staticmethod
    def _clone_episode_payloads(raw_payloads: object) -> list[JsonDict]:
        return MainWindow._as_json_list(raw_payloads)

    def _collect_episode_payloads(self, parent_item: QStandardItem) -> list[JsonDict]:
        payloads: list[JsonDict] = []
        for row in range(parent_item.rowCount()):
            ep_item = parent_item.child(row, 0)
            payloads.extend(self._item_role_payloads(ep_item, ROLE_EPISODE_DATA))
        return payloads

    @staticmethod
    def _delete_tree_path(path: str) -> tuple[bool, str]:
        if not path:
            return True, ""
        if not os.path.exists(path):
            return True, ""
        try:
            shutil.rmtree(path)
        except Exception as e:
            if os.path.exists(path):
                return False, str(e)
        if os.path.exists(path):
            return False, "Directory still exists after delete attempt."
        return True, ""

    @staticmethod
    def _delete_file_path(path: str) -> tuple[bool, str]:
        if not path:
            return True, ""
        if not os.path.exists(path):
            return True, ""
        try:
            os.remove(path)
        except Exception as e:
            if os.path.exists(path):
                return False, str(e)
        if os.path.exists(path):
            return False, "File still exists after delete attempt."
        return True, ""

    def _ctx_monitor_series(
        self,
        item: QStandardItem,
        *,
        label: str,
        series_id: int,
    ) -> None:
        def _action() -> None:
            series_data = self.api.get_series_by_id(series_id)
            series_data["monitored"] = True
            self.api.update_series(series_data)

        def _on_success(_result: object) -> None:
            self._update_mon_column(item, True)
            self.status_label.setText(f"Monitored: {label}")

        self._run_api_action(
            f"Monitoring: {label}...",
            _action,
            _on_success,
            "Failed to monitor",
        )

    def _ctx_monitor_season(
        self,
        item: QStandardItem,
        *,
        label: str,
        series_id: int,
    ) -> None:
        season_num = self._item_role_int(item, ROLE_SEASON_NUM)
        if season_num is None:
            self.status_label.setText("No season number available")
            return

        def _action() -> None:
            series_data = self.api.get_series_by_id(series_id)
            seasons_obj = series_data.get("seasons", [])
            seasons = (
                cast("list[JsonDict]", seasons_obj)
                if isinstance(seasons_obj, list)
                else []
            )
            for season in seasons:
                if season.get("seasonNumber") == season_num:
                    season["monitored"] = True
                    break
            self.api.update_series(series_data)

        def _on_success(_result: object) -> None:
            self._update_mon_column(item, True)
            self.status_label.setText(f"Monitored: {label}")

        self._run_api_action(
            f"Monitoring: {label}...",
            _action,
            _on_success,
            "Failed to monitor",
        )

    def _ctx_monitor_episode(
        self,
        item: QStandardItem,
        *,
        label: str,
    ) -> None:
        ep_payloads = self._item_role_payloads(item, ROLE_EPISODE_DATA)
        if not ep_payloads:
            self.status_label.setText("No episode payload found for monitor action")
            return

        def _action() -> None:
            for episode in ep_payloads:
                episode["monitored"] = True
                self.api.update_episode(episode)

        def _on_success(_result: object) -> None:
            self._update_mon_column(item, True)
            self.status_label.setText(f"Monitored: {label}")

        self._run_api_action(
            f"Monitoring: {label}...",
            _action,
            _on_success,
            "Failed to monitor",
        )

    def _ctx_monitor(self, item: QStandardItem, node_type: str):
        series_id = self._item_role_int(item, ROLE_SERIES_ID)
        label = item.text()
        if series_id is None:
            self.status_label.setText("No series ID available")
            return
        handlers = {
            "series": self._ctx_monitor_series,
            "season": self._ctx_monitor_season,
        }
        handler = handlers.get(node_type)
        if handler is not None:
            handler(item, label=label, series_id=series_id)
            return
        if node_type == "episode":
            self._ctx_monitor_episode(item, label=label)

    def _ctx_unmonitor_series(
        self,
        item: QStandardItem,
        *,
        label: str,
        series_id: int,
    ) -> None:
        def _action() -> JsonDict:
            series_data = self.api.get_series_by_id(series_id)
            series_data["monitored"] = False
            self.api.update_series(series_data)
            return {"episode_failures": 0}

        def _on_success(result: object) -> None:
            self._update_mon_column(item, False)
            msg = f"Unmonitored: {label}"
            failures = self._result_int(result, "episode_failures")
            if failures:
                msg += f" ({failures} episode update errors)"
            self.status_label.setText(msg)

        self._run_api_action(
            f"Unmonitoring: {label}...",
            _action,
            _on_success,
            "Failed to unmonitor",
        )

    def _ctx_unmonitor_season(
        self,
        item: QStandardItem,
        *,
        label: str,
        series_id: int,
    ) -> None:
        season_num = self._item_role_int(item, ROLE_SEASON_NUM)
        if season_num is None:
            self.status_label.setText("No season number available")
            return
        season_episodes = self._collect_episode_payloads(item)

        def _action() -> JsonDict:
            series_data = self.api.get_series_by_id(series_id)
            seasons_obj = series_data.get("seasons", [])
            seasons = (
                cast("list[JsonDict]", seasons_obj)
                if isinstance(seasons_obj, list)
                else []
            )
            for season in seasons:
                if season.get("seasonNumber") == season_num:
                    season["monitored"] = False
                    break
            self.api.update_series(series_data)
            failures = 0
            for episode in season_episodes:
                episode["monitored"] = False
                try:
                    self.api.update_episode(episode)
                except Exception:
                    failures += 1
            return {"episode_failures": failures}

        def _on_success(result: object) -> None:
            self._update_mon_column(item, False)
            for row in range(item.rowCount()):
                ep_item = item.child(row, 0)
                self._update_mon_column(ep_item, False)
            msg = f"Unmonitored: {label}"
            failures = self._result_int(result, "episode_failures")
            if failures:
                msg += f" ({failures} episode update errors)"
            self.status_label.setText(msg)

        self._run_api_action(
            f"Unmonitoring: {label}...",
            _action,
            _on_success,
            "Failed to unmonitor",
        )

    def _ctx_unmonitor_episode(
        self,
        item: QStandardItem,
        *,
        label: str,
    ) -> None:
        ep_payloads = self._item_role_payloads(item, ROLE_EPISODE_DATA)
        if not ep_payloads:
            self.status_label.setText("No episode payload found for unmonitor action")
            return

        def _action() -> JsonDict:
            for episode in ep_payloads:
                episode["monitored"] = False
                self.api.update_episode(episode)
            return {"episode_failures": 0}

        def _on_success(_result: object) -> None:
            self._update_mon_column(item, False)
            self.status_label.setText(f"Unmonitored: {label}")

        self._run_api_action(
            f"Unmonitoring: {label}...",
            _action,
            _on_success,
            "Failed to unmonitor",
        )

    def _ctx_unmonitor(self, item: QStandardItem, node_type: str):
        series_id = self._item_role_int(item, ROLE_SERIES_ID)
        label = item.text()
        if series_id is None:
            self.status_label.setText("No series ID available")
            return
        handlers = {
            "series": self._ctx_unmonitor_series,
            "season": self._ctx_unmonitor_season,
        }
        handler = handlers.get(node_type)
        if handler is not None:
            handler(item, label=label, series_id=series_id)
            return
        if node_type == "episode":
            self._ctx_unmonitor_episode(item, label=label)

    def _ctx_auto_search(self, item: QStandardItem, node_type: str):
        series_id = self._item_role_int(item, ROLE_SERIES_ID)
        label = item.text()
        if series_id is None:
            self.status_label.setText("No series ID available")
            return

        if node_type == "series":

            def _action():
                self.api.series_search(series_id)
                return None

            def _on_success(_result: object) -> None:
                self.status_label.setText(f"Auto search started: {label}")

            self._run_api_action(
                f"Starting auto search: {label}...",
                _action,
                _on_success,
                "Auto search failed",
            )
            return

        if node_type == "season":
            season_num = self._item_role_int(item, ROLE_SEASON_NUM)
            if season_num is None:
                self.status_label.setText("No season number available")
                return

            def _action():
                self.api.season_search(series_id, season_num)
                return None

            def _on_success(_result: object) -> None:
                self.status_label.setText(f"Auto search started: {label}")

            self._run_api_action(
                f"Starting auto search: {label}...",
                _action,
                _on_success,
                "Auto search failed",
            )
            return

        if node_type == "episode":
            ep_payloads = self._item_role_payloads(item, ROLE_EPISODE_DATA)
            ep_ids = [
                episode_id
                for ep in ep_payloads
                if isinstance((episode_id := ep.get("id")), int)
            ]
            if not ep_ids:
                QMessageBox.warning(
                    self, "No Episodes", "No episode IDs found for auto search."
                )
                return

            def _action():
                self.api.episode_search(ep_ids)
                return None

            def _on_success(_result: object) -> None:
                self.status_label.setText(f"Auto search started: {label}")

            self._run_api_action(
                f"Starting auto search: {label}...",
                _action,
                _on_success,
                "Auto search failed",
            )

    def _run_manual_search_dialog(
        self,
        *,
        label: str,
        releases: object,
    ) -> None:
        release_list = self._as_json_list(releases)
        if not release_list:
            QMessageBox.information(self, "No Results", "No releases found.")
            self.status_label.setText("No releases found")
            return
        dlg = ManualSearchDialog(
            self,
            label,
            release_list,
            settings=self._settings,
        )
        if dlg.exec() != _DIALOG_ACCEPTED or not dlg.selected_release:
            self.status_label.setText("Manual search cancelled")
            return
        rel = dlg.selected_release

        def _download_action() -> None:
            guid = rel.get("guid")
            indexer_id = rel.get("indexerId")
            if not isinstance(guid, str) or not isinstance(indexer_id, int):
                return
            self.api.download_release(guid, indexer_id)

        def _on_download_success(_result: object) -> None:
            self.status_label.setText(f"Download queued: {rel.get('title', '?')}")

        self._run_api_action(
            "Queueing selected release...",
            _download_action,
            _on_download_success,
            "Failed to queue download",
        )

    def _ctx_manual_search_episode(self, item: QStandardItem, *, label: str) -> None:
        ep_payloads = self._item_role_payloads(item, ROLE_EPISODE_DATA)
        ep_ids = [
            episode_id
            for ep in ep_payloads
            if isinstance((episode_id := ep.get("id")), int)
        ]
        if not ep_ids:
            QMessageBox.warning(
                self, "No Episodes", "No episode IDs found for manual search."
            )
            return

        def _search_action() -> object:
            return self.api.get_release(ep_ids[0])

        self._run_api_action(
            f"Searching releases for episode: {label}...",
            _search_action,
            lambda releases: self._run_manual_search_dialog(
                label=label,
                releases=releases,
            ),
            "Manual search failed",
        )

    def _ctx_manual_search_season(self, item: QStandardItem, *, label: str) -> None:
        series_id = self._item_role_int(item, ROLE_SERIES_ID)
        season_num = self._item_role_int(item, ROLE_SEASON_NUM)
        if series_id is None or season_num is None:
            QMessageBox.warning(
                self,
                "No Series/Season",
                "No series or season info found for manual search.",
            )
            return

        def _search_action() -> object:
            return self.api.get_release_by_season(series_id, season_num)

        self._run_api_action(
            f"Searching releases for season: {label}...",
            _search_action,
            lambda releases: self._run_manual_search_dialog(
                label=label,
                releases=releases,
            ),
            "Manual search failed",
        )

    def _ctx_manual_search_series(self, item: QStandardItem, *, label: str) -> None:
        series_id = self._item_role_int(item, ROLE_SERIES_ID)
        if series_id is None:
            QMessageBox.warning(
                self, "No Series", "No series ID found for manual search."
            )
            return

        def _search_action() -> object:
            return self.api.get_release_by_series(series_id)

        self._run_api_action(
            f"Searching releases for series: {label}...",
            _search_action,
            lambda releases: self._run_manual_search_dialog(
                label=label,
                releases=releases,
            ),
            "Manual search failed",
        )

    def _ctx_manual_search(self, item: QStandardItem, node_type: str):
        label = item.text()
        if node_type == "episode":
            self._ctx_manual_search_episode(item, label=label)
            return
        if node_type == "season":
            self._ctx_manual_search_season(item, label=label)
            return
        if node_type == "series":
            self._ctx_manual_search_series(item, label=label)

    def _warn_disk_delete_failure(self, fs_error: str) -> None:
        """Show one standard warning dialog for disk-delete failures."""

        if fs_error:
            QMessageBox.warning(
                self,
                "Delete Warning",
                f"Disk delete failed:\n{fs_error}",
            )

    def _unmonitor_series_season(self, series_id: int, season_num: int) -> bool:
        """Persist one season-level unmonitor flag on the series payload."""

        try:
            series_data = self.api.get_series_by_id(series_id)
            seasons_obj = series_data.get("seasons", [])
            seasons = (
                cast("list[JsonDict]", seasons_obj)
                if isinstance(seasons_obj, list)
                else []
            )
            for season in seasons:
                if season.get("seasonNumber") == season_num:
                    season["monitored"] = False
                    break
            self.api.update_series(series_data)
            return True
        except Exception:
            return False

    def _ctx_delete_series_from_disk(
        self,
        item: QStandardItem,
        *,
        label: str,
    ) -> None:
        series_path = self._item_role_str(item, ROLE_SERIES_PATH)
        series_file_ids: list[int] = []
        for s_row in range(item.rowCount()):
            season_item = item.child(s_row, 0)
            for ep_row in range(season_item.rowCount()):
                ep_item = season_item.child(ep_row, 0)
                file_id = self._item_role_int(ep_item, ROLE_FILE_ID)
                if file_id:
                    series_file_ids.append(file_id)

        def _action() -> JsonDict:
            api_failures = 0
            for file_id in series_file_ids:
                try:
                    self.api.delete_episode_file(file_id)
                except Exception:
                    api_failures += 1
            fs_deleted, fs_error = self._delete_tree_path(series_path)
            return {
                "api_failures": api_failures,
                "fs_deleted": fs_deleted,
                "fs_error": fs_error,
            }

        def _on_success(result: object) -> None:
            msg = f"Deleted from disk: {label}"
            api_failures = self._result_int(result, "api_failures")
            fs_deleted = self._result_bool(result, "fs_deleted")
            fs_error = self._result_str(result, "fs_error")
            if api_failures:
                msg += f" ({api_failures} API delete errors)"
            if not fs_deleted:
                msg += " (disk delete failed)"
                self._warn_disk_delete_failure(fs_error)
            else:
                parent = item.parent() or self.model.invisibleRootItem()
                parent.removeRow(item.row())
            self.status_label.setText(msg)

        self._run_api_action(
            f"Deleting from disk: {label}...",
            _action,
            _on_success,
            "Delete from disk failed",
        )

    def _ctx_delete_season_from_disk(
        self,
        item: QStandardItem,
        *,
        label: str,
    ) -> None:
        season_path = self._item_role_str(item, ROLE_SEASON_PATH)
        file_ids: list[int] = []
        for row in range(item.rowCount()):
            ep_item = item.child(row, 0)
            file_id = self._item_role_int(ep_item, ROLE_FILE_ID)
            if file_id:
                file_ids.append(file_id)

        def _action() -> JsonDict:
            api_failures = 0
            for file_id in file_ids:
                try:
                    self.api.delete_episode_file(file_id)
                except Exception:
                    api_failures += 1
            fs_deleted, fs_error = self._delete_tree_path(season_path)
            return {
                "api_failures": api_failures,
                "fs_deleted": fs_deleted,
                "fs_error": fs_error,
            }

        def _on_success(result: object) -> None:
            msg = f"Deleted from disk: {label}"
            api_failures = self._result_int(result, "api_failures")
            fs_deleted = self._result_bool(result, "fs_deleted")
            fs_error = self._result_str(result, "fs_error")
            if api_failures:
                msg += f" ({api_failures} API delete errors)"
            if not fs_deleted:
                msg += " (disk delete failed)"
                self._warn_disk_delete_failure(fs_error)
            else:
                parent = item.parent() or self.model.invisibleRootItem()
                parent.removeRow(item.row())
            self.status_label.setText(msg)

        self._run_api_action(
            f"Deleting from disk: {label}...",
            _action,
            _on_success,
            "Delete from disk failed",
        )

    def _ctx_delete_episode_from_disk(
        self,
        item: QStandardItem,
        *,
        label: str,
    ) -> None:
        file_id = self._item_role_int(item, ROLE_FILE_ID)
        file_path = self._item_role_str(item, ROLE_FILE_PATH)
        if not file_id and not file_path:
            self.status_label.setText("No file found for this episode")
            return

        def _action() -> JsonDict:
            api_failures = 0
            if file_id:
                try:
                    self.api.delete_episode_file(file_id)
                except Exception:
                    api_failures += 1
            fs_deleted, fs_error = self._delete_file_path(file_path)
            return {
                "api_failures": api_failures,
                "fs_deleted": fs_deleted,
                "fs_error": fs_error,
            }

        def _on_success(result: object) -> None:
            msg = f"Deleted from disk: {label}"
            api_failures = self._result_int(result, "api_failures")
            fs_deleted = self._result_bool(result, "fs_deleted")
            fs_error = self._result_str(result, "fs_error")
            if api_failures:
                msg += f" ({api_failures} API delete errors)"
            if not fs_deleted:
                msg += " (disk delete failed)"
                self._warn_disk_delete_failure(fs_error)
            else:
                parent = item.parent() or self.model.invisibleRootItem()
                parent.removeRow(item.row())
            self.status_label.setText(msg)

        self._run_api_action(
            f"Deleting from disk: {label}...",
            _action,
            _on_success,
            "Delete from disk failed",
        )

    def _ctx_delete_from_disk(self, item: QStandardItem, node_type: str):
        label = item.text()
        reply = QMessageBox.question(
            self,
            "Delete from Disk",
            f'Delete "{label}" files from disk?\n\n'
            "Monitoring status will not change.\nThis cannot be undone.",
            _MSG_YES | _MSG_NO,
            _MSG_NO,
        )
        if reply != _MSG_YES:
            return
        if node_type == "series":
            self._ctx_delete_series_from_disk(item, label=label)
            return
        if node_type == "season":
            self._ctx_delete_season_from_disk(item, label=label)
            return
        if node_type == "episode":
            self._ctx_delete_episode_from_disk(item, label=label)

    def _show_quality_profile_dialog(
        self,
        *,
        label: str,
        current_qp_id: int,
    ) -> tuple[int, str] | None:
        """Prompt for a quality profile selection and return the chosen pair."""

        dlg = QDialog(self)
        dlg.setWindowTitle("Change Quality Profile")
        lay = QVBoxLayout(dlg)
        lay.addWidget(QLabel(f"Quality profile for: {label}"))
        combo = QComboBox()
        current_idx = 0
        for profile in self._quality_profiles:
            profile_name = profile.get("name")
            profile_id = profile.get("id")
            if not isinstance(profile_name, str) or not isinstance(profile_id, int):
                continue
            combo_index = combo.count()
            combo.addItem(profile_name, profile_id)
            if profile_id == current_qp_id:
                current_idx = combo_index
        combo.setCurrentIndex(current_idx)
        lay.addWidget(combo)
        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        _ = btns.accepted.connect(dlg.accept)
        _ = btns.rejected.connect(dlg.reject)
        lay.addWidget(btns)
        if dlg.exec() != _DIALOG_ACCEPTED:
            self.status_label.setText("Quality profile change cancelled")
            return None
        new_qp_id = combo.currentData()
        if not isinstance(new_qp_id, int):
            self.status_label.setText("Selected quality profile is invalid")
            return None
        return new_qp_id, combo.currentText()

    def _apply_quality_profile_change(
        self,
        item: QStandardItem,
        *,
        label: str,
        series_payload: JsonDict,
        current_qp_id: int,
    ) -> None:
        """Prompt for, validate, and apply one quality profile change."""

        selection = self._show_quality_profile_dialog(
            label=label,
            current_qp_id=current_qp_id,
        )
        if selection is None:
            return
        new_qp_id, selected_name = selection
        if new_qp_id == current_qp_id:
            self.status_label.setText("Quality profile unchanged")
            return

        def _update_action() -> None:
            payload: JsonDict = dict(series_payload)
            payload["qualityProfileId"] = new_qp_id
            self.api.update_series(payload)

        def _on_updated(_result: object) -> None:
            row_idx = item.index().row()
            parent = item.parent() or self.model.invisibleRootItem()
            qp_item = parent.child(row_idx, 3)
            if qp_item:
                qp_item.setText(selected_name)
            self.status_label.setText(
                f"Quality profile changed to {selected_name} for {label}"
            )

        self._run_api_action(
            f"Updating quality profile: {label}...",
            _update_action,
            _on_updated,
            "Failed to update quality profile",
        )

    def _ctx_change_quality_profile(self, item: QStandardItem):
        """Change quality profile for a series via a combo-box dialog."""
        node_type = self._item_role_str(item, ROLE_NODE_TYPE)
        if node_type != "series":
            self.status_label.setText("Quality profile can only be changed on a series")
            return
        series_id = self._item_role_int(item, ROLE_SERIES_ID)
        if series_id is None:
            self.status_label.setText("No series ID available")
            return
        if not self._quality_profiles:
            QMessageBox.critical(self, "Error", "No quality profiles available.")
            return
        label = item.text()

        def _fetch_action():
            return self.api.get_series_by_id(series_id)

        def _on_fetched(series_data: object) -> None:
            series_payload = self._as_json_dict(series_data)
            if series_payload is None:
                self.status_label.setText("Failed to load quality profile data")
                return
            current_qp_value = series_payload.get("qualityProfileId", 0)
            current_qp_id = current_qp_value if isinstance(current_qp_value, int) else 0
            self._apply_quality_profile_change(
                item,
                label=label,
                series_payload=series_payload,
                current_qp_id=current_qp_id,
            )

        self._run_api_action(
            f"Loading quality profile options: {label}...",
            _fetch_action,
            _on_fetched,
            "Failed to fetch series",
        )

    def _ctx_unmonitor_delete_series(
        self,
        item: QStandardItem,
        *,
        label: str,
        series_id: int,
    ) -> None:
        def _action() -> JsonDict:
            self.api.delete_series(series_id, delete_files=True)
            return {"api_failures": 0}

        def _on_success(_result: object) -> None:
            parent = item.parent() or self.model.invisibleRootItem()
            parent.removeRow(item.row())
            self.status_label.setText(f"Deleted & unmonitored: {label}")

        self._run_api_action(
            f"Deleting & unmonitoring: {label}...",
            _action,
            _on_success,
            "Unmonitor & delete failed",
        )

    def _ctx_unmonitor_delete_season(
        self,
        item: QStandardItem,
        *,
        label: str,
        series_id: int,
    ) -> None:
        season_num = self._item_role_int(item, ROLE_SEASON_NUM)
        season_path = self._item_role_str(item, ROLE_SEASON_PATH)
        if season_num is None:
            self.status_label.setText("No season number available")
            return
        season_file_ids: list[int] = []
        for row in range(item.rowCount()):
            ep_item = item.child(row, 0)
            file_id = self._item_role_int(ep_item, ROLE_FILE_ID)
            if file_id:
                season_file_ids.append(file_id)
        season_episodes = self._collect_episode_payloads(item)

        def _action() -> JsonDict:
            api_failures = 0
            for file_id in season_file_ids:
                try:
                    self.api.delete_episode_file(file_id)
                except Exception:
                    api_failures += 1
            for episode in season_episodes:
                episode["monitored"] = False
                try:
                    self.api.update_episode(episode)
                except Exception:
                    api_failures += 1
            if not self._unmonitor_series_season(series_id, season_num):
                api_failures += 1
            fs_deleted, fs_error = self._delete_tree_path(season_path)
            return {
                "api_failures": api_failures,
                "fs_deleted": fs_deleted,
                "fs_error": fs_error,
            }

        def _on_success(result: object) -> None:
            msg = f"Deleted & unmonitored: {label}"
            api_failures = self._result_int(result, "api_failures")
            fs_deleted = self._result_bool(result, "fs_deleted")
            fs_error = self._result_str(result, "fs_error")
            if api_failures:
                msg += f" ({api_failures} API errors)"
            if not fs_deleted:
                msg += " (disk delete failed)"
                self._warn_disk_delete_failure(fs_error)
                self._update_mon_column(item, False)
                for row in range(item.rowCount()):
                    ep_item = item.child(row, 0)
                    self._update_mon_column(ep_item, False)
            else:
                parent = item.parent() or self.model.invisibleRootItem()
                parent.removeRow(item.row())
            self.status_label.setText(msg)

        self._run_api_action(
            f"Deleting & unmonitoring: {label}...",
            _action,
            _on_success,
            "Unmonitor & delete failed",
        )

    def _ctx_unmonitor_delete_episode(
        self,
        item: QStandardItem,
        *,
        label: str,
    ) -> None:
        file_id = self._item_role_int(item, ROLE_FILE_ID)
        file_path = self._item_role_str(item, ROLE_FILE_PATH)
        ep_payloads = self._item_role_payloads(item, ROLE_EPISODE_DATA)

        def _action() -> JsonDict:
            api_failures = 0
            for episode in ep_payloads:
                episode["monitored"] = False
                try:
                    self.api.update_episode(episode)
                except Exception:
                    api_failures += 1
            if file_id:
                try:
                    self.api.delete_episode_file(file_id)
                except Exception:
                    api_failures += 1
            fs_deleted, fs_error = self._delete_file_path(file_path)
            return {
                "api_failures": api_failures,
                "fs_deleted": fs_deleted,
                "fs_error": fs_error,
            }

        def _on_success(result: object) -> None:
            msg = f"Deleted & unmonitored: {label}"
            api_failures = self._result_int(result, "api_failures")
            fs_deleted = self._result_bool(result, "fs_deleted")
            fs_error = self._result_str(result, "fs_error")
            if api_failures:
                msg += f" ({api_failures} API errors)"
            if not fs_deleted:
                msg += " (disk delete failed)"
                self._warn_disk_delete_failure(fs_error)
                self._update_mon_column(item, False)
            else:
                parent = item.parent() or self.model.invisibleRootItem()
                parent.removeRow(item.row())
            self.status_label.setText(msg)

        self._run_api_action(
            f"Deleting & unmonitoring: {label}...",
            _action,
            _on_success,
            "Unmonitor & delete failed",
        )

    def _ctx_unmonitor_delete(self, item: QStandardItem, node_type: str):
        series_id = self._item_role_int(item, ROLE_SERIES_ID)
        label = item.text()
        if series_id is None:
            self.status_label.setText("No series ID available")
            return
        reply = QMessageBox.question(
            self,
            "Unmonitor & Delete",
            f'Unmonitor "{label}" and delete files from disk?\n\n'
            "This cannot be undone.",
            _MSG_YES | _MSG_NO,
            _MSG_NO,
        )
        if reply != _MSG_YES:
            return
        if node_type == "series":
            self._ctx_unmonitor_delete_series(item, label=label, series_id=series_id)
            return
        if node_type == "season":
            self._ctx_unmonitor_delete_season(
                item,
                label=label,
                series_id=series_id,
            )
            return
        if node_type == "episode":
            self._ctx_unmonitor_delete_episode(item, label=label)

    # ── keyboard / double-click actions ────────────────────────

    def _current_item(self) -> QStandardItem | None:
        idx = self.tree.currentIndex()
        if not idx.isValid():
            return None
        if idx.column() != 0:
            idx = idx.siblingAtColumn(0)
        return self.model.itemFromIndex(idx)

    def _on_enter(self):
        item = self._current_item()
        if item:
            self._activate_item(item)

    def _on_double_click(self, index: QModelIndex):
        if index.column() != 0:
            index = index.siblingAtColumn(0)
        item = self.model.itemFromIndex(index)
        if item:
            self._activate_item(item)

    def _activate_item(self, item: QStandardItem):
        node_type = self._item_role_str(item, ROLE_NODE_TYPE)
        if node_type == "series":
            path = self._item_role_str(item, ROLE_SERIES_PATH)
            if path and os.path.isdir(path):
                open_path_in_default_app(path)
            else:
                self.status_label.setText(f"Directory not found: {path}")
        elif node_type == "season":
            path = self._item_role_str(item, ROLE_SEASON_PATH)
            if path and os.path.isdir(path):
                open_path_in_default_app(path)
            else:
                self.status_label.setText(f"Directory not found: {path}")
        elif node_type == "episode":
            path = self._item_role_str(item, ROLE_FILE_PATH)
            if path and os.path.isfile(path):
                open_path_in_default_app(path)
            else:
                self.status_label.setText(f"File not found: {path}")

    def _on_delete(self):
        if not self._ensure_action_idle():
            return
        item = self._current_item()
        if not item:
            return
        node_type = self._item_role_str(item, ROLE_NODE_TYPE)
        if node_type == "series":
            self._delete_series(item)
        elif node_type == "season":
            self._delete_season(item)

    def _delete_series(self, item: QStandardItem):
        series_id = self._item_role_int(item, ROLE_SERIES_ID)
        title = item.text()
        if series_id is None:
            self.status_label.setText("No series ID available")
            return
        reply = QMessageBox.question(
            self,
            "Delete Series",
            f'Delete "{title}" from Sonarr AND from disk?\n\nThis cannot be undone.',
            _MSG_YES | _MSG_NO,
            _MSG_NO,
        )
        if reply != _MSG_YES:
            return

        def _action():
            self.api.delete_series(series_id, delete_files=True)
            return None

        def _on_success(_result: object) -> None:
            parent = item.parent() or self.model.invisibleRootItem()
            parent.removeRow(item.row())
            self.status_label.setText(f"Deleted series: {title}")

        self._run_api_action(
            f"Deleting series: {title}...",
            _action,
            _on_success,
            "Failed to delete series",
        )

    def _delete_season(self, item: QStandardItem):
        # reuse the unmonitor & delete logic
        self._ctx_unmonitor_delete(item, "season")

    # ── Add Show ───────────────────────────────────────────────

    def _add_show(self):
        if not self._ensure_action_idle():
            return
        dlg = AddShowDialog(self, self.api)
        if dlg.exec() == _DIALOG_ACCEPTED and dlg.added_series:
            self._refresh()

    def closeEvent(self, event: QCloseEvent) -> None:
        """Save column widths and stop worker before closing."""
        widths = [self.tree.columnWidth(c) for c in range(len(self._columns))]
        self._settings.set_value(self._ui_key("column_widths"), widths)
        self._settings.sync()
        if not self._stop_action_worker(5000):
            QMessageBox.warning(
                self,
                "Action still running",
                "An action is still in progress.\n"
                "Please try closing again in a few seconds.",
            )
            event.ignore()
            return
        if not self._stop_worker(5000):
            reply = QMessageBox.question(
                self,
                "Background task still running",
                "A refresh is still in progress and did not stop in time.\n"
                "Force close now?",
                _MSG_YES | _MSG_NO,
                _MSG_NO,
            )
            if reply != _MSG_YES:
                event.ignore()
                return
            worker = self.worker
            if worker and worker.isRunning():
                worker.terminate()
                worker.wait(1500)
                if worker.isRunning():
                    QMessageBox.critical(
                        self,
                        "Unable to close safely",
                        "Background thread is still running after force-close "
                        "attempt.\n"
                        "Please try closing again in a few seconds.",
                    )
                    event.ignore()
                    return
            self.worker = None
        self.action_worker = None
        super().closeEvent(event)

    def _clear_cache_and_refresh(self):
        """Clear the ffprobe cache and reload all data."""
        if not self._ensure_action_idle("Action still running; try again in a moment"):
            return
        if not self._stop_worker(5000):
            QMessageBox.warning(
                self,
                "Refresh still running",
                "Cannot clear cache while a refresh is still running.\n"
                "Please wait and try again.",
            )
            self.status_label.setText("Refresh still running; try again in a moment")
            return
        clear_probe_cache()
        self.status_label.setText("Cache cleared, refreshing...")
        self._refresh()

    def _refresh(self):
        """Reload all data from Sonarr."""
        if not self._ensure_action_idle("Action still running; try again"):
            return
        if not self._stop_worker(5000):
            QMessageBox.warning(
                self,
                "Refresh still running",
                "Previous refresh is still stopping.\n"
                "Please try again in a few seconds.",
            )
            self.status_label.setText("Previous refresh still stopping; try again")
            return
        self.model.removeRows(0, self.model.rowCount())
        self.progress_bar.show()
        self.progress_bar.setRange(0, 0)
        self.status_label.setText("Refreshing...")
        self._start_worker()
