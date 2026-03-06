"""Main Qt window for Sonarr UI helper."""

import os
import shutil
from collections.abc import Callable
from pathlib import Path
from typing import Any

from PySide6.QtCore import QModelIndex, QSettings, Qt
from PySide6.QtGui import (
    QAction,
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

from .api import SonarrAPI
from .dialogs.add_show import AddShowDialog
from .dialogs.manual_search import ManualSearchDialog
from .helpers import _open_path, fmt_size
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

SETTINGS_ORG_NAME = 'SonarrUIHelper'
SETTINGS_APP_NAME = 'SonarrUIHelper'


class MainWindow(QMainWindow):
    def __init__(self, api: SonarrAPI, loader_api: SonarrAPI | None = None, settings: dict = None):
        super().__init__()
        self.api = api
        self.loader_api = loader_api or api
        self.cfg = settings or {}
        self.preferences_store = QSettings(
            QSettings.IniFormat,
            QSettings.UserScope,
            SETTINGS_ORG_NAME,
            SETTINGS_APP_NAME,
        )
        self._settings = self.preferences_store
        self.setWindowTitle('Sonarr UI Helper')
        self.resize(1600, 800)
        self._last_worker_error = ''
        self._worker_warning_count = 0

        # ── Menu bar ──────────────────────────────────────────
        self._build_menu_bar()

        # central widget
        central = QWidget()
        layout = QVBoxLayout(central)
        layout.setContentsMargins(4, 4, 4, 4)

        # toolbar (with mnemonics via &)
        toolbar = QHBoxLayout()
        btn_expand_all = QPushButton('Expand &All')
        btn_expand_all.clicked.connect(self._expand_all)
        btn_expand_series = QPushButton('Expand &Series')
        btn_expand_series.clicked.connect(self._expand_series)
        btn_collapse_seasons = QPushButton('Collapse S&easons')
        btn_collapse_seasons.clicked.connect(self._collapse_all_seasons)
        btn_collapse_series = QPushButton('&Collapse Series')
        btn_collapse_series.clicked.connect(self._collapse_all_series)
        btn_add_show = QPushButton('A&dd Show')
        btn_add_show.clicked.connect(self._add_show)
        toolbar.addWidget(btn_expand_all)
        toolbar.addWidget(btn_expand_series)
        toolbar.addWidget(btn_collapse_seasons)
        toolbar.addWidget(btn_collapse_series)
        btn_refresh = QPushButton('&Refresh')
        btn_refresh.clicked.connect(self._refresh)
        self.chk_show_missing = QCheckBox('Show &Missing')
        self.chk_show_missing.setChecked(False)
        self.chk_show_missing.toggled.connect(self._toggle_missing)
        toolbar.addWidget(self.chk_show_missing)
        toolbar.addStretch()
        toolbar.addWidget(btn_add_show)
        toolbar.addWidget(btn_refresh)
        layout.addLayout(toolbar)

        self.tree = QTreeView()
        self.tree.setAlternatingRowColors(True)
        self.tree.setUniformRowHeights(True)
        self.tree.setAnimated(False)
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._on_context_menu)
        self.tree.doubleClicked.connect(self._on_double_click)

        self.model = QStandardItemModel()
        self._columns = ['Name', 'Size', 'Mon', 'Quality Profile', 'Resolution', 'V.Bitrate',
                         'V.Codec', 'HDR', 'A.Codec', 'A.Bitrate', 'Audio Lang', 'Sub Lang']
        self.model.setHorizontalHeaderLabels(self._columns)
        self.tree.setModel(self.model)

        layout.addWidget(self.tree)
        self.setCentralWidget(central)

        # status bar with progress
        self.status_label = QLabel('Loading…')
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # indeterminate
        self.progress_bar.setMaximumWidth(200)
        sb = QStatusBar()
        sb.addWidget(self.status_label, 1)
        sb.addPermanentWidget(self.progress_bar)
        self.setStatusBar(sb)

        # keyboard: Delete key
        del_shortcut = QShortcut(QKeySequence(Qt.Key_Delete), self.tree)
        del_shortcut.activated.connect(self._on_delete)

        # keyboard: Enter/Return to activate
        enter_shortcut = QShortcut(QKeySequence(Qt.Key_Return), self.tree)
        enter_shortcut.activated.connect(self._on_enter)

        # cache quality profiles for name lookup
        self._quality_profiles = []
        self._qp_map = {}  # id -> name
        try:
            self._quality_profiles = api.get_quality_profiles()
            self._qp_map = {p['id']: p['name'] for p in self._quality_profiles}
        except Exception:
            pass

        # start loading
        self.worker: LoadWorker | None = None
        self.action_worker: ApiActionWorker | None = None
        self._start_worker()

    # ── menu bar ───────────────────────────────────────────────

    def _build_menu_bar(self):
        mb = self.menuBar()

        # File menu
        file_menu = mb.addMenu('&File')
        act = file_menu.addAction('&Add Show')
        act.setShortcut(QKeySequence('Ctrl+N'))
        act.triggered.connect(self._add_show)
        act = file_menu.addAction('&Refresh')
        act.setShortcut(QKeySequence('F5'))
        act.triggered.connect(self._refresh)
        act = file_menu.addAction('&Clear Cache && Refresh')
        act.setShortcut(QKeySequence('Ctrl+F5'))
        act.triggered.connect(self._clear_cache_and_refresh)
        file_menu.addSeparator()
        act = file_menu.addAction('E&xit')
        act.setShortcuts([QKeySequence('Ctrl+Q'), QKeySequence('Alt+X')])
        act.triggered.connect(self.close)

        # View menu
        view_menu = mb.addMenu('&View')
        act = view_menu.addAction('E&xpand All')
        act.setShortcut(QKeySequence('Ctrl+E'))
        act.triggered.connect(self._expand_all)
        act = view_menu.addAction('Expand &Series')
        act.setShortcut(QKeySequence('Ctrl+Shift+E'))
        act.triggered.connect(self._expand_series)
        act = view_menu.addAction('Collapse S&easons')
        act.setShortcut(QKeySequence('Ctrl+W'))
        act.triggered.connect(self._collapse_all_seasons)
        act = view_menu.addAction('&Collapse All')
        act.setShortcut(QKeySequence('Ctrl+Shift+W'))
        act.triggered.connect(self._collapse_all_series)
        view_menu.addSeparator()
        self.act_show_missing = QAction('Show &Missing', self)
        self.act_show_missing.setCheckable(True)
        self.act_show_missing.setChecked(False)
        self.act_show_missing.setShortcut(QKeySequence('Ctrl+M'))
        self.act_show_missing.toggled.connect(self._toggle_missing_from_menu)
        view_menu.addAction(self.act_show_missing)
        act = view_menu.addAction('Fit &Columns')
        act.triggered.connect(self._fit_columns)
        act = view_menu.addAction('&Reset View')
        act.setShortcut(QKeySequence('Ctrl+Shift+R'))
        act.triggered.connect(self._reset_view_settings)

        # Actions menu
        actions_menu = mb.addMenu('&Actions')
        act = actions_menu.addAction('&Monitor')
        act.setShortcut(QKeySequence('M'))
        act.triggered.connect(lambda: self._ctx_on_selected('monitor'))
        act = actions_menu.addAction('&Auto Search')
        act.setShortcut(QKeySequence('S'))
        act.triggered.connect(lambda: self._ctx_on_selected('auto_search'))
        act = actions_menu.addAction('Ma&nual Search')
        act.setShortcut(QKeySequence('N'))
        act.triggered.connect(lambda: self._ctx_on_selected('manual_search'))
        actions_menu.addSeparator()
        act = actions_menu.addAction('Change &Quality Profile')
        act.setShortcut(QKeySequence('Q'))
        act.triggered.connect(lambda: self._ctx_on_selected('change_quality_profile'))
        actions_menu.addSeparator()
        act = actions_menu.addAction('&Unmonitor')
        act.setShortcut(QKeySequence('U'))
        act.triggered.connect(lambda: self._ctx_on_selected('unmonitor'))
        act = actions_menu.addAction('&Delete from Disk')
        act.setShortcut(QKeySequence('D'))
        act.triggered.connect(lambda: self._ctx_on_selected('delete_from_disk'))
        act = actions_menu.addAction('Unmonitor && De&lete')
        act.setShortcut(QKeySequence('Ctrl+Delete'))
        act.triggered.connect(lambda: self._ctx_on_selected('unmonitor_delete'))
        actions_menu.addSeparator()
        act = actions_menu.addAction('&Open in Explorer')
        act.setShortcut(QKeySequence('O'))
        act.triggered.connect(self._on_enter)

        # Tools menu
        tools_menu = mb.addMenu('&Tools')
        act = tools_menu.addAction('Edit .ini file')
        act.triggered.connect(self._edit_ini_file)

        # Help menu
        help_menu = mb.addMenu('&Help')
        act = help_menu.addAction('&Keyboard Shortcuts')
        act.setShortcut(QKeySequence('F1'))
        act.triggered.connect(self._show_help)

    def _toggle_missing_from_menu(self, checked: bool):
        """Sync the menu checkbox with the toolbar checkbox."""
        self.chk_show_missing.setChecked(checked)

    def _reset_view_settings(self):
        reply = QMessageBox.question(
            self,
            'Reset View',
            'Reset all saved UI view settings to defaults?\n'
            'This clears saved column widths, splitter positions, and other QSettings state.',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        self._settings.clear()
        self._settings.sync()

        # Re-apply in-memory defaults immediately.
        self.chk_show_missing.setChecked(False)
        self._apply_default_column_widths()
        self.status_label.setText('View settings reset to defaults')

    def _edit_ini_file(self):
        self._settings.sync()
        ini_path = Path(self.preferences_store.fileName())
        try:
            ini_path.parent.mkdir(parents=True, exist_ok=True)
            ini_path.touch(exist_ok=True)
            _open_path(str(ini_path))
            self.status_label.setText(f'Opened settings file: {ini_path}')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to open settings file:\n{e}')
            self.status_label.setText('Failed to open settings file')

    def _ctx_on_selected(self, action_name: str):
        """Dispatch a context-menu action on the currently selected tree item."""
        if not self._ensure_action_idle():
            return
        item = self._current_item()
        if not item:
            self.status_label.setText('No item selected')
            return
        node_type = item.data(ROLE_NODE_TYPE)
        if action_name == 'monitor':
            self._ctx_monitor(item, node_type)
        elif action_name == 'auto_search':
            self._ctx_auto_search(item, node_type)
        elif action_name == 'manual_search':
            self._ctx_manual_search(item, node_type)
        elif action_name == 'change_quality_profile':
            self._ctx_change_quality_profile(item)
        elif action_name == 'unmonitor':
            self._ctx_unmonitor(item, node_type)
        elif action_name == 'delete_from_disk':
            self._ctx_delete_from_disk(item, node_type)
        elif action_name == 'unmonitor_delete':
            self._ctx_unmonitor_delete(item, node_type)

    def _show_help(self):
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
            "<tr><td><code>Delete</code></td><td>Remove series/season from Sonarr (deletes files)</td></tr>"
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
            "<tr><td><code>Q</code></td><td>Change Quality Profile (series only)</td></tr>"
            "<tr><td><code>U</code></td><td>Unmonitor</td></tr>"
            "<tr><td><code>D</code></td><td>Delete files from disk (keep in Sonarr)</td></tr>"
            "<tr><td><code>Ctrl+Delete</code></td><td>Unmonitor &amp; Delete from disk</td></tr>"
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
        QMessageBox.information(self, 'Keyboard Shortcuts', help_text)

    # ── populate tree ───────────────────────────────────────────

    def _on_progress(self, text: str):
        self.status_label.setText(text)

    def _on_worker_progress(self, worker: LoadWorker, text: str):
        if worker is not self.worker:
            return
        if text.startswith('Error:'):
            self._last_worker_error = text
        elif text.startswith('Warning:'):
            self._worker_warning_count += 1
        self._on_progress(text)

    def _on_worker_series_ready(self, worker: LoadWorker, series_list: list):
        if worker is not self.worker:
            return
        self._on_data_loaded(series_list)

    def _start_worker(self):
        self._last_worker_error = ''
        self._worker_warning_count = 0
        worker = LoadWorker(self.loader_api)
        worker.progress.connect(lambda text, w=worker: self._on_worker_progress(w, text))
        worker.series_ready.connect(lambda data, w=worker: self._on_worker_series_ready(w, data))
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

    def _ensure_action_idle(self, status_text: str = 'Another action is still running') -> bool:
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
        action: Callable[[], Any],
        on_success: Callable[[Any], None],
        error_text: str,
    ) -> bool:
        if not self._ensure_action_idle():
            return False

        worker = ApiActionWorker(action)
        self.action_worker = worker
        self.status_label.setText(start_text)
        QApplication.setOverrideCursor(Qt.WaitCursor)

        def _on_done(result: Any, error: Any, w: ApiActionWorker = worker):
            if self.action_worker is not w:
                return
            self.action_worker = None
            QApplication.restoreOverrideCursor()
            try:
                if error is not None:
                    QMessageBox.critical(self, 'Error', f'{error_text}:\n{error}')
                    self.status_label.setText(error_text)
                    return
                on_success(result)
            finally:
                w.deleteLater()

        worker.finished_action.connect(_on_done)
        worker.start()
        return True

    @staticmethod
    def _make_row(cols: int) -> list:
        """Create a list of empty, non-editable QStandardItems."""
        row = []
        for _ in range(cols):
            item = QStandardItem('')
            item.setEditable(False)
            row.append(item)
        return row

    def _highlight_row(self, row: list, sub_langs: str):
        """Apply light red background if configured sub language is missing.

        highlight_missing_subs is a label that maps to english_language_codes,
        e.g. highlight_missing_subs = "english" with
        english_language_codes = ["eng", "en", "english"] means any of those
        codes count as a match.
        """
        hl = self.cfg.get('highlight_missing_subs', '').strip().lower()
        if not hl:
            return
        # expand via the language codes array
        codes = [c.lower() for c in self.cfg.get('english_language_codes', [hl])]
        if not codes:
            codes = [hl]
        langs = {l.strip().lower() for l in sub_langs.split(',') if l.strip()}
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
        self.status_label.setText('Columns fitted to contents')

    def _on_data_loaded(self, series_list: list):
        self.progress_bar.hide()
        self.model.removeRows(0, self.model.rowCount())

        if not series_list:
            if self._last_worker_error:
                self.status_label.setText(self._last_worker_error)
            elif self._worker_warning_count:
                self.status_label.setText(f'No series loaded ({self._worker_warning_count} warnings)')
            else:
                self.status_label.setText('No series with downloaded episodes found.')
            return

        total_series = len(series_list)
        missing_color = QColor(128, 128, 128)  # grey for missing episodes

        for s in series_list:
            # ── Series row ──
            series_label = f"{s['title']}, {s['year']}" if s['year'] else s['title']
            series_item = QStandardItem(series_label)
            series_item.setEditable(False)
            series_item.setData('series', ROLE_NODE_TYPE)
            series_item.setData(s['series_id'], ROLE_SERIES_ID)
            series_item.setData(s['path'], ROLE_SERIES_PATH)
            font = series_item.font()
            font.setBold(True)
            series_item.setFont(font)

            series_row = self._make_row(len(self._columns))
            series_row[0] = series_item
            series_row[1].setText(fmt_size(s['total_size']))
            series_row[2].setText('Y' if s['series_data'].get('monitored', False) else 'N')
            qp_id = s['series_data'].get('qualityProfileId', 0)
            series_row[3].setText(self._qp_map.get(qp_id, str(qp_id)))

            all_season_nums = s['all_season_nums']
            downloaded_seasons = s['seasons']
            missing_seasons = s['missing_seasons']

            for sn in all_season_nums:
                downloaded_eps = downloaded_seasons.get(sn, [])
                missing_eps = missing_seasons.get(sn, [])
                season_size = sum(e['size_bytes'] for e in downloaded_eps)

                # build season path from first downloaded episode's directory
                season_path = ''
                if downloaded_eps and downloaded_eps[0]['file_path']:
                    season_path = str(Path(downloaded_eps[0]['file_path']).parent)

                season_label = f"Season {sn}" if sn > 0 else "Specials"
                has_downloaded = len(downloaded_eps) > 0
                has_missing = len(missing_eps) > 0
                if has_downloaded and has_missing:
                    season_label += f" ({len(downloaded_eps)} downloaded, {len(missing_eps)} missing)"
                elif not has_downloaded:
                    season_label += f" ({len(missing_eps)} missing)"

                season_item = QStandardItem(season_label)
                season_item.setEditable(False)
                season_item.setData('season', ROLE_NODE_TYPE)
                season_item.setData(s['series_id'], ROLE_SERIES_ID)
                season_item.setData(sn, ROLE_SEASON_NUM)
                season_item.setData(season_path, ROLE_SEASON_PATH)
                season_item.setData(not has_downloaded, ROLE_IS_MISSING)
                if not has_downloaded:
                    season_item.setForeground(missing_color)

                # season monitored status from series metadata
                season_monitored = True
                for s_info in s['series_data'].get('seasons', []):
                    if s_info.get('seasonNumber') == sn:
                        season_monitored = s_info.get('monitored', True)
                        break

                season_row = self._make_row(len(self._columns))
                season_row[0] = season_item
                if season_size > 0:
                    season_row[1].setText(fmt_size(season_size))
                season_row[2].setText('Y' if season_monitored else 'N')

                # downloaded episodes
                for ep in downloaded_eps:
                    ep_label = f"E{ep['episode_number']:02d} - {ep['file_name']}"
                    ep_item = QStandardItem(ep_label)
                    ep_item.setEditable(False)
                    ep_item.setData('episode', ROLE_NODE_TYPE)
                    ep_item.setData(ep['file_path'], ROLE_FILE_PATH)
                    ep_item.setData(ep['file_id'], ROLE_FILE_ID)
                    ep_item.setData(ep['episode_data'], ROLE_EPISODE_DATA)
                    ep_item.setData(s['series_id'], ROLE_SERIES_ID)
                    ep_item.setData(sn, ROLE_SEASON_NUM)
                    ep_item.setData(False, ROLE_IS_MISSING)

                    ep_row = self._make_row(len(self._columns))
                    ep_row[0] = ep_item
                    ep_row[1].setText(fmt_size(ep['size_bytes']))
                    ep_monitored = any(e.get('monitored', False) for e in ep.get('episode_data', []))
                    ep_row[2].setText('Y' if ep_monitored else 'N')
                    ep_row[4].setText(ep['video_resolution'])
                    ep_row[5].setText(ep['video_bitrate'])
                    ep_row[6].setText(ep['video_codec'])
                    ep_row[7].setText(ep['hdr'])
                    ep_row[8].setText(ep['audio_codec'])
                    ep_row[9].setText(ep['audio_bitrate'])
                    ep_row[10].setText(ep['audio_langs'])
                    ep_row[11].setText(ep['sub_langs'])
                    self._highlight_row(ep_row, ep['sub_langs'])

                    season_item.appendRow(ep_row)

                # missing episodes
                for mep in missing_eps:
                    ep_num = mep.get('episodeNumber', 0)
                    ep_title = mep.get('title', '')
                    monitored = mep.get('monitored', False)
                    ep_label = f"E{ep_num:02d} - {ep_title}"
                    ep_item = QStandardItem(ep_label)
                    ep_item.setEditable(False)
                    ep_item.setForeground(missing_color)
                    ep_item.setData('episode', ROLE_NODE_TYPE)
                    ep_item.setData(s['series_id'], ROLE_SERIES_ID)
                    ep_item.setData(sn, ROLE_SEASON_NUM)
                    ep_item.setData(True, ROLE_IS_MISSING)
                    ep_item.setData([mep], ROLE_EPISODE_DATA)

                    ep_row = self._make_row(len(self._columns))
                    ep_row[0] = ep_item
                    ep_row[2].setText('Y' if monitored else 'N')

                    season_item.appendRow(ep_row)

                series_item.appendRow(season_row)

            self.model.appendRow(series_row)

        # expand all to season level (skip Specials / season 0)
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

        # resize columns
        header = self.tree.header()
        header.setSectionResizeMode(QHeaderView.Interactive)
        saved = self._settings.value('column_widths')
        if saved and len(saved) == len(self._columns):
            for col, w in enumerate(saved):
                self.tree.setColumnWidth(col, int(w))
        else:
            self._apply_default_column_widths()

        status_text = f'{total_series} series loaded'
        if self._worker_warning_count:
            status_text += f' ({self._worker_warning_count} warnings)'
        self.status_label.setText(status_text)

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

    def _on_context_menu(self, pos):
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

        node_type = item.data(ROLE_NODE_TYPE)
        menu = QMenu(self)

        act_monitor = menu.addAction('Monitor')
        act_auto_search = menu.addAction('Auto Search')
        act_manual_search = menu.addAction('Manual Search')
        menu.addSeparator()
        act_change_qp = menu.addAction('Change Quality Profile')
        act_change_qp.setEnabled(node_type == 'series')
        menu.addSeparator()
        act_unmonitor = menu.addAction('Unmonitor')
        act_delete_disk = menu.addAction('Delete From Disk')
        act_unmonitor_delete = menu.addAction('Unmonitor && Delete From Disk')

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

    def _update_mon_column(self, item: QStandardItem, monitored: bool):
        """Update the 'Mon' column (col 2) for the row containing item."""
        parent = item.parent() or self.model.invisibleRootItem()
        mon_item = parent.child(item.row(), 2)
        if mon_item:
            mon_item.setText('Y' if monitored else 'N')

    @staticmethod
    def _clone_episode_payloads(raw_payloads: Any) -> list[dict]:
        payloads: list[dict] = []
        for ep in raw_payloads or []:
            if isinstance(ep, dict):
                payloads.append(dict(ep))
        return payloads

    def _collect_episode_payloads(self, parent_item: QStandardItem) -> list[dict]:
        payloads: list[dict] = []
        for row in range(parent_item.rowCount()):
            ep_item = parent_item.child(row, 0)
            payloads.extend(self._clone_episode_payloads(ep_item.data(ROLE_EPISODE_DATA)))
        return payloads

    @staticmethod
    def _delete_tree_path(path: str) -> tuple[bool, str]:
        if not path:
            return True, ''
        if not os.path.exists(path):
            return True, ''
        try:
            shutil.rmtree(path)
        except Exception as e:
            if os.path.exists(path):
                return False, str(e)
        if os.path.exists(path):
            return False, 'Directory still exists after delete attempt.'
        return True, ''

    @staticmethod
    def _delete_file_path(path: str) -> tuple[bool, str]:
        if not path:
            return True, ''
        if not os.path.exists(path):
            return True, ''
        try:
            os.remove(path)
        except Exception as e:
            if os.path.exists(path):
                return False, str(e)
        if os.path.exists(path):
            return False, 'File still exists after delete attempt.'
        return True, ''

    def _ctx_monitor(self, item: QStandardItem, node_type: str):
        series_id = item.data(ROLE_SERIES_ID)
        label = item.text()

        if node_type == 'series':
            def _action():
                series_data = self.api.get_series_by_id(series_id)
                series_data['monitored'] = True
                self.api.update_series(series_data)
                return None

            def _on_success(_result):
                self._update_mon_column(item, True)
                self.status_label.setText(f'Monitored: {label}')

            self._run_api_action(f'Monitoring: {label}...', _action, _on_success, 'Failed to monitor')
            return

        if node_type == 'season':
            season_num = item.data(ROLE_SEASON_NUM)

            def _action():
                series_data = self.api.get_series_by_id(series_id)
                for s in series_data.get('seasons', []):
                    if s.get('seasonNumber') == season_num:
                        s['monitored'] = True
                        break
                self.api.update_series(series_data)
                return None

            def _on_success(_result):
                self._update_mon_column(item, True)
                self.status_label.setText(f'Monitored: {label}')

            self._run_api_action(f'Monitoring: {label}...', _action, _on_success, 'Failed to monitor')
            return

        if node_type == 'episode':
            ep_payloads = self._clone_episode_payloads(item.data(ROLE_EPISODE_DATA))
            if not ep_payloads:
                self.status_label.setText('No episode payload found for monitor action')
                return

            def _action():
                for ep in ep_payloads:
                    ep['monitored'] = True
                    self.api.update_episode(ep)
                return None

            def _on_success(_result):
                self._update_mon_column(item, True)
                self.status_label.setText(f'Monitored: {label}')

            self._run_api_action(f'Monitoring: {label}...', _action, _on_success, 'Failed to monitor')

    def _ctx_unmonitor(self, item: QStandardItem, node_type: str):
        series_id = item.data(ROLE_SERIES_ID)
        label = item.text()

        if node_type == 'series':
            def _action():
                series_data = self.api.get_series_by_id(series_id)
                series_data['monitored'] = False
                self.api.update_series(series_data)
                return {'episode_failures': 0}

            def _on_success(result):
                self._update_mon_column(item, False)
                msg = f'Unmonitored: {label}'
                failures = result.get('episode_failures', 0)
                if failures:
                    msg += f' ({failures} episode update errors)'
                self.status_label.setText(msg)

            self._run_api_action(f'Unmonitoring: {label}...', _action, _on_success, 'Failed to unmonitor')
            return

        if node_type == 'season':
            season_num = item.data(ROLE_SEASON_NUM)
            season_episodes = self._collect_episode_payloads(item)

            def _action():
                series_data = self.api.get_series_by_id(series_id)
                for s in series_data.get('seasons', []):
                    if s.get('seasonNumber') == season_num:
                        s['monitored'] = False
                        break
                self.api.update_series(series_data)
                failures = 0
                for ep in season_episodes:
                    ep['monitored'] = False
                    try:
                        self.api.update_episode(ep)
                    except Exception:
                        failures += 1
                return {'episode_failures': failures}

            def _on_success(result):
                self._update_mon_column(item, False)
                for row in range(item.rowCount()):
                    ep_item = item.child(row, 0)
                    self._update_mon_column(ep_item, False)
                msg = f'Unmonitored: {label}'
                failures = result.get('episode_failures', 0)
                if failures:
                    msg += f' ({failures} episode update errors)'
                self.status_label.setText(msg)

            self._run_api_action(f'Unmonitoring: {label}...', _action, _on_success, 'Failed to unmonitor')
            return

        if node_type == 'episode':
            ep_payloads = self._clone_episode_payloads(item.data(ROLE_EPISODE_DATA))
            if not ep_payloads:
                self.status_label.setText('No episode payload found for unmonitor action')
                return

            def _action():
                for ep in ep_payloads:
                    ep['monitored'] = False
                    self.api.update_episode(ep)
                return {'episode_failures': 0}

            def _on_success(_result):
                self._update_mon_column(item, False)
                self.status_label.setText(f'Unmonitored: {label}')

            self._run_api_action(f'Unmonitoring: {label}...', _action, _on_success, 'Failed to unmonitor')

    def _ctx_auto_search(self, item: QStandardItem, node_type: str):
        series_id = item.data(ROLE_SERIES_ID)
        label = item.text()

        if node_type == 'series':
            def _action():
                self.api.series_search(series_id)
                return None

            def _on_success(_result):
                self.status_label.setText(f'Auto search started: {label}')

            self._run_api_action(f'Starting auto search: {label}...', _action, _on_success, 'Auto search failed')
            return

        if node_type == 'season':
            season_num = item.data(ROLE_SEASON_NUM)

            def _action():
                self.api.season_search(series_id, season_num)
                return None

            def _on_success(_result):
                self.status_label.setText(f'Auto search started: {label}')

            self._run_api_action(f'Starting auto search: {label}...', _action, _on_success, 'Auto search failed')
            return

        if node_type == 'episode':
            ep_payloads = self._clone_episode_payloads(item.data(ROLE_EPISODE_DATA))
            ep_ids = [ep['id'] for ep in ep_payloads if 'id' in ep]
            if not ep_ids:
                QMessageBox.warning(self, 'No Episodes', 'No episode IDs found for auto search.')
                return

            def _action():
                self.api.episode_search(ep_ids)
                return None

            def _on_success(_result):
                self.status_label.setText(f'Auto search started: {label}')

            self._run_api_action(f'Starting auto search: {label}...', _action, _on_success, 'Auto search failed')

    def _ctx_manual_search(self, item: QStandardItem, node_type: str):
        label = item.text()
        series_id = item.data(ROLE_SERIES_ID)
        season_num = item.data(ROLE_SEASON_NUM)
        if node_type == 'episode':
            ep_payloads = self._clone_episode_payloads(item.data(ROLE_EPISODE_DATA))
            ep_ids = [ep['id'] for ep in ep_payloads if 'id' in ep]
            if not ep_ids:
                QMessageBox.warning(self, 'No Episodes', 'No episode IDs found for manual search.')
                return
            query_label = 'episode'
            def _search_action():
                return self.api.get_release(ep_ids[0])
        elif node_type == 'season':
            if series_id is None or season_num is None:
                QMessageBox.warning(self, 'No Series/Season', 'No series/season info found for manual search.')
                return
            query_label = 'season'
            def _search_action():
                return self.api.get_release_by_season(series_id, season_num)
        elif node_type == 'series':
            if series_id is None:
                QMessageBox.warning(self, 'No Series', 'No series ID found for manual search.')
                return
            query_label = 'series'
            def _search_action():
                return self.api.get_release_by_series(series_id)
        else:
            return
        def _on_search_success(releases: list[dict]):
            if not releases:
                QMessageBox.information(self, 'No Results', 'No releases found.')
                self.status_label.setText('No releases found')
                return
            dlg = ManualSearchDialog(self, label, releases, settings=self._settings)
            if dlg.exec() != QDialog.Accepted or not dlg.selected_release:
                self.status_label.setText('Manual search cancelled')
                return
            rel = dlg.selected_release
            def _download_action():
                self.api.download_release(rel['guid'], rel['indexerId'])
                return None
            def _on_download_success(_result):
                self.status_label.setText(f'Download queued: {rel.get("title", "?")}')
            self._run_api_action(
                'Queueing selected release...',
                _download_action,
                _on_download_success,
                'Failed to queue download',
            )
        self._run_api_action(
            f'Searching releases for {query_label}: {label}...',
            _search_action,
            _on_search_success,
            'Manual search failed',
        )
    def _ctx_delete_from_disk(self, item: QStandardItem, node_type: str):
        label = item.text()
        reply = QMessageBox.question(
            self, 'Delete from Disk',
            f'Delete "{label}" files from disk?\n\nMonitoring status will not change.\nThis cannot be undone.',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return
        if node_type == 'series':
            series_path = item.data(ROLE_SERIES_PATH)
            file_ids: list[int] = []
            for s_row in range(item.rowCount()):
                season_item = item.child(s_row, 0)
                for ep_row in range(season_item.rowCount()):
                    ep_item = season_item.child(ep_row, 0)
                    file_id = ep_item.data(ROLE_FILE_ID)
                    if file_id:
                        file_ids.append(file_id)
            def _action():
                api_failures = 0
                for file_id in file_ids:
                    try:
                        self.api.delete_episode_file(file_id)
                    except Exception:
                        api_failures += 1
                fs_deleted, fs_error = self._delete_tree_path(series_path)
                return {'api_failures': api_failures, 'fs_deleted': fs_deleted, 'fs_error': fs_error}
            def _on_success(result):
                msg = f'Deleted from disk: {label}'
                if result['api_failures']:
                    msg += f" ({result['api_failures']} API delete errors)"
                if not result['fs_deleted']:
                    msg += ' (disk delete failed)'
                    if result['fs_error']:
                        QMessageBox.warning(self, 'Delete Warning', f'Disk delete failed:\n{result["fs_error"]}')
                else:
                    parent = item.parent() or self.model.invisibleRootItem()
                    parent.removeRow(item.row())
                self.status_label.setText(msg)
            self._run_api_action(f'Deleting from disk: {label}...', _action, _on_success, 'Delete from disk failed')
            return
        if node_type == 'season':
            season_path = item.data(ROLE_SEASON_PATH)
            file_ids = []
            for row in range(item.rowCount()):
                ep_item = item.child(row, 0)
                file_id = ep_item.data(ROLE_FILE_ID)
                if file_id:
                    file_ids.append(file_id)
            def _action():
                api_failures = 0
                for file_id in file_ids:
                    try:
                        self.api.delete_episode_file(file_id)
                    except Exception:
                        api_failures += 1
                fs_deleted, fs_error = self._delete_tree_path(season_path)
                return {'api_failures': api_failures, 'fs_deleted': fs_deleted, 'fs_error': fs_error}
            def _on_success(result):
                msg = f'Deleted from disk: {label}'
                if result['api_failures']:
                    msg += f" ({result['api_failures']} API delete errors)"
                if not result['fs_deleted']:
                    msg += ' (disk delete failed)'
                    if result['fs_error']:
                        QMessageBox.warning(self, 'Delete Warning', f'Disk delete failed:\n{result["fs_error"]}')
                else:
                    parent = item.parent() or self.model.invisibleRootItem()
                    parent.removeRow(item.row())
                self.status_label.setText(msg)
            self._run_api_action(f'Deleting from disk: {label}...', _action, _on_success, 'Delete from disk failed')
            return
        if node_type == 'episode':
            file_id = item.data(ROLE_FILE_ID)
            file_path = item.data(ROLE_FILE_PATH)
            if not file_id and not file_path:
                self.status_label.setText('No file found for this episode')
                return
            def _action():
                api_failures = 0
                if file_id:
                    try:
                        self.api.delete_episode_file(file_id)
                    except Exception:
                        api_failures += 1
                fs_deleted, fs_error = self._delete_file_path(file_path)
                return {'api_failures': api_failures, 'fs_deleted': fs_deleted, 'fs_error': fs_error}
            def _on_success(result):
                msg = f'Deleted from disk: {label}'
                if result['api_failures']:
                    msg += f" ({result['api_failures']} API delete errors)"
                if not result['fs_deleted']:
                    msg += ' (disk delete failed)'
                    if result['fs_error']:
                        QMessageBox.warning(self, 'Delete Warning', f'Disk delete failed:\n{result["fs_error"]}')
                else:
                    parent = item.parent() or self.model.invisibleRootItem()
                    parent.removeRow(item.row())
                self.status_label.setText(msg)
            self._run_api_action(f'Deleting from disk: {label}...', _action, _on_success, 'Delete from disk failed')
    def _ctx_change_quality_profile(self, item: QStandardItem):
        """Change quality profile for a series via a combo-box dialog."""
        node_type = item.data(ROLE_NODE_TYPE)
        if node_type != 'series':
            self.status_label.setText('Quality profile can only be changed on a series')
            return
        series_id = item.data(ROLE_SERIES_ID)
        if not self._quality_profiles:
            QMessageBox.critical(self, 'Error', 'No quality profiles available.')
            return
        label = item.text()
        def _fetch_action():
            return self.api.get_series_by_id(series_id)
        def _on_fetched(series_data: dict):
            current_qp_id = series_data.get('qualityProfileId', 0)
            dlg = QDialog(self)
            dlg.setWindowTitle('Change Quality Profile')
            lay = QVBoxLayout(dlg)
            lay.addWidget(QLabel(f'Quality profile for: {label}'))
            combo = QComboBox()
            current_idx = 0
            for i, p in enumerate(self._quality_profiles):
                combo.addItem(p['name'], p['id'])
                if p['id'] == current_qp_id:
                    current_idx = i
            combo.setCurrentIndex(current_idx)
            lay.addWidget(combo)
            btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            btns.accepted.connect(dlg.accept)
            btns.rejected.connect(dlg.reject)
            lay.addWidget(btns)
            if dlg.exec() != QDialog.Accepted:
                self.status_label.setText('Quality profile change cancelled')
                return
            new_qp_id = combo.currentData()
            if new_qp_id == current_qp_id:
                self.status_label.setText('Quality profile unchanged')
                return
            selected_name = combo.currentText()
            def _update_action():
                payload = dict(series_data)
                payload['qualityProfileId'] = new_qp_id
                self.api.update_series(payload)
                return None
            def _on_updated(_result):
                row_idx = item.index().row()
                parent = item.parent() or self.model.invisibleRootItem()
                qp_item = parent.child(row_idx, 3)
                if qp_item:
                    qp_item.setText(selected_name)
                self.status_label.setText(f'Quality profile changed to {selected_name} for {label}')
            self._run_api_action(
                f'Updating quality profile: {label}...',
                _update_action,
                _on_updated,
                'Failed to update quality profile',
            )
        self._run_api_action(
            f'Loading quality profile options: {label}...',
            _fetch_action,
            _on_fetched,
            'Failed to fetch series',
        )
    def _ctx_unmonitor_delete(self, item: QStandardItem, node_type: str):
        series_id = item.data(ROLE_SERIES_ID)
        label = item.text()
        reply = QMessageBox.question(
            self, 'Unmonitor & Delete',
            f'Unmonitor "{label}" and delete files from disk?\n\nThis cannot be undone.',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return
        if node_type == 'series':
            def _action():
                self.api.delete_series(series_id, delete_files=True)
                return {'api_failures': 0}
            def _on_success(_result):
                parent = item.parent() or self.model.invisibleRootItem()
                parent.removeRow(item.row())
                self.status_label.setText(f'Deleted & unmonitored: {label}')
            self._run_api_action(f'Deleting & unmonitoring: {label}...', _action, _on_success, 'Unmonitor & delete failed')
            return
        if node_type == 'season':
            season_num = item.data(ROLE_SEASON_NUM)
            season_path = item.data(ROLE_SEASON_PATH)
            file_ids: list[int] = []
            for row in range(item.rowCount()):
                ep_item = item.child(row, 0)
                file_id = ep_item.data(ROLE_FILE_ID)
                if file_id:
                    file_ids.append(file_id)
            season_episodes = self._collect_episode_payloads(item)
            def _action():
                api_failures = 0
                for file_id in file_ids:
                    try:
                        self.api.delete_episode_file(file_id)
                    except Exception:
                        api_failures += 1
                for ep in season_episodes:
                    ep['monitored'] = False
                    try:
                        self.api.update_episode(ep)
                    except Exception:
                        api_failures += 1
                try:
                    series_data = self.api.get_series_by_id(series_id)
                    for s in series_data.get('seasons', []):
                        if s.get('seasonNumber') == season_num:
                            s['monitored'] = False
                            break
                    self.api.update_series(series_data)
                except Exception:
                    api_failures += 1
                fs_deleted, fs_error = self._delete_tree_path(season_path)
                return {'api_failures': api_failures, 'fs_deleted': fs_deleted, 'fs_error': fs_error}
            def _on_success(result):
                msg = f'Deleted & unmonitored: {label}'
                if result['api_failures']:
                    msg += f" ({result['api_failures']} API errors)"
                if not result['fs_deleted']:
                    msg += ' (disk delete failed)'
                    if result['fs_error']:
                        QMessageBox.warning(self, 'Delete Warning', f'Disk delete failed:\n{result["fs_error"]}')
                    self._update_mon_column(item, False)
                    for row in range(item.rowCount()):
                        ep_item = item.child(row, 0)
                        self._update_mon_column(ep_item, False)
                else:
                    parent = item.parent() or self.model.invisibleRootItem()
                    parent.removeRow(item.row())
                self.status_label.setText(msg)
            self._run_api_action(f'Deleting & unmonitoring: {label}...', _action, _on_success, 'Unmonitor & delete failed')
            return
        if node_type == 'episode':
            file_id = item.data(ROLE_FILE_ID)
            file_path = item.data(ROLE_FILE_PATH)
            ep_payloads = self._clone_episode_payloads(item.data(ROLE_EPISODE_DATA))
            def _action():
                api_failures = 0
                for ep in ep_payloads:
                    ep['monitored'] = False
                    try:
                        self.api.update_episode(ep)
                    except Exception:
                        api_failures += 1
                if file_id:
                    try:
                        self.api.delete_episode_file(file_id)
                    except Exception:
                        api_failures += 1
                fs_deleted, fs_error = self._delete_file_path(file_path)
                return {'api_failures': api_failures, 'fs_deleted': fs_deleted, 'fs_error': fs_error}
            def _on_success(result):
                msg = f'Deleted & unmonitored: {label}'
                if result['api_failures']:
                    msg += f" ({result['api_failures']} API errors)"
                if not result['fs_deleted']:
                    msg += ' (disk delete failed)'
                    if result['fs_error']:
                        QMessageBox.warning(self, 'Delete Warning', f'Disk delete failed:\n{result["fs_error"]}')
                    self._update_mon_column(item, False)
                else:
                    parent = item.parent() or self.model.invisibleRootItem()
                    parent.removeRow(item.row())
                self.status_label.setText(msg)
            self._run_api_action(f'Deleting & unmonitoring: {label}...', _action, _on_success, 'Unmonitor & delete failed')
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
        node_type = item.data(ROLE_NODE_TYPE)
        if node_type == 'series':
            path = item.data(ROLE_SERIES_PATH)
            if path and os.path.isdir(path):
                _open_path(path)
            else:
                self.status_label.setText(f'Directory not found: {path}')
        elif node_type == 'season':
            path = item.data(ROLE_SEASON_PATH)
            if path and os.path.isdir(path):
                _open_path(path)
            else:
                self.status_label.setText(f'Directory not found: {path}')
        elif node_type == 'episode':
            path = item.data(ROLE_FILE_PATH)
            if path and os.path.isfile(path):
                _open_path(path)
            else:
                self.status_label.setText(f'File not found: {path}')

    def _on_delete(self):
        if not self._ensure_action_idle():
            return
        item = self._current_item()
        if not item:
            return
        node_type = item.data(ROLE_NODE_TYPE)
        if node_type == 'series':
            self._delete_series(item)
        elif node_type == 'season':
            self._delete_season(item)

    def _delete_series(self, item: QStandardItem):
        series_id = item.data(ROLE_SERIES_ID)
        title = item.text()
        reply = QMessageBox.question(
            self, 'Delete Series',
            f'Delete "{title}" from Sonarr AND from disk?\n\nThis cannot be undone.',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        def _action():
            self.api.delete_series(series_id, delete_files=True)
            return None

        def _on_success(_result):
            parent = item.parent() or self.model.invisibleRootItem()
            parent.removeRow(item.row())
            self.status_label.setText(f'Deleted series: {title}')

        self._run_api_action(f'Deleting series: {title}...', _action, _on_success, 'Failed to delete series')

    def _delete_season(self, item: QStandardItem):
        # reuse the unmonitor & delete logic
        self._ctx_unmonitor_delete(item, 'season')

    # ── Add Show ───────────────────────────────────────────────

    def _add_show(self):
        if not self._ensure_action_idle():
            return
        dlg = AddShowDialog(self, self.api)
        if dlg.exec() == QDialog.Accepted and dlg.added_series:
            self._refresh()

    def closeEvent(self, event):
        """Save column widths and stop worker before closing."""
        widths = [self.tree.columnWidth(c) for c in range(len(self._columns))]
        self._settings.setValue('column_widths', widths)
        self._settings.sync()
        if not self._stop_action_worker(5000):
            QMessageBox.warning(
                self,
                'Action still running',
                'An action is still in progress.\nPlease try closing again in a few seconds.',
            )
            event.ignore()
            return
        if not self._stop_worker(5000):
            reply = QMessageBox.question(
                self,
                'Background task still running',
                'A refresh is still in progress and did not stop in time.\n'
                'Force close now?',
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                event.ignore()
                return
            worker = self.worker
            if worker and worker.isRunning():
                worker.terminate()
                worker.wait(1500)
                if worker.isRunning():
                    QMessageBox.critical(
                        self,
                        'Unable to close safely',
                        'Background thread is still running after force-close attempt.\n'
                        'Please try closing again in a few seconds.',
                    )
                    event.ignore()
                    return
            self.worker = None
        self.action_worker = None
        super().closeEvent(event)

    def _clear_cache_and_refresh(self):
        """Clear the ffprobe cache and reload all data."""
        if not self._ensure_action_idle('Action still running; try again in a moment'):
            return
        if not self._stop_worker(5000):
            QMessageBox.warning(
                self,
                'Refresh still running',
                'Cannot clear cache while a refresh is still running.\n'
                'Please wait and try again.',
            )
            self.status_label.setText('Refresh still running; try again in a moment')
            return
        clear_probe_cache()
        self.status_label.setText('Cache cleared, refreshing...')
        self._refresh()

    def _refresh(self):
        """Reload all data from Sonarr."""
        if not self._ensure_action_idle('Action still running; try again'):
            return
        if not self._stop_worker(5000):
            QMessageBox.warning(
                self,
                'Refresh still running',
                'Previous refresh is still stopping.\n'
                'Please try again in a few seconds.',
            )
            self.status_label.setText('Previous refresh still stopping; try again')
            return
        self.model.removeRows(0, self.model.rowCount())
        self.progress_bar.show()
        self.progress_bar.setRange(0, 0)
        self.status_label.setText('Refreshing...')
        self._start_worker()
