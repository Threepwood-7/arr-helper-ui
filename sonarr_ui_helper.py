#!/usr/bin/env python3
"""
Sonarr UI Helper - PySide6 TreeView for browsing Sonarr series/episodes.
Uses ffprobe for media info and shares config with media_quality_checker.
"""

import os
import sys
import json
import platform
import subprocess
import shutil
from pathlib import Path
from typing import Dict, List, Optional

import requests
import tomli
from ffprobe_utils import find_ffprobe
from PySide6.QtCore import Qt, QThread, Signal, QModelIndex, QSortFilterProxyModel, QSettings, QStringListModel
from PySide6.QtGui import QStandardItemModel, QStandardItem, QFont, QColor, QKeySequence, QShortcut, QAction
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QTreeView, QVBoxLayout, QWidget,
    QHeaderView, QMessageBox, QProgressBar, QLabel, QHBoxLayout,
    QStatusBar, QPushButton, QMenu, QDialog, QLineEdit,
    QDialogButtonBox, QAbstractItemView, QListWidget,
    QListWidgetItem, QCheckBox, QMenuBar, QComboBox, QCompleter,
)


# ── Sonarr API helpers ──────────────────────────────────────────────

class SonarrAPI:
    def __init__(self, url: str, api_key: str, http_user: str = '', http_pass: str = ''):
        self.url = url.rstrip('/')
        self.api_key = api_key
        self.headers = {'X-Api-Key': api_key}
        self.auth = (http_user, http_pass) if http_user else None

    def _get(self, endpoint: str):
        r = requests.get(f"{self.url}/api/v3/{endpoint}", headers=self.headers, auth=self.auth, timeout=300)
        r.raise_for_status()
        return r.json()

    def _post(self, endpoint: str, data: dict):
        r = requests.post(f"{self.url}/api/v3/{endpoint}", headers=self.headers, auth=self.auth, json=data, timeout=300)
        r.raise_for_status()
        return r.json() if r.text else {}

    def _delete(self, endpoint: str, params: dict = None):
        r = requests.delete(f"{self.url}/api/v3/{endpoint}", headers=self.headers, auth=self.auth, params=params or {}, timeout=300)
        r.raise_for_status()
        return r

    def _put(self, endpoint: str, data: dict):
        r = requests.put(f"{self.url}/api/v3/{endpoint}", headers=self.headers, auth=self.auth, json=data, timeout=300)
        r.raise_for_status()
        return r.json()

    def get_series(self) -> List[dict]:
        return self._get('series')

    def get_series_by_id(self, series_id: int) -> dict:
        return self._get(f'series/{series_id}')

    def get_episodes(self, series_id: int) -> List[dict]:
        return self._get(f'episode?seriesId={series_id}')

    def get_episode_files(self, series_id: int) -> List[dict]:
        return self._get(f'episodefile?seriesId={series_id}')

    def update_series(self, series: dict) -> dict:
        return self._put(f'series/{series["id"]}', series)

    def update_episode(self, episode: dict) -> dict:
        return self._put(f'episode/{episode["id"]}', episode)

    def delete_series(self, series_id: int, delete_files: bool = True):
        self._delete(f'series/{series_id}', {'deleteFiles': str(delete_files).lower()})

    def delete_episode_file(self, file_id: int):
        self._delete(f'episodefile/{file_id}')

    # search / commands
    def command(self, body: dict) -> dict:
        return self._post('command', body)

    def series_search(self, series_id: int):
        return self.command({'name': 'SeriesSearch', 'seriesId': series_id})

    def season_search(self, series_id: int, season_number: int):
        return self.command({'name': 'SeasonSearch', 'seriesId': series_id, 'seasonNumber': season_number})

    def episode_search(self, episode_ids: List[int]):
        return self.command({'name': 'EpisodeSearch', 'episodeIds': episode_ids})

    def get_release(self, episode_id: int) -> List[dict]:
        return self._get(f'release?episodeId={episode_id}')

    def download_release(self, guid: str, indexer_id: int) -> dict:
        return self._post('release', {'guid': guid, 'indexerId': indexer_id})

    # lookup / add
    def lookup_series(self, term: str) -> List[dict]:
        return self._get(f'series/lookup?term={requests.utils.quote(term)}')

    def add_series(self, series: dict) -> dict:
        return self._post('series', series)

    def get_root_folders(self) -> List[dict]:
        return self._get('rootfolder')

    def get_quality_profiles(self) -> List[dict]:
        return self._get('qualityprofile')


# ── ffprobe cache & helper ──────────────────────────────────────────

_FFPROBE: str | None = None          # resolved at startup in main()
_PROBE_CACHE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'z_fprobe.cache')
_probe_cache: Dict = {}


def _load_probe_cache():
    global _probe_cache
    if os.path.exists(_PROBE_CACHE_PATH):
        try:
            with open(_PROBE_CACHE_PATH, 'r') as f:
                _probe_cache = json.load(f)
        except Exception:
            _probe_cache = {}


def _save_probe_cache():
    try:
        with open(_PROBE_CACHE_PATH, 'w') as f:
            json.dump(_probe_cache, f, indent=1)
    except Exception:
        pass


def probe_file(file_path: str) -> Dict:
    """Return dict with codecs, resolution, bitrates, HDR, languages, size."""
    info: Dict = {
        'video_codec': '',
        'video_resolution': '',
        'video_bitrate': '',
        'audio_codec': '',
        'audio_bitrate': '',
        'hdr': '',
        'audio_langs': [],
        'sub_langs': [],
        'size_bytes': 0,
    }
    try:
        info['size_bytes'] = os.path.getsize(file_path)
    except OSError:
        pass

    # check cache — keyed by path, invalidated if size changed or fields missing
    cached = _probe_cache.get(file_path)
    if cached and cached.get('size_bytes') == info['size_bytes'] and 'video_resolution' in cached:
        return cached

    try:
        cmd = [
            _FFPROBE or 'ffprobe', '-v', 'quiet', '-print_format', 'json',
            '-show_streams', '-show_format', file_path,
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        if result.returncode != 0:
            return info
        data = json.loads(result.stdout)
        for s in data.get('streams', []):
            codec_type = s.get('codec_type', '')
            lang = s.get('tags', {}).get('language', '')
            if codec_type == 'video' and not info['video_codec']:
                info['video_codec'] = s.get('codec_name', '').upper()
                w = s.get('width', 0)
                h = s.get('height', 0)
                if w and h:
                    info['video_resolution'] = f'{w}x{h}'
                vbr = int(s.get('bit_rate', 0) or 0)
                if vbr:
                    info['video_bitrate'] = f'{vbr // 1000} kbps'
                color_transfer = s.get('color_transfer', '')
                color_space = s.get('color_space', '')
                side_data = s.get('side_data_list', [])
                has_hdr_transfer = color_transfer in ('smpte2084', 'arib-std-b67')
                has_hdr_space = color_space in ('bt2020nc', 'bt2020c')
                has_dovi = any(
                    sd.get('side_data_type', '') in ('DOVI configuration record', 'Dolby Vision configuration')
                    for sd in side_data
                ) if side_data else False
                if has_dovi:
                    info['hdr'] = 'DV'
                elif has_hdr_transfer or has_hdr_space:
                    info['hdr'] = 'HDR'
                else:
                    info['hdr'] = 'SDR'
            elif codec_type == 'audio':
                if not info['audio_codec']:
                    info['audio_codec'] = s.get('codec_name', '').upper()
                    abr = int(s.get('bit_rate', 0) or 0)
                    if abr:
                        info['audio_bitrate'] = f'{abr // 1000} kbps'
                if lang:
                    info['audio_langs'].append(lang)
            elif codec_type == 'subtitle':
                if lang:
                    info['sub_langs'].append(lang)

        if not info['video_bitrate']:
            fmt_br = int(data.get('format', {}).get('bit_rate', 0) or 0)
            if fmt_br:
                info['video_bitrate'] = f'{fmt_br // 1000} kbps'
    except Exception:
        pass

    _probe_cache[file_path] = info
    return info


def _open_path(path: str):
    """Open a file or directory with the system default handler (cross-platform)."""
    system = platform.system()
    if system == 'Windows':
        os.startfile(path)
    elif system == 'Darwin':
        subprocess.Popen(['open', path])
    else:
        subprocess.Popen(['xdg-open', path])


def fmt_size(size_bytes: int) -> str:
    if size_bytes <= 0:
        return '0 GB'
    gb = size_bytes / (1024 ** 3)
    if gb >= 1:
        return f'{gb:.2f} GB'
    mb = size_bytes / (1024 ** 2)
    return f'{mb:.0f} MB'


# ── Background loader ──────────────────────────────────────────────

class LoadWorker(QThread):
    """Fetches series + episode data from Sonarr, probes files, emits results."""
    progress = Signal(str)
    series_ready = Signal(list)

    def __init__(self, api: SonarrAPI):
        super().__init__()
        self.api = api

    def run(self):
        self.progress.emit('Fetching series list…')
        try:
            all_series = self.api.get_series()
        except Exception as e:
            self.progress.emit(f'Error: {e}')
            self.series_ready.emit([])
            return

        result = []
        for idx, series in enumerate(all_series):
            if self.isInterruptionRequested():
                return
            series_id = series['id']
            title = series.get('title', '?')
            year = series.get('year', '')
            series_path = series.get('path', '')

            self.progress.emit(f'Loading {idx + 1}/{len(all_series)}: {title}')

            try:
                ep_files = self.api.get_episode_files(series_id)
            except Exception:
                ep_files = []

            try:
                episodes = self.api.get_episodes(series_id)
            except Exception:
                episodes = []

            # map episode_file_id -> episode metadata
            file_to_eps: Dict[int, list] = {}
            for ep in episodes:
                fid = ep.get('episodeFileId', 0)
                if fid:
                    file_to_eps.setdefault(fid, []).append(ep)

            # group downloaded episodes by season
            seasons: Dict[int, list] = {}
            for ef in ep_files:
                if self.isInterruptionRequested():
                    return
                file_path = ef.get('path', '')
                file_id = ef.get('id')
                eps_for_file = file_to_eps.get(file_id, [])
                season_num = eps_for_file[0].get('seasonNumber', 0) if eps_for_file else 0
                ep_num = eps_for_file[0].get('episodeNumber', 0) if eps_for_file else 0

                probe = probe_file(file_path)

                ep_entry = {
                    'episode_number': ep_num,
                    'file_name': Path(file_path).name if file_path else '',
                    'file_path': file_path,
                    'file_id': file_id,
                    'size_bytes': probe['size_bytes'],
                    'video_resolution': probe['video_resolution'],
                    'video_bitrate': probe['video_bitrate'],
                    'video_codec': probe['video_codec'],
                    'hdr': probe['hdr'],
                    'audio_codec': probe['audio_codec'],
                    'audio_bitrate': probe['audio_bitrate'],
                    'audio_langs': ', '.join(dict.fromkeys(probe['audio_langs'])),
                    'sub_langs': ', '.join(dict.fromkeys(probe['sub_langs'])),
                    'episode_data': eps_for_file,
                }
                seasons.setdefault(season_num, []).append(ep_entry)

            # sort episodes inside each season
            for sn in seasons:
                seasons[sn].sort(key=lambda e: e['episode_number'])

            # collect missing episodes per season (no file)
            missing_seasons: Dict[int, list] = {}
            for ep in episodes:
                if ep.get('episodeFileId', 0) == 0:
                    sn = ep.get('seasonNumber', 0)
                    missing_seasons.setdefault(sn, []).append(ep)
            for sn in missing_seasons:
                missing_seasons[sn].sort(key=lambda e: e.get('episodeNumber', 0))

            # collect all season numbers from the series metadata
            all_season_nums = set()
            for s_info in series.get('seasons', []):
                all_season_nums.add(s_info.get('seasonNumber', 0))
            # also include seasons from episodes
            for ep in episodes:
                all_season_nums.add(ep.get('seasonNumber', 0))

            total_size = sum(e['size_bytes'] for s in seasons.values() for e in s)

            result.append({
                'series_id': series_id,
                'title': title,
                'year': year,
                'path': series_path,
                'total_size': total_size,
                'seasons': seasons,
                'missing_seasons': missing_seasons,
                'all_season_nums': sorted(all_season_nums),
                'series_data': series,
            })

        result.sort(key=lambda s: s['title'].lower())
        _save_probe_cache()
        self.progress.emit('Done')
        self.series_ready.emit(result)


# ── Custom data roles ──────────────────────────────────────────────

ROLE_NODE_TYPE = Qt.UserRole + 1    # 'series' | 'season' | 'episode'
ROLE_SERIES_ID = Qt.UserRole + 2
ROLE_SERIES_PATH = Qt.UserRole + 3
ROLE_SEASON_NUM = Qt.UserRole + 4
ROLE_SEASON_PATH = Qt.UserRole + 5
ROLE_FILE_PATH = Qt.UserRole + 6
ROLE_FILE_ID = Qt.UserRole + 7
ROLE_EPISODE_DATA = Qt.UserRole + 8
ROLE_IS_MISSING = Qt.UserRole + 9


# ── Manual Search Dialog ──────────────────────────────────────────

ROLE_RELEASE = Qt.UserRole + 20  # stores the release dict on each row


class _SearchFilterProxy(QSortFilterProxyModel):
    """Filters on title text, quality, and indexer; sorts numerically via UserRole."""

    def __init__(self, col_quality=2, col_indexer=3, parent=None):
        super().__init__(parent)
        self._quality = ''
        self._indexer = ''
        self._col_quality = col_quality
        self._col_indexer = col_indexer

    def set_quality(self, quality: str):
        self._quality = quality
        self.invalidateFilter()

    def set_indexer(self, indexer: str):
        self._indexer = indexer
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row: int, source_parent: QModelIndex) -> bool:
        if not super().filterAcceptsRow(source_row, source_parent):
            return False
        model = self.sourceModel()
        if self._quality:
            idx = model.index(source_row, self._col_quality, source_parent)
            if (model.data(idx, Qt.DisplayRole) or '') != self._quality:
                return False
        if self._indexer:
            idx = model.index(source_row, self._col_indexer, source_parent)
            if (model.data(idx, Qt.DisplayRole) or '') != self._indexer:
                return False
        return True

    def lessThan(self, left: QModelIndex, right: QModelIndex) -> bool:
        lv = left.data(Qt.UserRole)
        rv = right.data(Qt.UserRole)
        if lv is not None and rv is not None:
            try:
                return float(lv) < float(rv)
            except (TypeError, ValueError):
                pass
        return super().lessThan(left, right)


class ManualSearchDialog(QDialog):
    """Shows release search results with sorting and filtering."""

    _COL_QUALITY = 2  # index of Quality column for filtering
    _COL_INDEXER = 3  # index of Indexer column for filtering

    def __init__(self, parent, title: str, releases: List[dict], settings: QSettings = None):
        super().__init__(parent)
        self.setWindowTitle(f'Manual Search: {title}')
        self.setWindowState(Qt.WindowMaximized)
        self.selected_release = None
        self._settings = settings

        layout = QVBoxLayout(self)

        # filter row: text filter + quality combo
        filter_row = QHBoxLayout()
        self.filter_input = QLineEdit()
        self.filter_input.setPlaceholderText('Filter results…')
        self.filter_input.textChanged.connect(self._apply_filters)
        # autocomplete from saved history
        self._filter_history = []
        if settings:
            self._filter_history = settings.value('search_filter_history', []) or []
        self._completer_model = QStringListModel(self._filter_history)
        completer = QCompleter(self._completer_model, self)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        self.filter_input.setCompleter(completer)
        filter_row.addWidget(self.filter_input, 1)

        filter_row.addWidget(QLabel('Quality:'))
        self.quality_combo = QComboBox()
        self.quality_combo.addItem('All', '')
        # collect unique quality names
        qualities = sorted({
            rel.get('quality', {}).get('quality', {}).get('name', '')
            for rel in releases
        } - {''})
        for q in qualities:
            self.quality_combo.addItem(q, q)
        self.quality_combo.currentIndexChanged.connect(self._apply_filters)
        filter_row.addWidget(self.quality_combo)

        filter_row.addWidget(QLabel('Indexer:'))
        self.indexer_combo = QComboBox()
        self.indexer_combo.addItem('All', '')
        indexers = sorted({rel.get('indexer', '') for rel in releases} - {''})
        for ix in indexers:
            self.indexer_combo.addItem(ix, ix)
        self.indexer_combo.currentIndexChanged.connect(self._apply_filters)
        filter_row.addWidget(self.indexer_combo)
        layout.addLayout(filter_row)

        # source model
        self._columns = ['Title', 'Size (GB)', 'Quality', 'Indexer', 'Age']
        self.source_model = QStandardItemModel()
        self.source_model.setHorizontalHeaderLabels(self._columns)

        # populate
        for rel in releases:
            title_item = QStandardItem(rel.get('title', ''))
            title_item.setEditable(False)
            title_item.setData(rel, ROLE_RELEASE)

            size = rel.get('size', 0)
            size_gb = size / (1024**3) if size > 0 else 0.0
            size_item = QStandardItem()
            size_item.setEditable(False)
            size_item.setData(round(size_gb, 2), Qt.DisplayRole)

            quality = rel.get('quality', {}).get('quality', {}).get('name', '')
            quality_item = QStandardItem(quality)
            quality_item.setEditable(False)

            indexer_item = QStandardItem(rel.get('indexer', ''))
            indexer_item.setEditable(False)

            age_hours = rel.get('ageHours', 0) or 0
            if age_hours:
                age_days = age_hours / 24
            else:
                age_days = rel.get('age', 0) or 0
                age_hours = age_days * 24
            if age_hours < 1:
                age_str = '< 1h'
            elif age_hours < 24:
                age_str = f'{int(age_hours)}h'
            else:
                age_str = f'{int(age_days)}d'
            age_item = QStandardItem(age_str)
            age_item.setEditable(False)
            age_item.setData(age_hours, Qt.UserRole)

            self.source_model.appendRow([title_item, size_item, quality_item,
                                         indexer_item, age_item])

        # proxy for sorting + filtering
        self.proxy = _SearchFilterProxy(self._COL_QUALITY, self._COL_INDEXER)
        self.proxy.setSourceModel(self.source_model)
        self.proxy.setFilterCaseSensitivity(Qt.CaseInsensitive)
        self.proxy.setFilterKeyColumn(0)  # filter on title

        # tree view (flat table mode)
        self.table = QTreeView()
        self.table.setModel(self.proxy)
        self.table.setRootIsDecorated(False)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setSortingEnabled(True)
        self.table.sortByColumn(4, Qt.AscendingOrder)  # default sort by age asc (newest first)
        header = self.table.header()
        header.setSectionResizeMode(QHeaderView.Interactive)
        saved = settings.value('search_column_widths') if settings else None
        if saved and len(saved) == len(self._columns):
            for col, w in enumerate(saved):
                self.table.setColumnWidth(col, int(w))
        else:
            self.table.setColumnWidth(0, 800)
            for col in range(1, len(self._columns)):
                self.table.resizeColumnToContents(col)
        self.table.doubleClicked.connect(self._on_double_click)

        layout.addWidget(self.table)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._on_ok)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _apply_filters(self):
        self.proxy.setFilterFixedString(self.filter_input.text())
        self.proxy.set_quality(self.quality_combo.currentData() or '')
        self.proxy.set_indexer(self.indexer_combo.currentData() or '')

    def _save_filter_history(self):
        text = self.filter_input.text().strip()
        if text and self._settings:
            if text not in self._filter_history:
                self._filter_history.append(text)
                # keep last 50
                self._filter_history = self._filter_history[-50:]
                self._settings.setValue('search_filter_history', self._filter_history)
                self._completer_model.setStringList(self._filter_history)

    def _save_column_widths(self):
        if self._settings:
            widths = [self.table.columnWidth(c) for c in range(len(self._columns))]
            self._settings.setValue('search_column_widths', widths)

    def _get_selected_release(self) -> Optional[dict]:
        indexes = self.table.selectionModel().selectedRows()
        if not indexes:
            return None
        source_idx = self.proxy.mapToSource(indexes[0])
        item = self.source_model.item(source_idx.row(), 0)
        return item.data(ROLE_RELEASE) if item else None

    def _on_double_click(self, _index: QModelIndex):
        rel = self._get_selected_release()
        if rel:
            self.selected_release = rel
            self._save_filter_history()
            self._save_column_widths()
            super().accept()

    def _on_ok(self):
        rel = self._get_selected_release()
        if rel:
            self.selected_release = rel
            self._save_filter_history()
            self._save_column_widths()
            super().accept()

    def reject(self):
        self._save_filter_history()
        self._save_column_widths()
        super().reject()


# ── Add Show Dialog ───────────────────────────────────────────────

class AddShowDialog(QDialog):
    """Search for a show on Sonarr and add it."""

    def __init__(self, parent, api: SonarrAPI):
        super().__init__(parent)
        self.api = api
        self.setWindowTitle('Add Show')
        pw = parent.width() if parent else 1000
        self.resize(int(pw * 0.7), 500)
        self.added_series = None

        layout = QVBoxLayout(self)

        # search bar
        search_row = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText('Search for a show…')
        self.search_input.returnPressed.connect(self._do_search)
        btn_search = QPushButton('Search')
        btn_search.clicked.connect(self._do_search)
        search_row.addWidget(self.search_input)
        search_row.addWidget(btn_search)
        layout.addLayout(search_row)

        # root folder and quality profile selectors
        options_row = QHBoxLayout()
        options_row.addWidget(QLabel('Root Folder:'))
        self.combo_root = QComboBox()
        options_row.addWidget(self.combo_root, 1)
        options_row.addWidget(QLabel('Quality Profile:'))
        self.combo_qp = QComboBox()
        options_row.addWidget(self.combo_qp, 1)
        layout.addLayout(options_row)

        # preload root folders and quality profiles
        self._root_folders = []
        self._quality_profiles = []
        try:
            self._root_folders = api.get_root_folders()
            for rf in self._root_folders:
                self.combo_root.addItem(rf['path'], rf['path'])
        except Exception:
            pass
        try:
            self._quality_profiles = api.get_quality_profiles()
            for qp in self._quality_profiles:
                self.combo_qp.addItem(qp['name'], qp['id'])
        except Exception:
            pass

        # results list
        self.results_list = QListWidget()
        self.results_list.setAlternatingRowColors(True)
        layout.addWidget(self.results_list)

        # add button
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText('Add Selected')
        buttons.accepted.connect(self._add_selected)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.lookup_results: List[dict] = []

    def _do_search(self):
        term = self.search_input.text().strip()
        if not term:
            return
        self.results_list.clear()
        self.lookup_results = []
        try:
            results = self.api.lookup_series(term)
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Search failed:\n{e}')
            return

        self.lookup_results = results
        for idx, r in enumerate(results):
            title = r.get('title', '?')
            year = r.get('year', '')
            overview = (r.get('overview', '') or '')[:120]
            status = r.get('status', '')
            already = 'id' in r and r.get('id', 0) > 0
            label = f"{title} ({year}) [{status}]"
            if already:
                label += ' [ALREADY IN SONARR]'
            if overview:
                label += f"\n  {overview}"
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, idx)
            self.results_list.addItem(item)

    def _add_selected(self):
        row = self.results_list.currentRow()
        if row < 0 or row >= len(self.lookup_results):
            QMessageBox.warning(self, 'No Selection', 'Please select a show first.')
            return

        lookup = self.lookup_results[row]

        # check if already added
        if lookup.get('id', 0) > 0:
            QMessageBox.information(self, 'Already Added', 'This series is already in Sonarr.')
            return

        root_path = self.combo_root.currentData()
        qp_id = self.combo_qp.currentData()
        if not root_path:
            QMessageBox.critical(self, 'Error', 'No root folder selected.')
            return
        if qp_id is None:
            QMessageBox.critical(self, 'Error', 'No quality profile selected.')
            return

        # build the payload from lookup data with required fields
        series_data = dict(lookup)  # copy so we don't mutate
        series_data['rootFolderPath'] = root_path
        series_data['qualityProfileId'] = qp_id
        series_data['monitored'] = False
        series_data['seasonFolder'] = True
        series_data['addOptions'] = {'searchForMissingEpisodes': False}
        # remove id=0 so Sonarr treats it as a new series
        series_data.pop('id', None)

        try:
            result = self.api.add_series(series_data)
            self.added_series = result
            QMessageBox.information(self, 'Added', f"Added: {result.get('title', '?')}")
            self.accept()
        except requests.exceptions.HTTPError as e:
            body = ''
            if e.response is not None:
                try:
                    body = e.response.text
                except Exception:
                    pass
            QMessageBox.critical(self, 'Error', f'Failed to add series:\n{e}\n\n{body}')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to add series:\n{e}')


# ── Main window ────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self, api: SonarrAPI, settings: dict = None):
        super().__init__()
        self.api = api
        self.cfg = settings or {}
        self._settings = QSettings('SonarrUIHelper', 'SonarrUIHelper')
        self.setWindowTitle('Sonarr UI Helper')
        self.resize(1600, 800)

        # ── Menu bar ──────────────────────────────────────────
        self._build_menu_bar()

        # central widget
        central = QWidget()
        layout = QVBoxLayout(central)
        layout.setContentsMargins(4, 4, 4, 4)

        # toolbar (with mnemonics via &)
        toolbar = QHBoxLayout()
        btn_expand_all = QPushButton('E&xpand All')
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
        self.worker = LoadWorker(api)
        self.worker.progress.connect(self._on_progress)
        self.worker.series_ready.connect(self._on_data_loaded)
        self.worker.start()

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
        act.setShortcut(QKeySequence('Ctrl+Q'))
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

        # Help menu
        help_menu = mb.addMenu('&Help')
        act = help_menu.addAction('&Keyboard Shortcuts')
        act.setShortcut(QKeySequence('F1'))
        act.triggered.connect(self._show_help)

    def _toggle_missing_from_menu(self, checked: bool):
        """Sync the menu checkbox with the toolbar checkbox."""
        self.chk_show_missing.setChecked(checked)

    def _ctx_on_selected(self, action_name: str):
        """Dispatch a context-menu action on the currently selected tree item."""
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
            "<tr><td><code>Alt+X</code></td><td>Expand All</td></tr>"
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

    def _on_data_loaded(self, series_list: list):
        self.progress_bar.hide()
        self.model.removeRows(0, self.model.rowCount())

        if not series_list:
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
            # sensible defaults: first column wide, rest auto-fit
            self.tree.setColumnWidth(0, 600)
            for col in range(1, len(self._columns)):
                self.tree.resizeColumnToContents(col)

        self.status_label.setText(f'{total_series} series loaded')

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

    def _ctx_monitor(self, item: QStandardItem, node_type: str):
        series_id = item.data(ROLE_SERIES_ID)
        try:
            if node_type == 'series':
                series_data = self.api.get_series_by_id(series_id)
                series_data['monitored'] = True
                self.api.update_series(series_data)
                self._update_mon_column(item, True)
                self.status_label.setText(f'Monitored: {item.text()}')
            elif node_type == 'season':
                season_num = item.data(ROLE_SEASON_NUM)
                series_data = self.api.get_series_by_id(series_id)
                for s in series_data.get('seasons', []):
                    if s.get('seasonNumber') == season_num:
                        s['monitored'] = True
                        break
                self.api.update_series(series_data)
                self._update_mon_column(item, True)
                self.status_label.setText(f'Monitored: {item.text()}')
            elif node_type == 'episode':
                ep_data_list = item.data(ROLE_EPISODE_DATA) or []
                for ep in ep_data_list:
                    ep['monitored'] = True
                    self.api.update_episode(ep)
                self._update_mon_column(item, True)
                self.status_label.setText(f'Monitored: {item.text()}')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to monitor:\n{e}')

    def _ctx_unmonitor(self, item: QStandardItem, node_type: str):
        series_id = item.data(ROLE_SERIES_ID)
        try:
            if node_type == 'series':
                series_data = self.api.get_series_by_id(series_id)
                series_data['monitored'] = False
                self.api.update_series(series_data)
                self._update_mon_column(item, False)
                self.status_label.setText(f'Unmonitored: {item.text()}')
            elif node_type == 'season':
                season_num = item.data(ROLE_SEASON_NUM)
                series_data = self.api.get_series_by_id(series_id)
                for s in series_data.get('seasons', []):
                    if s.get('seasonNumber') == season_num:
                        s['monitored'] = False
                        break
                self.api.update_series(series_data)
                self._update_mon_column(item, False)
                # also unmonitor episodes in this season
                self._unmonitor_season_episodes(item)
                self.status_label.setText(f'Unmonitored: {item.text()}')
            elif node_type == 'episode':
                ep_data_list = item.data(ROLE_EPISODE_DATA) or []
                for ep in ep_data_list:
                    ep['monitored'] = False
                    self.api.update_episode(ep)
                self._update_mon_column(item, False)
                self.status_label.setText(f'Unmonitored: {item.text()}')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to unmonitor:\n{e}')

    def _unmonitor_season_episodes(self, season_item: QStandardItem):
        """Unmonitor all episodes under a season item."""
        for row in range(season_item.rowCount()):
            ep_item = season_item.child(row, 0)
            ep_data_list = ep_item.data(ROLE_EPISODE_DATA) or []
            for ep in ep_data_list:
                ep['monitored'] = False
                try:
                    self.api.update_episode(ep)
                except Exception:
                    pass
            self._update_mon_column(ep_item, False)

    def _ctx_auto_search(self, item: QStandardItem, node_type: str):
        series_id = item.data(ROLE_SERIES_ID)
        try:
            if node_type == 'series':
                self.api.series_search(series_id)
                self.status_label.setText(f'Auto search started: {item.text()}')
            elif node_type == 'season':
                season_num = item.data(ROLE_SEASON_NUM)
                self.api.season_search(series_id, season_num)
                self.status_label.setText(f'Auto search started: {item.text()}')
            elif node_type == 'episode':
                ep_data_list = item.data(ROLE_EPISODE_DATA) or []
                ep_ids = [ep['id'] for ep in ep_data_list if 'id' in ep]
                if ep_ids:
                    self.api.episode_search(ep_ids)
                    self.status_label.setText(f'Auto search started: {item.text()}')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Auto search failed:\n{e}')

    def _ctx_manual_search(self, item: QStandardItem, node_type: str):
        # get episode ids to search for
        ep_ids = []
        if node_type == 'episode':
            ep_data_list = item.data(ROLE_EPISODE_DATA) or []
            ep_ids = [ep['id'] for ep in ep_data_list if 'id' in ep]
        elif node_type == 'season':
            # collect all episode ids from child items
            for row in range(item.rowCount()):
                ep_item = item.child(row, 0)
                ep_data_list = ep_item.data(ROLE_EPISODE_DATA) or []
                ep_ids.extend(ep['id'] for ep in ep_data_list if 'id' in ep)
        elif node_type == 'series':
            # collect from all seasons -> all episodes
            for s_row in range(item.rowCount()):
                season_item = item.child(s_row, 0)
                for ep_row in range(season_item.rowCount()):
                    ep_item = season_item.child(ep_row, 0)
                    ep_data_list = ep_item.data(ROLE_EPISODE_DATA) or []
                    ep_ids.extend(ep['id'] for ep in ep_data_list if 'id' in ep)

        if not ep_ids:
            QMessageBox.warning(self, 'No Episodes', 'No episode IDs found for manual search.')
            return

        # search using the first episode id (Sonarr release endpoint is per-episode)
        self.status_label.setText('Searching for releases…')
        QApplication.setOverrideCursor(Qt.WaitCursor)
        QApplication.processEvents()
        try:
            releases = self.api.get_release(ep_ids[0])
        except Exception as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, 'Error', f'Manual search failed:\n{e}')
            self.status_label.setText('Search failed')
            return
        QApplication.restoreOverrideCursor()

        if not releases:
            QMessageBox.information(self, 'No Results', 'No releases found.')
            self.status_label.setText('No releases found')
            return

        dlg = ManualSearchDialog(self, item.text(), releases, settings=self._settings)
        if dlg.exec() == QDialog.Accepted and dlg.selected_release:
            rel = dlg.selected_release
            try:
                self.api.download_release(rel['guid'], rel['indexerId'])
                self.status_label.setText(f'Download queued: {rel.get("title", "?")}')
            except Exception as e:
                QMessageBox.critical(self, 'Error', f'Failed to queue download:\n{e}')
        else:
            self.status_label.setText('Manual search cancelled')

    def _ctx_delete_from_disk(self, item: QStandardItem, node_type: str):
        label = item.text()
        reply = QMessageBox.question(
            self, 'Delete from Disk',
            f'Delete "{label}" files from disk?\n\nMonitoring status will not change.\nThis cannot be undone.',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        try:
            if node_type == 'series':
                series_path = item.data(ROLE_SERIES_PATH)
                # delete episode files from Sonarr DB
                for s_row in range(item.rowCount()):
                    season_item = item.child(s_row, 0)
                    for ep_row in range(season_item.rowCount()):
                        ep_item = season_item.child(ep_row, 0)
                        file_id = ep_item.data(ROLE_FILE_ID)
                        if file_id:
                            try:
                                self.api.delete_episode_file(file_id)
                            except Exception:
                                pass
                if series_path and os.path.isdir(series_path):
                    shutil.rmtree(series_path, ignore_errors=True)
                parent = item.parent() or self.model.invisibleRootItem()
                parent.removeRow(item.row())
                self.status_label.setText(f'Deleted from disk: {label}')

            elif node_type == 'season':
                season_path = item.data(ROLE_SEASON_PATH)
                for row in range(item.rowCount()):
                    ep_item = item.child(row, 0)
                    file_id = ep_item.data(ROLE_FILE_ID)
                    if file_id:
                        try:
                            self.api.delete_episode_file(file_id)
                        except Exception:
                            pass
                if season_path and os.path.isdir(season_path):
                    shutil.rmtree(season_path, ignore_errors=True)
                parent = item.parent() or self.model.invisibleRootItem()
                parent.removeRow(item.row())
                self.status_label.setText(f'Deleted from disk: {label}')

            elif node_type == 'episode':
                file_id = item.data(ROLE_FILE_ID)
                file_path = item.data(ROLE_FILE_PATH)
                if file_id:
                    try:
                        self.api.delete_episode_file(file_id)
                    except Exception:
                        pass
                if file_path and os.path.isfile(file_path):
                    try:
                        os.remove(file_path)
                    except Exception:
                        pass
                parent = item.parent() or self.model.invisibleRootItem()
                parent.removeRow(item.row())
                self.status_label.setText(f'Deleted from disk: {label}')

        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed:\n{e}')

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

        try:
            series_data = self.api.get_series_by_id(series_id)
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to fetch series:\n{e}')
            return

        current_qp_id = series_data.get('qualityProfileId', 0)

        dlg = QDialog(self)
        dlg.setWindowTitle('Change Quality Profile')
        lay = QVBoxLayout(dlg)
        lay.addWidget(QLabel(f'Quality profile for: {item.text()}'))
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
            return

        new_qp_id = combo.currentData()
        if new_qp_id == current_qp_id:
            return

        try:
            series_data['qualityProfileId'] = new_qp_id
            self.api.update_series(series_data)
            # update the tree cell
            row_idx = item.index().row()
            parent = item.parent() or self.model.invisibleRootItem()
            qp_item = parent.child(row_idx, 3)
            if qp_item:
                qp_item.setText(combo.currentText())
            self.status_label.setText(f'Quality profile changed to {combo.currentText()} for {item.text()}')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to update quality profile:\n{e}')

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

        try:
            if node_type == 'series':
                self.api.delete_series(series_id, delete_files=True)
                parent = item.parent() or self.model.invisibleRootItem()
                parent.removeRow(item.row())
                self.status_label.setText(f'Deleted & unmonitored: {label}')

            elif node_type == 'season':
                season_num = item.data(ROLE_SEASON_NUM)
                season_path = item.data(ROLE_SEASON_PATH)

                # delete episode files from Sonarr
                for row in range(item.rowCount()):
                    ep_item = item.child(row, 0)
                    file_id = ep_item.data(ROLE_FILE_ID)
                    if file_id:
                        try:
                            self.api.delete_episode_file(file_id)
                        except Exception:
                            pass

                # unmonitor episodes
                self._unmonitor_season_episodes(item)

                # unmonitor the season
                try:
                    series_data = self.api.get_series_by_id(series_id)
                    for s in series_data.get('seasons', []):
                        if s.get('seasonNumber') == season_num:
                            s['monitored'] = False
                            break
                    self.api.update_series(series_data)
                except Exception:
                    pass

                # delete season directory
                if season_path and os.path.isdir(season_path):
                    shutil.rmtree(season_path, ignore_errors=True)

                parent = item.parent() or self.model.invisibleRootItem()
                parent.removeRow(item.row())
                self.status_label.setText(f'Deleted & unmonitored: {label}')

            elif node_type == 'episode':
                file_id = item.data(ROLE_FILE_ID)
                file_path = item.data(ROLE_FILE_PATH)
                # unmonitor
                ep_data_list = item.data(ROLE_EPISODE_DATA) or []
                for ep in ep_data_list:
                    ep['monitored'] = False
                    try:
                        self.api.update_episode(ep)
                    except Exception:
                        pass
                # delete file from Sonarr
                if file_id:
                    try:
                        self.api.delete_episode_file(file_id)
                    except Exception:
                        pass
                # delete from disk
                if file_path and os.path.isfile(file_path):
                    try:
                        os.remove(file_path)
                    except Exception:
                        pass
                parent = item.parent() or self.model.invisibleRootItem()
                parent.removeRow(item.row())
                self.status_label.setText(f'Deleted & unmonitored: {label}')

        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed:\n{e}')

    # ── keyboard / double-click actions ────────────────────────

    def _current_item(self) -> Optional[QStandardItem]:
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
        try:
            self.api.delete_series(series_id, delete_files=True)
            parent = item.parent() or self.model.invisibleRootItem()
            parent.removeRow(item.row())
            self.status_label.setText(f'Deleted series: {title}')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to delete series:\n{e}')

    def _delete_season(self, item: QStandardItem):
        # reuse the unmonitor & delete logic
        self._ctx_unmonitor_delete(item, 'season')

    # ── Add Show ───────────────────────────────────────────────

    def _add_show(self):
        dlg = AddShowDialog(self, self.api)
        if dlg.exec() == QDialog.Accepted and dlg.added_series:
            self._refresh()

    def closeEvent(self, event):
        """Save column widths and stop worker before closing."""
        widths = [self.tree.columnWidth(c) for c in range(len(self._columns))]
        self._settings.setValue('column_widths', widths)
        if self.worker.isRunning():
            self.worker.requestInterruption()
            self.worker.wait(5000)
        super().closeEvent(event)

    def _clear_cache_and_refresh(self):
        """Clear the ffprobe cache and reload all data."""
        global _probe_cache
        _probe_cache = {}
        _save_probe_cache()
        self.status_label.setText('Cache cleared, refreshing…')
        self._refresh()

    def _refresh(self):
        """Reload all data from Sonarr."""
        if self.worker.isRunning():
            try:
                self.worker.progress.disconnect()
                self.worker.series_ready.disconnect()
            except (RuntimeError, TypeError):
                pass
            self.worker.requestInterruption()
            self.worker.wait(5000)
        self.model.removeRows(0, self.model.rowCount())
        self.progress_bar.show()
        self.progress_bar.setRange(0, 0)
        self.status_label.setText('Refreshing…')
        self.worker = LoadWorker(self.api)
        self.worker.progress.connect(self._on_progress)
        self.worker.series_ready.connect(self._on_data_loaded)
        self.worker.start()


# ── entry point ────────────────────────────────────────────────────

def main():
    _load_probe_cache()

    # check ffprobe availability
    global _FFPROBE
    _FFPROBE = find_ffprobe()
    if not _FFPROBE:
        app = QApplication.instance() or QApplication(sys.argv)
        QMessageBox.critical(
            None, 'ffprobe not found',
            'ffprobe is required but was not found in PATH or\n'
            'common installation locations.\n\n'
            'Please install ffmpeg/ffprobe and try again.\n'
            'https://ffmpeg.org/download.html',
        )
        sys.exit(1)

    # load config (same as media_quality_checker)
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.toml')
    if not os.path.exists(config_path):
        print(f'Error: config.toml not found at {config_path}')
        sys.exit(1)

    with open(config_path, 'rb') as f:
        config = tomli.load(f)

    sonarr = config.get('sonarr', {})
    if not sonarr.get('enabled', True):
        print('Sonarr is disabled in config.toml')
        sys.exit(1)

    api = SonarrAPI(
        sonarr['url'], sonarr['api_key'],
        http_user=sonarr.get('http_basic_auth_username', ''),
        http_pass=sonarr.get('http_basic_auth_password', ''),
    )

    app = QApplication.instance() or QApplication(sys.argv)
    app.setStyle('Fusion')
    win = MainWindow(api, settings=config.get('settings', {}))
    win.showMaximized()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
