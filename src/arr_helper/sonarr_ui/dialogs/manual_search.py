"""Manual search dialog for Sonarr UI."""

from PySide6.QtCore import QModelIndex, QSortFilterProxyModel, QStringListModel, Qt
from PySide6.QtGui import QStandardItem, QStandardItemModel
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QCompleter,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QTreeView,
    QVBoxLayout,
)
from threep_commons.settings import QSettingsValueStore

from ..roles import ROLE_RELEASE


class _SearchFilterProxy(QSortFilterProxyModel):
    """Filters on title text, quality, and indexer; sorts numerically via UserRole."""

    def __init__(self, col_quality=2, col_indexer=3, parent=None):
        super().__init__(parent)
        self._quality = ""
        self._indexer = ""
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
            if (model.data(idx, Qt.DisplayRole) or "") != self._quality:
                return False
        if self._indexer:
            idx = model.index(source_row, self._col_indexer, source_parent)
            if (model.data(idx, Qt.DisplayRole) or "") != self._indexer:
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

    @staticmethod
    def _ui_key(name: str) -> str:
        return f"ui/sonarr_ui/manual_search/{name}"

    def __init__(
        self,
        parent,
        title: str,
        releases: list[dict],
        settings: QSettingsValueStore | None = None,
    ):
        super().__init__(parent)
        self.setWindowTitle(f"Manual Search: {title}")
        self.setWindowState(Qt.WindowMaximized)
        self.selected_release = None
        self._settings = settings

        layout = QVBoxLayout(self)

        # filter row: text filter + quality combo
        filter_row = QHBoxLayout()
        self.filter_input = QLineEdit()
        self.filter_input.setPlaceholderText("Filter results…")
        self.filter_input.textChanged.connect(self._apply_filters)
        # autocomplete from saved history
        self._filter_history = []
        if settings:
            self._filter_history = (
                settings.value(self._ui_key("search_filter_history"), []) or []
            )
        self._completer_model = QStringListModel(self._filter_history)
        completer = QCompleter(self._completer_model, self)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        self.filter_input.setCompleter(completer)
        filter_row.addWidget(self.filter_input, 1)

        filter_row.addWidget(QLabel("Quality:"))
        self.quality_combo = QComboBox()
        self.quality_combo.addItem("All", "")
        # collect unique quality names
        qualities = sorted(
            {
                rel.get("quality", {}).get("quality", {}).get("name", "")
                for rel in releases
            }
            - {""}
        )
        for q in qualities:
            self.quality_combo.addItem(q, q)
        self.quality_combo.currentIndexChanged.connect(self._apply_filters)
        filter_row.addWidget(self.quality_combo)

        filter_row.addWidget(QLabel("Indexer:"))
        self.indexer_combo = QComboBox()
        self.indexer_combo.addItem("All", "")
        indexers = sorted({rel.get("indexer", "") for rel in releases} - {""})
        for ix in indexers:
            self.indexer_combo.addItem(ix, ix)
        self.indexer_combo.currentIndexChanged.connect(self._apply_filters)
        filter_row.addWidget(self.indexer_combo)
        layout.addLayout(filter_row)

        # source model
        self._columns = ["Title", "Size (GB)", "Quality", "Indexer", "Age"]
        self.source_model = QStandardItemModel()
        self.source_model.setHorizontalHeaderLabels(self._columns)

        # populate
        for rel in releases:
            title_item = QStandardItem(rel.get("title", ""))
            title_item.setEditable(False)
            title_item.setData(rel, ROLE_RELEASE)

            size = rel.get("size", 0)
            size_gb = size / (1024**3) if size > 0 else 0.0
            size_item = QStandardItem()
            size_item.setEditable(False)
            size_item.setData(round(size_gb, 2), Qt.DisplayRole)

            quality = rel.get("quality", {}).get("quality", {}).get("name", "")
            quality_item = QStandardItem(quality)
            quality_item.setEditable(False)

            indexer_item = QStandardItem(rel.get("indexer", ""))
            indexer_item.setEditable(False)

            age_hours = rel.get("ageHours", 0) or 0
            if age_hours:
                age_days = age_hours / 24
            else:
                age_days = rel.get("age", 0) or 0
                age_hours = age_days * 24
            if age_hours < 1:
                age_str = "< 1h"
            elif age_hours < 24:
                age_str = f"{int(age_hours)}h"
            else:
                age_str = f"{int(age_days)}d"
            age_item = QStandardItem(age_str)
            age_item.setEditable(False)
            age_item.setData(age_hours, Qt.UserRole)

            self.source_model.appendRow(
                [title_item, size_item, quality_item, indexer_item, age_item]
            )

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
        self.table.sortByColumn(
            4, Qt.AscendingOrder
        )  # default sort by age asc (newest first)
        header = self.table.header()
        header.setSectionResizeMode(QHeaderView.Interactive)
        saved = (
            settings.value(self._ui_key("search_column_widths")) if settings else None
        )
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
        self.proxy.set_quality(self.quality_combo.currentData() or "")
        self.proxy.set_indexer(self.indexer_combo.currentData() or "")

    def _save_filter_history(self):
        text = self.filter_input.text().strip()
        if text and self._settings and text not in self._filter_history:
            self._filter_history.append(text)
            # keep last 50
            self._filter_history = self._filter_history[-50:]
            self._settings.set_value(
                self._ui_key("search_filter_history"), self._filter_history
            )
            self._completer_model.setStringList(self._filter_history)

    def _save_column_widths(self):
        if self._settings:
            widths = [self.table.columnWidth(c) for c in range(len(self._columns))]
            self._settings.set_value(self._ui_key("search_column_widths"), widths)

    def _get_selected_release(self) -> dict | None:
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
