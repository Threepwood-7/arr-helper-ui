"""Manual search dialog for Sonarr UI."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast, final

from PySide6.QtCore import (
    QModelIndex,
    QPersistentModelIndex,
    QSortFilterProxyModel,
    QStringListModel,
    Qt,
)
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
    QWidget,
)

from ..roles import ROLE_RELEASE

if TYPE_CHECKING:
    from threep_commons.settings import QSettingsValueStore

    from ..api import JsonDict

_DISPLAY_ROLE = int(Qt.ItemDataRole.DisplayRole)
_USER_ROLE = int(Qt.ItemDataRole.UserRole)


def _release_title(release: JsonDict) -> str:
    return str(release.get("title", ""))


def _release_quality(release: JsonDict) -> str:
    quality = cast("JsonDict", release.get("quality", {}))
    nested = cast("JsonDict", quality.get("quality", {}))
    return str(nested.get("name", ""))


def _release_indexer(release: JsonDict) -> str:
    return str(release.get("indexer", ""))


def _release_size_gb(release: JsonDict) -> float:
    size_obj = release.get("size", 0)
    size = int(size_obj) if isinstance(size_obj, int | float) else 0
    return size / (1024**3) if size > 0 else 0.0


def _release_age_hours(release: JsonDict) -> int:
    age_hours_obj = release.get("ageHours", 0)
    if isinstance(age_hours_obj, int | float) and age_hours_obj > 0:
        return int(age_hours_obj)
    age_days_obj = release.get("age", 0)
    if isinstance(age_days_obj, int | float) and age_days_obj > 0:
        return int(age_days_obj * 24)
    return 0


def _normalize_string_list(value: object) -> list[str]:
    if isinstance(value, list | tuple):
        values = cast("list[object] | tuple[object, ...]", value)
        return [str(item) for item in values]
    return []


def _normalize_int_list(value: object, *, size: int) -> list[int] | None:
    if not isinstance(value, list | tuple):
        return None
    values = cast("list[object] | tuple[object, ...]", value)
    if len(values) != size:
        return None
    widths: list[int] = []
    for item in values:
        if isinstance(item, int | float | str):
            try:
                widths.append(int(item))
            except (TypeError, ValueError, OverflowError):
                return None
            continue
        return None
    return widths


class _SearchFilterProxy(QSortFilterProxyModel):
    """Filters on title text, quality, and indexer; sorts numerically via UserRole."""

    def __init__(
        self,
        col_quality: int = 2,
        col_indexer: int = 3,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._quality = ""
        self._indexer = ""
        self._col_quality = col_quality
        self._col_indexer = col_indexer

    def set_quality(self, quality: str) -> None:
        self._quality = quality
        self.invalidateFilter()

    def set_indexer(self, indexer: str) -> None:
        self._indexer = indexer
        self.invalidateFilter()

    @staticmethod
    def _display_text(
        index: QModelIndex | QPersistentModelIndex,
        role: int = _DISPLAY_ROLE,
    ) -> str:
        value = index.data(role)
        return str(value or "")

    def filterAcceptsRow(
        self,
        source_row: int,
        source_parent: QModelIndex | QPersistentModelIndex,
    ) -> bool:
        if not super().filterAcceptsRow(source_row, source_parent):
            return False
        model = self.sourceModel()
        if self._quality:
            index = model.index(source_row, self._col_quality, source_parent)
            if self._display_text(index) != self._quality:
                return False
        if self._indexer:
            index = model.index(source_row, self._col_indexer, source_parent)
            if self._display_text(index) != self._indexer:
                return False
        return True

    def lessThan(
        self,
        left: QModelIndex | QPersistentModelIndex,
        right: QModelIndex | QPersistentModelIndex,
    ) -> bool:
        left_value = left.data(_USER_ROLE)
        right_value = right.data(_USER_ROLE)
        if isinstance(left_value, int | float) and isinstance(right_value, int | float):
            return float(left_value) < float(right_value)
        return super().lessThan(left, right)


@final
class ManualSearchDialog(QDialog):
    """Shows release search results with sorting and filtering."""

    _COL_QUALITY = 2
    _COL_INDEXER = 3

    @staticmethod
    def _ui_key(name: str) -> str:
        return f"ui/sonarr_ui/manual_search/{name}"

    def __init__(
        self,
        parent: QWidget | None,
        title: str,
        releases: list[JsonDict],
        settings: QSettingsValueStore | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"Manual Search: {title}")
        self.setWindowState(Qt.WindowState.WindowMaximized)
        self.selected_release: JsonDict | None = None
        self._settings = settings
        self._filter_history: list[str] = []

        layout = QVBoxLayout(self)

        filter_row = QHBoxLayout()
        self.filter_input = QLineEdit()
        self.filter_input.setPlaceholderText("Filter results…")
        _ = self.filter_input.textChanged.connect(self._apply_filters)
        if settings is not None:
            self._filter_history = _normalize_string_list(
                settings.value(self._ui_key("search_filter_history"), [])
            )
        self._completer_model = QStringListModel(self._filter_history)
        completer = QCompleter(self._completer_model, self)
        completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        completer.setFilterMode(Qt.MatchFlag.MatchContains)
        self.filter_input.setCompleter(completer)
        filter_row.addWidget(self.filter_input, 1)

        filter_row.addWidget(QLabel("Quality:"))
        self.quality_combo = QComboBox()
        self.quality_combo.addItem("All", "")
        for quality in sorted(
            {_release_quality(release) for release in releases} - {""}
        ):
            self.quality_combo.addItem(quality, quality)
        _ = self.quality_combo.currentIndexChanged.connect(self._apply_filters)
        filter_row.addWidget(self.quality_combo)

        filter_row.addWidget(QLabel("Indexer:"))
        self.indexer_combo = QComboBox()
        self.indexer_combo.addItem("All", "")
        for indexer in sorted(
            {_release_indexer(release) for release in releases} - {""}
        ):
            self.indexer_combo.addItem(indexer, indexer)
        _ = self.indexer_combo.currentIndexChanged.connect(self._apply_filters)
        filter_row.addWidget(self.indexer_combo)
        layout.addLayout(filter_row)

        self._columns = ["Title", "Size (GB)", "Quality", "Indexer", "Age"]
        self.source_model = QStandardItemModel()
        self.source_model.setHorizontalHeaderLabels(self._columns)

        for release in releases:
            title_item = QStandardItem(_release_title(release))
            title_item.setEditable(False)
            title_item.setData(release, ROLE_RELEASE)

            size_item = QStandardItem()
            size_item.setEditable(False)
            size_item.setData(round(_release_size_gb(release), 2), _DISPLAY_ROLE)

            quality_item = QStandardItem(_release_quality(release))
            quality_item.setEditable(False)

            indexer_item = QStandardItem(_release_indexer(release))
            indexer_item.setEditable(False)

            age_hours = _release_age_hours(release)
            if age_hours < 1:
                age_text = "< 1h"
            elif age_hours < 24:
                age_text = f"{age_hours}h"
            else:
                age_text = f"{age_hours // 24}d"
            age_item = QStandardItem(age_text)
            age_item.setEditable(False)
            age_item.setData(age_hours, _USER_ROLE)

            self.source_model.appendRow(
                [title_item, size_item, quality_item, indexer_item, age_item]
            )

        self.proxy = _SearchFilterProxy(self._COL_QUALITY, self._COL_INDEXER)
        self.proxy.setSourceModel(self.source_model)
        self.proxy.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.proxy.setFilterKeyColumn(0)

        self.table = QTreeView()
        self.table.setModel(self.proxy)
        self.table.setRootIsDecorated(False)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setSortingEnabled(True)
        self.table.sortByColumn(4, Qt.SortOrder.AscendingOrder)
        header = self.table.header()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        saved_widths = (
            _normalize_int_list(
                settings.value(self._ui_key("search_column_widths")),
                size=len(self._columns),
            )
            if settings is not None
            else None
        )
        if saved_widths is not None:
            for column, width in enumerate(saved_widths):
                self.table.setColumnWidth(column, width)
        else:
            self.table.setColumnWidth(0, 800)
            for column in range(1, len(self._columns)):
                self.table.resizeColumnToContents(column)
        _ = self.table.doubleClicked.connect(self._on_double_click)
        layout.addWidget(self.table)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        _ = buttons.accepted.connect(self._on_ok)
        _ = buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _apply_filters(self) -> None:
        self.proxy.setFilterFixedString(self.filter_input.text())
        self.proxy.set_quality(str(self.quality_combo.currentData() or ""))
        self.proxy.set_indexer(str(self.indexer_combo.currentData() or ""))

    def _save_filter_history(self) -> None:
        text = self.filter_input.text().strip()
        if text and self._settings is not None and text not in self._filter_history:
            self._filter_history.append(text)
            self._filter_history = self._filter_history[-50:]
            self._settings.set_value(
                self._ui_key("search_filter_history"), self._filter_history
            )
            self._completer_model.setStringList(self._filter_history)

    def _save_column_widths(self) -> None:
        if self._settings is not None:
            widths = [
                self.table.columnWidth(column) for column in range(len(self._columns))
            ]
            self._settings.set_value(self._ui_key("search_column_widths"), widths)

    def _get_selected_release(self) -> JsonDict | None:
        indexes = self.table.selectionModel().selectedRows()
        if not indexes:
            return None
        source_index = self.proxy.mapToSource(indexes[0])
        item = self.source_model.item(source_index.row(), 0)
        payload = item.data(ROLE_RELEASE)
        return cast("JsonDict", payload) if isinstance(payload, dict) else None

    def _on_double_click(self, _index: QModelIndex) -> None:
        release = self._get_selected_release()
        if release is not None:
            self.selected_release = release
            self._save_filter_history()
            self._save_column_widths()
            super().accept()

    def _on_ok(self) -> None:
        release = self._get_selected_release()
        if release is not None:
            self.selected_release = release
            self._save_filter_history()
            self._save_column_widths()
            super().accept()

    def reject(self) -> None:
        self._save_filter_history()
        self._save_column_widths()
        super().reject()
