"""Add-show dialog for Sonarr UI."""

import contextlib

import requests
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)

from ..api import SonarrAPI


class AddShowDialog(QDialog):
    """Search for a show on Sonarr and add it."""

    def __init__(self, parent, api: SonarrAPI):
        super().__init__(parent)
        self.api = api
        self.setWindowTitle("Add Show")
        pw = parent.width() if parent else 1000
        self.resize(int(pw * 0.7), 500)
        self.added_series = None

        layout = QVBoxLayout(self)

        # search bar
        search_row = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search for a show…")
        self.search_input.returnPressed.connect(self._do_search)
        btn_search = QPushButton("Search")
        btn_search.clicked.connect(self._do_search)
        search_row.addWidget(self.search_input)
        search_row.addWidget(btn_search)
        layout.addLayout(search_row)

        # root folder and quality profile selectors
        options_row = QHBoxLayout()
        options_row.addWidget(QLabel("Root Folder:"))
        self.combo_root = QComboBox()
        options_row.addWidget(self.combo_root, 1)
        options_row.addWidget(QLabel("Quality Profile:"))
        self.combo_qp = QComboBox()
        options_row.addWidget(self.combo_qp, 1)
        layout.addLayout(options_row)

        # preload root folders and quality profiles
        self._root_folders = []
        self._quality_profiles = []
        try:
            self._root_folders = api.get_root_folders()
            for rf in self._root_folders:
                self.combo_root.addItem(rf["path"], rf["path"])
        except Exception:
            pass
        try:
            self._quality_profiles = api.get_quality_profiles()
            for qp in self._quality_profiles:
                self.combo_qp.addItem(qp["name"], qp["id"])
        except Exception:
            pass

        # results list
        self.results_list = QListWidget()
        self.results_list.setAlternatingRowColors(True)
        layout.addWidget(self.results_list)

        # add button
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Add Selected")
        buttons.accepted.connect(self._add_selected)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.lookup_results: list[dict] = []

    def _do_search(self):
        term = self.search_input.text().strip()
        if not term:
            return
        self.results_list.clear()
        self.lookup_results = []
        try:
            results = self.api.lookup_series(term)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Search failed:\n{e}")
            return

        self.lookup_results = results
        for idx, r in enumerate(results):
            title = r.get("title", "?")
            year = r.get("year", "")
            overview = (r.get("overview", "") or "")[:120]
            status = r.get("status", "")
            already = "id" in r and r.get("id", 0) > 0
            label = f"{title} ({year}) [{status}]"
            if already:
                label += " [ALREADY IN SONARR]"
            if overview:
                label += f"\n  {overview}"
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, idx)
            self.results_list.addItem(item)

    def _add_selected(self):
        row = self.results_list.currentRow()
        if row < 0 or row >= len(self.lookup_results):
            QMessageBox.warning(self, "No Selection", "Please select a show first.")
            return

        lookup = self.lookup_results[row]

        # check if already added
        if lookup.get("id", 0) > 0:
            QMessageBox.information(
                self, "Already Added", "This series is already in Sonarr."
            )
            return

        root_path = self.combo_root.currentData()
        qp_id = self.combo_qp.currentData()
        if not root_path:
            QMessageBox.critical(self, "Error", "No root folder selected.")
            return
        if qp_id is None:
            QMessageBox.critical(self, "Error", "No quality profile selected.")
            return

        # build the payload from lookup data with required fields
        series_data = dict(lookup)  # copy so we don't mutate
        series_data["rootFolderPath"] = root_path
        series_data["qualityProfileId"] = qp_id
        series_data["monitored"] = False
        series_data["seasonFolder"] = True
        series_data["addOptions"] = {"searchForMissingEpisodes": False}
        # remove id=0 so Sonarr treats it as a new series
        series_data.pop("id", None)

        try:
            result = self.api.add_series(series_data)
            self.added_series = result
            QMessageBox.information(self, "Added", f"Added: {result.get('title', '?')}")
            self.accept()
        except requests.exceptions.HTTPError as e:
            body = ""
            if e.response is not None:
                with contextlib.suppress(Exception):
                    body = e.response.text
            QMessageBox.critical(self, "Error", f"Failed to add series:\n{e}\n\n{body}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to add series:\n{e}")
