"""First-run setup wizard for arr-helper runtime configuration."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

import requests
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

if TYPE_CHECKING:
    from ...media_checker.config import Config


class _ServiceGroup(QGroupBox):
    def __init__(
        self,
        title: str,
        enabled: bool,
        url: str,
        api_key: str,
        http_user: str,
        http_password: str,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(title, parent)
        self.chk_enabled = QCheckBox("Enabled")
        self.chk_enabled.setChecked(bool(enabled))
        self.txt_url = QLineEdit(url)
        self.txt_api_key = QLineEdit(api_key)
        self.txt_http_user = QLineEdit(http_user)
        self.txt_http_password = QLineEdit(http_password)
        self.txt_http_password.setEchoMode(QLineEdit.EchoMode.Password)

        layout = QFormLayout(self)
        layout.addRow("", self.chk_enabled)
        layout.addRow("URL", self.txt_url)
        layout.addRow("API Key", self.txt_api_key)
        layout.addRow("HTTP User", self.txt_http_user)
        layout.addRow("HTTP Password", self.txt_http_password)

    def to_dict(self) -> dict[str, Any]:
        return {
            "enabled": bool(self.chk_enabled.isChecked()),
            "url": str(self.txt_url.text() or "").strip(),
            "api_key": str(self.txt_api_key.text() or "").strip(),
            "http_basic_auth_username": str(self.txt_http_user.text() or "").strip(),
            "http_basic_auth_password": str(
                self.txt_http_password.text() or ""
            ).strip(),
        }


class ArrSetupWizardDialog(QDialog):
    def __init__(self, initial: dict[str, Any], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("arr-helper Setup Wizard")
        self.resize(760, 620)

        root = QVBoxLayout(self)
        root.addWidget(
            QLabel(
                "Set Sonarr/Radarr connection settings. Values are stored in the shared settings INI."
            )
        )

        sonarr = initial.get("sonarr", {}) if isinstance(initial, dict) else {}
        radarr = initial.get("radarr", {}) if isinstance(initial, dict) else {}
        settings = initial.get("settings", {}) if isinstance(initial, dict) else {}

        self.sonarr_group = _ServiceGroup(
            "Sonarr",
            bool(sonarr.get("enabled", True)),
            str(sonarr.get("url", "") or ""),
            str(sonarr.get("api_key", "") or ""),
            str(sonarr.get("http_basic_auth_username", "") or ""),
            str(sonarr.get("http_basic_auth_password", "") or ""),
        )
        self.radarr_group = _ServiceGroup(
            "Radarr",
            bool(radarr.get("enabled", True)),
            str(radarr.get("url", "") or ""),
            str(radarr.get("api_key", "") or ""),
            str(radarr.get("http_basic_auth_username", "") or ""),
            str(radarr.get("http_basic_auth_password", "") or ""),
        )

        root.addWidget(self.sonarr_group)
        root.addWidget(self.radarr_group)

        test_row = QHBoxLayout()
        btn_test_sonarr = QPushButton("Test Sonarr Connection")
        btn_test_sonarr.clicked.connect(self._test_sonarr_connection)
        btn_test_radarr = QPushButton("Test Radarr Connection")
        btn_test_radarr.clicked.connect(self._test_radarr_connection)
        test_row.addWidget(btn_test_sonarr)
        test_row.addWidget(btn_test_radarr)
        test_row.addStretch(1)
        root.addLayout(test_row)

        options = QGroupBox("Checker Settings")
        options_layout = QFormLayout(options)
        self.chk_dry_run = QCheckBox("Dry run")
        self.chk_dry_run.setChecked(bool(settings.get("dry_run", False)))
        self.chk_interactive = QCheckBox("Interactive mode")
        self.chk_interactive.setChecked(bool(settings.get("interactive", True)))
        self.chk_req_audio = QCheckBox("Require English audio")
        self.chk_req_audio.setChecked(bool(settings.get("require_english_audio", True)))
        self.chk_req_subs = QCheckBox("Require English subtitles")
        self.chk_req_subs.setChecked(bool(settings.get("require_english_subs", True)))
        self.txt_lang_codes = QLineEdit(
            ", ".join(settings.get("english_language_codes", ["eng", "en", "english"]))
        )
        self.txt_highlight = QLineEdit(
            str(settings.get("highlight_missing_subs", "") or "")
        )

        options_layout.addRow("", self.chk_dry_run)
        options_layout.addRow("", self.chk_interactive)
        options_layout.addRow("", self.chk_req_audio)
        options_layout.addRow("", self.chk_req_subs)
        options_layout.addRow("English language codes", self.txt_lang_codes)
        options_layout.addRow("Highlight missing subs", self.txt_highlight)

        root.addWidget(options)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        row = QHBoxLayout()
        row.addStretch(1)
        row.addWidget(buttons)
        root.addLayout(row)

    def _on_accept(self) -> None:
        sonarr = self.sonarr_group.to_dict()
        radarr = self.radarr_group.to_dict()

        enabled = []
        if sonarr["enabled"]:
            enabled.append(("sonarr", sonarr))
        if radarr["enabled"]:
            enabled.append(("radarr", radarr))
        if not enabled:
            QMessageBox.warning(
                self, "Validation", "Enable at least one of Sonarr or Radarr."
            )
            return
        for name, service in enabled:
            if not str(service.get("url", "")).strip():
                QMessageBox.warning(
                    self, "Validation", f"[{name}] URL is required when enabled."
                )
                return
            if not str(service.get("api_key", "")).strip():
                QMessageBox.warning(
                    self, "Validation", f"[{name}] API key is required when enabled."
                )
                return
        self.accept()

    def to_config(self) -> dict[str, Any]:
        lang_codes = [
            token.strip()
            for token in str(self.txt_lang_codes.text() or "").split(",")
            if token.strip()
        ]
        if not lang_codes:
            lang_codes = ["eng", "en", "english"]
        return {
            "sonarr": self.sonarr_group.to_dict(),
            "radarr": self.radarr_group.to_dict(),
            "settings": {
                "dry_run": bool(self.chk_dry_run.isChecked()),
                "interactive": bool(self.chk_interactive.isChecked()),
                "require_english_audio": bool(self.chk_req_audio.isChecked()),
                "require_english_subs": bool(self.chk_req_subs.isChecked()),
                "english_language_codes": lang_codes,
                "highlight_missing_subs": str(self.txt_highlight.text() or "").strip(),
            },
        }

    def _test_sonarr_connection(self) -> None:
        self._test_service_connection("Sonarr", self.sonarr_group.to_dict())

    def _test_radarr_connection(self) -> None:
        self._test_service_connection("Radarr", self.radarr_group.to_dict())

    def _test_service_connection(
        self, service_name: str, service: dict[str, Any]
    ) -> None:
        if not bool(service.get("enabled", True)):
            QMessageBox.information(
                self,
                "Connection Test",
                f"{service_name} is disabled. Enable it before testing.",
            )
            return

        url = str(service.get("url", "") or "").strip()
        api_key = str(service.get("api_key", "") or "").strip()
        http_user = str(service.get("http_basic_auth_username", "") or "").strip()
        http_password = str(service.get("http_basic_auth_password", "") or "").strip()
        if not url:
            QMessageBox.warning(
                self, "Connection Test", f"{service_name} URL is required."
            )
            return
        if not api_key:
            QMessageBox.warning(
                self, "Connection Test", f"{service_name} API key is required."
            )
            return

        endpoint = f"{url.rstrip('/')}/api/v3/system/status"
        headers = {"X-Api-Key": api_key}
        auth = (http_user, http_password) if http_user else None
        try:
            response = requests.get(endpoint, headers=headers, auth=auth, timeout=10)
            response.raise_for_status()
            payload = response.json() if response.text else {}
            app_name = str(payload.get("appName", service_name))
            version = str(payload.get("version", "unknown"))
            QMessageBox.information(
                self,
                "Connection Test",
                f"{service_name} connection successful.\nDetected: {app_name} {version}",
            )
        except requests.RequestException as exc:
            QMessageBox.critical(
                self,
                "Connection Test Failed",
                f"{service_name} connection failed:\n{exc}",
            )


def run_setup_wizard(config_loader: Config, parent: QWidget | None = None) -> bool:
    dialog = ArrSetupWizardDialog(config_loader.config, parent=parent)
    if dialog.exec() != QDialog.DialogCode.Accepted:
        return False
    config_loader.save(dialog.to_config())
    return True
