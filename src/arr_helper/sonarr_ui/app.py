"""Application entrypoint for Sonarr UI."""

from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication, QMessageBox
from threep_commons.paths import configure_qsettings

from ..constants import APP_IDENTITY
from ..core.ffprobe import find_ffprobe
from ..media_checker.config import Config
from .api import SonarrAPI
from .dialogs.setup_wizard import run_setup_wizard
from .main_window import MainWindow
from .probe_cache import load_probe_cache, set_ffprobe_path


def _get_or_create_qapp() -> QApplication:
    existing = QApplication.instance()
    if isinstance(existing, QApplication):
        return existing
    return QApplication(sys.argv)


def _show_validation_failure_and_exit(errors: list[str], config_loader: Config) -> None:
    details = "\n".join(f"  - {item}" for item in errors)
    print("Configuration validation failed:")
    if details:
        print(details)
    print(f"\nOpen and edit settings INI: {config_loader.config_path}")
    sys.exit(1)


def main() -> int:
    """Launch the Sonarr UI after validating config and ffprobe availability."""
    configure_qsettings(APP_IDENTITY)
    load_probe_cache()

    ffprobe_path = find_ffprobe()
    set_ffprobe_path(ffprobe_path)
    if not ffprobe_path:
        app = _get_or_create_qapp()
        QMessageBox.critical(
            None,
            "ffprobe not found",
            "ffprobe is required but was not found in PATH or\n"
            "common installation locations.\n\n"
            "Please install ffmpeg/ffprobe and try again.\n"
            "https://ffmpeg.org/download.html",
        )
        return 1

    app = _get_or_create_qapp()
    config_loader = Config()

    missing = config_loader.validate(context="sonarr_ui")
    if missing:
        if not run_setup_wizard(config_loader):
            print("Configuration wizard was cancelled. Exiting.")
            return 1
        config_loader.reload()
        missing = config_loader.validate(context="sonarr_ui")
        if missing:
            _show_validation_failure_and_exit(missing, config_loader)

    sonarr = config_loader.get_sonarr_config()
    if sonarr is None:
        print(f"Sonarr is disabled in {config_loader.config_path}")
        return 1

    api = SonarrAPI(
        str(sonarr.get("url", "")),
        str(sonarr.get("api_key", "")),
        http_user=str(sonarr.get("http_basic_auth_username", "")),
        http_pass=str(sonarr.get("http_basic_auth_password", "")),
        request_timeout=300,
    )
    loader_api = SonarrAPI(
        str(sonarr.get("url", "")),
        str(sonarr.get("api_key", "")),
        http_user=str(sonarr.get("http_basic_auth_username", "")),
        http_pass=str(sonarr.get("http_basic_auth_password", "")),
        request_timeout=300,
    )

    app.setStyle("Fusion")
    win = MainWindow(api, loader_api=loader_api, settings=config_loader.get_settings())
    win.showMaximized()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
