"""Application entrypoint for Sonarr UI."""

import sys

from PySide6.QtWidgets import QApplication, QMessageBox

from ..core.ffprobe import find_ffprobe
from ..media_checker.config import Config
from .api import SonarrAPI
from .main_window import MainWindow
from .probe_cache import _load_probe_cache, set_ffprobe_path


def main():
    _load_probe_cache()

    ffprobe_path = find_ffprobe()
    set_ffprobe_path(ffprobe_path)
    if not ffprobe_path:
        app = QApplication.instance() or QApplication(sys.argv)
        QMessageBox.critical(
            None,
            'ffprobe not found',
            'ffprobe is required but was not found in PATH or\n'
            'common installation locations.\n\n'
            'Please install ffmpeg/ffprobe and try again.\n'
            'https://ffmpeg.org/download.html',
        )
        sys.exit(1)

    config_loader = Config()
    config_loader.validate()
    config = config_loader.config
    sonarr = config_loader.get_sonarr_config()
    if sonarr is None:
        print(f"Sonarr is disabled in {config_loader.config_path}")
        sys.exit(1)

    sonarr = dict(sonarr)
    if not sonarr.get('enabled', True):
        print(f"Sonarr is disabled in {config_loader.config_path}")
        sys.exit(1)

    errors = []
    placeholders = {'your_sonarr_api_key_here', 'your-sonarr-api-key-here', ''}
    if 'url' not in sonarr or not sonarr['url'].strip():
        errors.append("[sonarr] 'url' is missing or empty")
    if 'api_key' not in sonarr:
        errors.append("[sonarr] 'api_key' is missing")
    elif sonarr['api_key'].strip().lower() in placeholders:
        errors.append("[sonarr] 'api_key' is still set to a placeholder value")
    if errors:
        print('Config validation failed:')
        for err in errors:
            print(f'  - {err}')
        print(f"\nPlease edit {config_loader.config_path}")
        sys.exit(1)

    api = SonarrAPI(
        sonarr['url'],
        sonarr['api_key'],
        http_user=sonarr.get('http_basic_auth_username', ''),
        http_pass=sonarr.get('http_basic_auth_password', ''),
        request_timeout=300,
    )
    loader_api = SonarrAPI(
        sonarr['url'],
        sonarr['api_key'],
        http_user=sonarr.get('http_basic_auth_username', ''),
        http_pass=sonarr.get('http_basic_auth_password', ''),
        request_timeout=300,
    )

    app = QApplication.instance() or QApplication(sys.argv)
    app.setStyle('Fusion')
    win = MainWindow(api, loader_api=loader_api, settings=config.get('settings', {}))
    win.showMaximized()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
