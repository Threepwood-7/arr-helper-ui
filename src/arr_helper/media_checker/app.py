"""CLI entrypoint for media quality checker."""

import sys

from rich import box
from rich.panel import Panel

from ..core.ffprobe import find_ffprobe
from .checker import MediaQualityChecker
from .config import Config


def main():
    config = Config()
    config.validate()

    ffprobe_path = find_ffprobe()
    if not ffprobe_path:
        print('Error: ffprobe not found in PATH or common install locations.')
        print('Please install ffmpeg/ffprobe: https://ffmpeg.org/download.html')
        sys.exit(1)

    settings = config.get_settings()
    dry_run = settings.get('dry_run', False)
    interactive = settings.get('interactive', False)
    require_audio = settings.get('require_english_audio', True)
    require_subs = settings.get('require_english_subs', True)
    english_codes = settings.get('english_language_codes', ['eng', 'en', 'english'])

    sonarr_config = config.get_sonarr_config()
    radarr_config = config.get_radarr_config()

    if not sonarr_config and not radarr_config:
        print('Error: Both Sonarr and Radarr are disabled in config.')
        print(f'Please enable at least one in {config.config_path}')
        sys.exit(1)

    def _http_auth(cfg):
        user = cfg.get('http_basic_auth_username', '') if cfg else ''
        password = cfg.get('http_basic_auth_password', '') if cfg else ''
        return (user, password) if user else None

    checker = MediaQualityChecker(
        sonarr_url=sonarr_config.get('url', '') if sonarr_config else '',
        sonarr_api=sonarr_config.get('api_key', '') if sonarr_config else '',
        radarr_url=radarr_config.get('url', '') if radarr_config else '',
        radarr_api=radarr_config.get('api_key', '') if radarr_config else '',
        require_audio=require_audio,
        require_subs=require_subs,
        english_codes=english_codes,
        interactive=interactive,
        config=config,
        sonarr_http_auth=_http_auth(sonarr_config),
        radarr_http_auth=_http_auth(radarr_config),
        ffprobe_path=ffprobe_path,
    )

    if interactive:
        if dry_run:
            checker.console.print(Panel('[bold yellow]DRY RUN MODE[/bold yellow] - No changes will be made', box=box.DOUBLE))
        else:
            checker.console.print('')
    else:
        if dry_run:
            print('=== DRY RUN MODE - No changes will be made ===\n')

    if sonarr_config:
        checker.process_sonarr(dry_run)

    if radarr_config:
        checker.process_radarr(dry_run)

    if interactive:
        checker.console.print(Panel('[bold green]Complete[/bold green]', box=box.DOUBLE))
    else:
        print('\n=== Complete ===')


if __name__ == '__main__':
    main()
