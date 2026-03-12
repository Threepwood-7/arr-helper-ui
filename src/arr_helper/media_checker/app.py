"""CLI entrypoint for media quality checker."""

from __future__ import annotations

import sys

from rich import box
from rich.panel import Panel
from threep_commons.paths import configure_qsettings

from ..constants import APP_IDENTITY
from ..core.ffprobe import find_ffprobe
from .checker import MediaQualityChecker
from .config import Config

HttpAuth = tuple[str, str]


def _print_config_errors_and_exit(config: Config, errors: list[str]) -> None:
    print("Configuration validation failed:")
    for err in errors:
        print(f"  - {err}")
    print(f"\nOpen and edit settings INI: {config.config_path}")
    print("Then run Sonarr UI setup wizard once: python -m arr_helper")
    sys.exit(1)


def main() -> int:
    configure_qsettings(APP_IDENTITY)
    config = Config()
    validation_errors = config.validate(context="media_checker")
    if validation_errors:
        _print_config_errors_and_exit(config, validation_errors)

    ffprobe_path = find_ffprobe()
    if not ffprobe_path:
        print("Error: ffprobe not found in PATH or common install locations.")
        print("Please install ffmpeg/ffprobe: https://ffmpeg.org/download.html")
        return 1

    settings = config.get_settings()
    dry_run = bool(settings.get("dry_run", False))
    interactive = bool(settings.get("interactive", False))
    require_audio = bool(settings.get("require_english_audio", True))
    require_subs = bool(settings.get("require_english_subs", True))
    english_codes_raw = settings.get("english_language_codes", ["eng", "en", "english"])

    sonarr_config = config.get_sonarr_config()
    radarr_config = config.get_radarr_config()

    if not sonarr_config and not radarr_config:
        print("Error: Both Sonarr and Radarr are disabled in config.")
        return 1

    def _http_auth(cfg: dict[str, object] | None) -> HttpAuth | None:
        user = str(cfg.get("http_basic_auth_username", "") if cfg else "")
        password = str(cfg.get("http_basic_auth_password", "") if cfg else "")
        return (user, password) if user else None

    english_codes_list = (
        [str(code) for code in english_codes_raw]
        if isinstance(english_codes_raw, list)
        else ["eng", "en", "english"]
    )

    checker = MediaQualityChecker(
        sonarr_url=str(sonarr_config.get("url", "") if sonarr_config else ""),
        sonarr_api=str(sonarr_config.get("api_key", "") if sonarr_config else ""),
        radarr_url=str(radarr_config.get("url", "") if radarr_config else ""),
        radarr_api=str(radarr_config.get("api_key", "") if radarr_config else ""),
        require_audio=require_audio,
        require_subs=require_subs,
        english_codes=english_codes_list,
        interactive=interactive,
        config=config,
        sonarr_http_auth=_http_auth(sonarr_config),
        radarr_http_auth=_http_auth(radarr_config),
        ffprobe_path=ffprobe_path,
    )

    if interactive:
        if dry_run:
            checker.console.print(
                Panel(
                    "[bold yellow]DRY RUN MODE[/bold yellow] - No changes will be made",
                    box=box.DOUBLE,
                )
            )
        else:
            checker.console.print("")
    else:
        if dry_run:
            print("=== DRY RUN MODE - No changes will be made ===\n")

    if sonarr_config:
        checker.process_sonarr(dry_run)

    if radarr_config:
        checker.process_radarr(dry_run)

    if interactive:
        checker.console.print(
            Panel("[bold green]Complete[/bold green]", box=box.DOUBLE)
        )
    else:
        print("\n=== Complete ===")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
