"""Media quality checking logic for Sonarr/Radarr."""

from __future__ import annotations

import json
import os
import subprocess
from pathlib import Path
from typing import TYPE_CHECKING, cast

import requests
from rich import box
from rich.console import Console
from rich.panel import Panel
from rich.prompt import Confirm
from rich.table import Table

if TYPE_CHECKING:
    from .config import Config

ConfigMap = dict[str, object]
ConfigList = list[ConfigMap]
HttpAuth = tuple[str, str]
RequestPayload = ConfigMap | None
RequestResult = ConfigMap | ConfigList


class MediaQualityChecker:
    def __init__(
        self,
        sonarr_url: str,
        sonarr_api: str,
        radarr_url: str,
        radarr_api: str,
        require_audio: bool = True,
        require_subs: bool = True,
        english_codes: list[str] | None = None,
        interactive: bool = False,
        config: Config | None = None,
        sonarr_http_auth: HttpAuth | None = None,
        radarr_http_auth: HttpAuth | None = None,
        ffprobe_path: str = "ffprobe",
    ) -> None:
        self.sonarr_url = sonarr_url.rstrip("/")
        self.sonarr_api = sonarr_api
        self.radarr_url = radarr_url.rstrip("/")
        self.radarr_api = radarr_api
        self.sonarr_http_auth = sonarr_http_auth
        self.radarr_http_auth = radarr_http_auth
        self.require_audio = require_audio
        self.require_subs = require_subs
        self.english_codes = [
            code.lower() for code in (english_codes or ["eng", "en", "english"])
        ]
        self.interactive = interactive
        self.console = Console()
        self.config = config
        self.ffprobe_path = ffprobe_path

        # Load caches
        self.user_cache = config.load_user_cache() if config else {}
        self.files_cache = config.load_files_cache() if config else {}

        # Keep cache entries keyed by file path, validated by a file signature
        # (size + mtime) to avoid stale "good/skipped" decisions.
        self._good_files_map = self._normalize_cache_map(
            self.files_cache.get("good_files", {})
        )
        self._skipped_files_map = self._normalize_cache_map(
            self.user_cache.get("skipped_files", {})
        )
        self.files_cache["good_files"] = self._good_files_map
        self.user_cache["skipped_files"] = self._skipped_files_map

    @staticmethod
    def _file_signature(file_path: str) -> str:
        try:
            stat = os.stat(file_path)
            return f"{stat.st_size}:{stat.st_mtime_ns}"
        except OSError:
            return ""

    def _normalize_cache_map(self, raw_value: object) -> dict[str, str]:
        normalized: dict[str, str] = {}
        if isinstance(raw_value, dict):
            for path, signature in raw_value.items():
                if isinstance(path, str) and isinstance(signature, str) and signature:
                    normalized[path] = signature
        elif isinstance(raw_value, list):
            # Backward compatibility with old list[str] format.
            for path in raw_value:
                if isinstance(path, str):
                    signature = self._file_signature(path)
                    if signature:
                        normalized[path] = signature
        return normalized

    def _is_cached_match(self, cache_map: dict[str, str], file_path: str) -> bool:
        cached_signature = cache_map.get(file_path)
        if not cached_signature:
            return False
        current_signature = self._file_signature(file_path)
        if current_signature and current_signature == cached_signature:
            return True
        # Drop stale cache entry when file contents changed or disappeared.
        cache_map.pop(file_path, None)
        return False

    def _add_good_file(self, file_path: str) -> None:
        """Add a file to the good files cache."""
        signature = self._file_signature(file_path)
        if signature:
            self._good_files_map[file_path] = signature

    def _add_skipped_file(self, file_path: str) -> None:
        """Add a file to the skipped files cache."""
        signature = self._file_signature(file_path)
        if signature:
            self._skipped_files_map[file_path] = signature

    def save_caches(self) -> None:
        """Save user cache and files cache to disk"""
        if self.config:
            self.config.save_user_cache(self.user_cache)
            self.config.save_files_cache(self.files_cache)

    def _make_request(
        self,
        url: str,
        api_key: str,
        endpoint: str,
        method: str = "GET",
        data: RequestPayload = None,
        auth: HttpAuth | None = None,
    ) -> RequestResult | None:
        """Make API request to Sonarr/Radarr"""
        headers = {"X-Api-Key": api_key}
        full_url = f"{url}/api/v3/{endpoint}"
        if auth is None:
            if url == self.sonarr_url:
                auth = self.sonarr_http_auth
            elif url == self.radarr_url:
                auth = self.radarr_http_auth

        try:
            if method == "GET":
                response = requests.get(
                    full_url, headers=headers, auth=auth, timeout=300
                )
            elif method == "PUT":
                response = requests.put(
                    full_url, headers=headers, auth=auth, json=data, timeout=300
                )
            elif method == "DELETE":
                response = requests.delete(
                    full_url, headers=headers, auth=auth, timeout=300
                )
            else:
                response = requests.post(
                    full_url, headers=headers, auth=auth, json=data, timeout=300
                )

            response.raise_for_status()
            if not response.text:
                return {}
            try:
                payload = response.json()
                return cast("RequestResult", payload)
            except ValueError as e:
                print(f"Error parsing JSON from {full_url}: {e}")
                return None
        except requests.exceptions.RequestException as e:
            print(f"Error making request to {full_url}: {e}")
            return None

    def check_file_streams(self, file_path: str) -> tuple[bool, bool]:
        """
        Check if file has English audio and subtitles using ffprobe
        Returns: (has_eng_audio, has_eng_subs)
        """
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            return False, False

        try:
            # Run ffprobe to get stream information
            cmd = [
                self.ffprobe_path,
                "-v",
                "quiet",
                "-print_format",
                "json",
                "-show_streams",
                file_path,
            ]

            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=60,
                creationflags=creationflags,
            )

            if result.returncode != 0:
                print(f"ffprobe error for {file_path}: {result.stderr}")
                return False, False

            payload = cast("ConfigMap", json.loads(result.stdout))
            streams_obj = payload.get("streams", [])
            streams = (
                cast("list[ConfigMap]", streams_obj)
                if isinstance(streams_obj, list)
                else []
            )

            has_eng_audio = False
            has_eng_subs = False

            for stream in streams:
                codec_type = str(stream.get("codec_type", ""))
                tags_obj = stream.get("tags", {})
                tags = cast("ConfigMap", tags_obj) if isinstance(tags_obj, dict) else {}
                language = str(tags.get("language", "")).lower()

                # Check for English audio
                if codec_type == "audio" and language in self.english_codes:
                    has_eng_audio = True

                # Check for English subtitles
                if codec_type == "subtitle" and language in self.english_codes:
                    has_eng_subs = True

            return has_eng_audio, has_eng_subs

        except subprocess.TimeoutExpired:
            print(f"ffprobe timeout for {file_path}")
            return False, False
        except json.JSONDecodeError as e:
            print(f"JSON decode error for {file_path}: {e}")
            return False, False
        except Exception as e:
            print(f"Error checking {file_path}: {e}")
            return False, False

    def should_redownload(self, has_eng_audio: bool, has_eng_subs: bool) -> bool:
        """Determine if file should be re-downloaded based on config"""
        if self.require_audio and not has_eng_audio:
            return True
        return bool(self.require_subs and not has_eng_subs)

    def get_episode_releases(
        self, episode_id: int, quality_profile_id: int | None = None
    ) -> ConfigList:
        """Get available releases for an episode from Sonarr"""
        endpoint = f"release?episodeId={episode_id}"
        if quality_profile_id:
            endpoint += f"&qualityProfileId={quality_profile_id}"

        releases = self._make_request(
            self.sonarr_url,
            self.sonarr_api,
            endpoint,
            auth=self.sonarr_http_auth,
        )
        return releases if isinstance(releases, list) else []

    def get_episodes_for_file(self, series_id: int, episode_file_id: int) -> list[int]:
        """Get episode IDs associated with an episode file"""
        episodes = self._make_request(
            self.sonarr_url,
            self.sonarr_api,
            f"episode?seriesId={series_id}",
            auth=self.sonarr_http_auth,
        )

        if not isinstance(episodes, list):
            return []

        # Find episodes that use this file
        episode_ids: list[int] = []
        for ep in episodes:
            if ep.get("episodeFileId") == episode_file_id:
                episode_id = ep.get("id")
                if isinstance(episode_id, int):
                    episode_ids.append(episode_id)

        return episode_ids

    def get_movie_releases(
        self, movie_id: int, quality_profile_id: int | None = None
    ) -> ConfigList:
        """Get available releases for a movie from Radarr"""
        endpoint = f"release?movieId={movie_id}"
        if quality_profile_id:
            endpoint += f"&qualityProfileId={quality_profile_id}"

        releases = self._make_request(
            self.radarr_url,
            self.radarr_api,
            endpoint,
            auth=self.radarr_http_auth,
        )
        return releases if isinstance(releases, list) else []

    def display_releases_and_select(
        self, releases: list[dict], title: str, file_path: str
    ) -> dict | None:
        """Display releases in a table and let user select one"""
        if not releases:
            self.console.print("[yellow]No alternative releases found[/yellow]")
            return None

        # Sort releases by quality profile match (preferred first) then size descending
        def sort_key(release):
            # Get quality score - higher is better
            quality = release.get("quality", {}).get("quality", {})
            quality_score = quality.get("id", 0)

            # Get size - larger is better for our sort
            size = release.get("size", 0)

            # Return tuple - sort by quality score desc, then size desc
            return (-quality_score, -size)

        releases = sorted(releases, key=sort_key)

        # Calculate dynamic width for title column
        max_title_len = (
            max(len(r.get("title", "")) for r in releases) if releases else 80
        )
        # Limit to 165 chars max, but use actual max + 3 if shorter
        title_width = min(165, max_title_len + 3)

        # Keep all releases, don't limit
        filtered_releases = releases
        search_term = ""

        while True:
            # Create table
            table_title = f"Available Releases for: {title}"
            if search_term:
                table_title += f" [Filter: '{search_term}']"

            table = Table(title=table_title, box=box.ROUNDED)
            table.add_column("#", style="cyan", width=3)
            table.add_column("Title", style="white", overflow="fold", width=title_width)
            table.add_column("Size", style="green", width=10)
            table.add_column("Quality", style="yellow", width=15)
            table.add_column("Indexer", style="blue", width=6)

            # Apply filter if search term exists
            if search_term:
                filtered_releases = [
                    r
                    for r in releases
                    if search_term.lower() in r.get("title", "").lower()
                ]
            else:
                filtered_releases = releases

            if not filtered_releases:
                self.console.print(
                    f"[yellow]No releases match filter: '{search_term}'[/yellow]"
                )
                self.console.print("[cyan]Press Enter to clear filter[/cyan]")
                input()
                search_term = ""
                continue

            for idx, release in enumerate(filtered_releases, 1):
                title_text = release.get("title", "Unknown")
                # Truncate only if longer than calculated width
                if len(title_text) > title_width:
                    title_text = title_text[:title_width]

                size = release.get("size", 0)
                size_gb = f"{size / (1024**3):.2f} GB" if size > 0 else "Unknown"
                quality = (
                    release.get("quality", {}).get("quality", {}).get("name", "Unknown")
                )
                indexer = release.get("indexer", "Unknown")[:6]  # Truncate to 6 chars

                table.add_row(str(idx), title_text, size_gb, quality, indexer)

            self.console.print(table)
            self.console.print(
                f"\n[dim]Showing {len(filtered_releases)} of {len(releases)} releases[/dim]"
            )

            # Ask user to select
            self.console.print("\n[bold cyan]Options:[/bold cyan]")
            self.console.print("  Enter release number to download")
            self.console.print("  Enter 's' to search/filter releases")
            self.console.print("  Enter 'c' to clear filter")
            self.console.print("  Enter 0 to skip (and remember permanently)")
            self.console.print("  Enter -1 to keep current file")

            try:
                choice_input = input("\n[Your choice]: ").strip()

                if choice_input.lower() == "s":
                    # Search/filter mode
                    search_term = input("Enter search term: ").strip()
                    continue
                elif choice_input.lower() == "c":
                    # Clear filter
                    search_term = ""
                    continue

                choice = int(choice_input)

                if choice == -1:
                    return None  # Keep current
                elif choice == 0:
                    # Skip and remember permanently
                    self._add_skipped_file(file_path)
                    self.save_caches()
                    self.console.print("[yellow]Marked to skip permanently[/yellow]")
                    return None
                elif 1 <= choice <= len(filtered_releases):
                    return filtered_releases[choice - 1]
                else:
                    self.console.print("[red]Invalid choice[/red]")
                    continue
            except ValueError:
                self.console.print(
                    "[red]Invalid input. Please enter a number, 's', 'c', 0, or -1[/red]"
                )
                continue
            except KeyboardInterrupt:
                self.console.print("\n[yellow]Skipped[/yellow]")
                return None

    def download_release(
        self, url: str, api_key: str, release: dict, is_sonarr: bool = True
    ) -> bool:
        """Download a specific release"""
        try:
            guid = release.get("guid")
            indexer_id = release.get("indexerId")

            if not guid or indexer_id is None:
                self.console.print("[red]Invalid release data[/red]")
                return False

            data = {"guid": guid, "indexerId": indexer_id}

            result = self._make_request(
                url,
                api_key,
                "release",
                method="POST",
                data=data,
                auth=self.sonarr_http_auth if is_sonarr else self.radarr_http_auth,
            )

            if result is not None:
                self.console.print("[green]+ Download queued successfully[/green]")
                return True
            else:
                self.console.print("[red]- Failed to queue download[/red]")
                return False

        except Exception as e:
            self.console.print(f"[red]Error downloading release: {e}[/red]")
            return False

    def _delete_file_and_trigger_search(
        self,
        *,
        url: str,
        api_key: str,
        file_endpoint: str,
        auth: tuple[str, str] | None,
        search_payload: dict | None,
        trigger_message: str,
        missing_search_message: str | None = None,
    ) -> None:
        """Delete a media file and optionally trigger an arr search command."""
        print("     Deleting file to trigger re-download...")
        self._make_request(
            url,
            api_key,
            file_endpoint,
            method="DELETE",
            auth=auth,
        )
        if search_payload:
            print(f"     {trigger_message}")
            self._make_request(
                url,
                api_key,
                "command",
                method="POST",
                data=search_payload,
                auth=auth,
            )
            return
        if missing_search_message:
            print(f"     {missing_search_message}")

    def process_sonarr(self, dry_run: bool = False):
        """Process all Sonarr series and check episode files"""
        if self.interactive:
            self.console.print(
                Panel("[bold cyan]Processing Sonarr[/bold cyan]", box=box.DOUBLE)
            )
        else:
            print("\n=== Processing Sonarr ===")

        # Get all series
        series_list = self._make_request(
            self.sonarr_url,
            self.sonarr_api,
            "series",
            auth=self.sonarr_http_auth,
        )
        if not series_list:
            msg = "Failed to fetch series from Sonarr"
            if self.interactive:
                self.console.print(f"[red]{msg}[/red]")
            else:
                print(msg)
            return

        if self.interactive:
            self.console.print(f"[green]Found {len(series_list)} series[/green]\n")
        else:
            print(f"Found {len(series_list)} series")

        for series in series_list:
            series_id = series["id"]
            series_title = series["title"]
            quality_profile_id = series.get("qualityProfileId")

            # Get episode files for this series
            episode_files = self._make_request(
                self.sonarr_url,
                self.sonarr_api,
                f"episodefile?seriesId={series_id}",
                auth=self.sonarr_http_auth,
            )

            if not episode_files:
                continue

            if self.interactive:
                self.console.print(
                    f"\n[bold white]Checking series:[/bold white] {series_title} [dim]({len(episode_files)} files)[/dim]"
                )
            else:
                print(f"\nChecking series: {series_title} ({len(episode_files)} files)")

            for ep_file in episode_files:
                file_path = ep_file.get("path")
                file_id = ep_file.get("id")

                if not file_path or not file_id:
                    continue

                # Check if file is already in good files cache
                if self._is_cached_match(self._good_files_map, file_path):
                    if self.interactive:
                        self.console.print(
                            f"[dim]Skipping (already verified as OK): {Path(file_path).name}[/dim]"
                        )
                    continue

                # Check for English streams
                has_eng_audio, has_eng_subs = self.check_file_streams(file_path)

                if self.should_redownload(has_eng_audio, has_eng_subs):
                    filename = Path(file_path).name

                    # Check if this file was previously skipped BEFORE showing anything
                    if self._is_cached_match(self._skipped_files_map, file_path):
                        if self.interactive:
                            self.console.print(
                                f"[dim]Skipping (previously marked to skip): {filename}[/dim]"
                            )
                        continue

                    if self.interactive:
                        # Interactive mode - show details and ask for confirmation
                        self.console.print(f"\n[red]X Issue found:[/red] {filename}")
                        self.console.print(
                            f"  [yellow]English audio:[/yellow] {'YES' if has_eng_audio else 'NO'}"
                        )
                        self.console.print(
                            f"  [yellow]English subs:[/yellow] {'YES' if has_eng_subs else 'NO'}"
                        )

                        # Get episode IDs for this file
                        episode_ids = self.get_episodes_for_file(series_id, file_id)

                        if episode_ids:
                            # Ask user what to do (default is No when pressing Enter)
                            view_alternatives = Confirm.ask(
                                "\n[bold cyan]View alternative releases?[/bold cyan]",
                                default=False,
                            )

                            if view_alternatives:
                                # Get releases for the first episode (they should be the same for multi-episode files)
                                releases = self.get_episode_releases(
                                    episode_ids[0], quality_profile_id
                                )
                                selected_release = self.display_releases_and_select(
                                    releases, filename, file_path
                                )

                                if selected_release and not dry_run:
                                    # Delete current file
                                    self.console.print(
                                        "[yellow]Deleting current file...[/yellow]"
                                    )
                                    self._make_request(
                                        self.sonarr_url,
                                        self.sonarr_api,
                                        f"episodefile/{file_id}",
                                        method="DELETE",
                                        auth=self.sonarr_http_auth,
                                    )

                                    # Download selected release
                                    self.download_release(
                                        self.sonarr_url,
                                        self.sonarr_api,
                                        selected_release,
                                        is_sonarr=True,
                                    )
                                elif selected_release and dry_run:
                                    self.console.print(
                                        "[dim]DRY RUN: Would download selected release[/dim]"
                                    )
                            else:
                                # User said No - save to cache to skip this file next time
                                self._add_skipped_file(file_path)
                                self.save_caches()
                                self.console.print(
                                    "[yellow]Marked to skip permanently[/yellow]"
                                )
                        else:
                            self.console.print(
                                "[yellow]No episode IDs found for this file[/yellow]"
                            )
                    else:
                        # Non-interactive mode - original behavior
                        print(f"  X {filename}")
                        print(
                            f"     English audio: {has_eng_audio}, English subs: {has_eng_subs}"
                        )

                        if not dry_run:
                            # Capture episode ids before deleting the file mapping.
                            episode_ids = self.get_episodes_for_file(series_id, file_id)
                            self._delete_file_and_trigger_search(
                                url=self.sonarr_url,
                                api_key=self.sonarr_api,
                                file_endpoint=f"episodefile/{file_id}",
                                auth=self.sonarr_http_auth,
                                search_payload={
                                    "name": "EpisodeSearch",
                                    "episodeIds": episode_ids,
                                }
                                if episode_ids
                                else None,
                                trigger_message="Triggering episode search...",
                                missing_search_message="Could not resolve episode IDs before delete; search not triggered",
                            )
                        else:
                            print("     [DRY RUN] Would delete and re-download")
                else:
                    # File is OK - add to cache
                    self._add_good_file(file_path)
                    self.save_caches()

                    if self.interactive:
                        self.console.print(f"[green]OK[/green] {Path(file_path).name}")
                    else:
                        print(f"  OK {Path(file_path).name}")

    def process_radarr(self, dry_run: bool = False):
        """Process all Radarr movies and check movie files"""
        if self.interactive:
            self.console.print(
                Panel("[bold cyan]Processing Radarr[/bold cyan]", box=box.DOUBLE)
            )
        else:
            print("\n=== Processing Radarr ===")

        # Get all movies
        movies = self._make_request(
            self.radarr_url,
            self.radarr_api,
            "movie",
            auth=self.radarr_http_auth,
        )
        if not movies:
            msg = "Failed to fetch movies from Radarr"
            if self.interactive:
                self.console.print(f"[red]{msg}[/red]")
            else:
                print(msg)
            return

        if self.interactive:
            self.console.print(f"[green]Found {len(movies)} movies[/green]\n")
        else:
            print(f"Found {len(movies)} movies")

        for movie in movies:
            if not movie.get("hasFile"):
                continue

            movie_id = movie["id"]
            movie_title = movie["title"]
            quality_profile_id = movie.get("qualityProfileId")
            movie_file = movie.get("movieFile", {})
            file_path = movie_file.get("path")
            file_id = movie_file.get("id")

            if not file_path or not file_id:
                continue

            # Check if file is already in good files cache
            if self._is_cached_match(self._good_files_map, file_path):
                if self.interactive:
                    self.console.print(
                        f"[dim]Skipping (already verified as OK): {movie_title}[/dim]"
                    )
                continue

            # Check for English streams
            has_eng_audio, has_eng_subs = self.check_file_streams(file_path)

            if self.should_redownload(has_eng_audio, has_eng_subs):
                filename = Path(file_path).name

                # Check if this file was previously skipped BEFORE showing anything
                if self._is_cached_match(self._skipped_files_map, file_path):
                    if self.interactive:
                        self.console.print(
                            f"[dim]Skipping (previously marked to skip): {movie_title}[/dim]"
                        )
                    continue

                if self.interactive:
                    # Interactive mode - show details and ask for confirmation
                    self.console.print(f"\n[red]X Issue found:[/red] {movie_title}")
                    self.console.print(f"  [dim]File:[/dim] {filename}")
                    self.console.print(
                        f"  [yellow]English audio:[/yellow] {'YES' if has_eng_audio else 'NO'}"
                    )
                    self.console.print(
                        f"  [yellow]English subs:[/yellow] {'YES' if has_eng_subs else 'NO'}"
                    )

                    # Ask user what to do (default is No when pressing Enter)
                    view_alternatives = Confirm.ask(
                        "\n[bold cyan]View alternative releases?[/bold cyan]",
                        default=False,
                    )

                    if view_alternatives:
                        # Get releases
                        releases = self.get_movie_releases(movie_id, quality_profile_id)
                        selected_release = self.display_releases_and_select(
                            releases, movie_title, file_path
                        )

                        if selected_release and not dry_run:
                            # Delete current file
                            self.console.print(
                                "[yellow]Deleting current file...[/yellow]"
                            )
                            self._make_request(
                                self.radarr_url,
                                self.radarr_api,
                                f"moviefile/{file_id}",
                                method="DELETE",
                                auth=self.radarr_http_auth,
                            )

                            # Download selected release
                            self.download_release(
                                self.radarr_url,
                                self.radarr_api,
                                selected_release,
                                is_sonarr=False,
                            )
                        elif selected_release and dry_run:
                            self.console.print(
                                "[dim]DRY RUN: Would download selected release[/dim]"
                            )
                    else:
                        # User said No - save to cache to skip this file next time
                        self._add_skipped_file(file_path)
                        self.save_caches()
                        self.console.print(
                            "[yellow]Marked to skip permanently[/yellow]"
                        )
                else:
                    # Non-interactive mode - original behavior
                    print(f"  X {movie_title}")
                    print(f"     File: {filename}")
                    print(
                        f"     English audio: {has_eng_audio}, English subs: {has_eng_subs}"
                    )

                    if not dry_run:
                        self._delete_file_and_trigger_search(
                            url=self.radarr_url,
                            api_key=self.radarr_api,
                            file_endpoint=f"moviefile/{file_id}",
                            auth=self.radarr_http_auth,
                            search_payload={
                                "name": "MoviesSearch",
                                "movieIds": [movie_id],
                            },
                            trigger_message="Triggering movie search...",
                        )
                    else:
                        print("     [DRY RUN] Would delete and re-download")
            else:
                # File is OK - add to cache
                self._add_good_file(file_path)
                self.save_caches()

                if self.interactive:
                    self.console.print(f"[green]OK[/green] {movie_title}")
                else:
                    print(f"  OK {movie_title}")
