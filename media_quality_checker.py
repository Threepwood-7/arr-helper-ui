#!/usr/bin/env python3
"""
Media Quality Checker for Sonarr/Radarr
Checks downloaded files for English audio and subtitles, triggers re-download if missing.
"""

import os
import sys
import json
import subprocess
from typing import Dict, List, Optional, Tuple
import requests
from pathlib import Path
import tomli
from ffprobe_utils import find_ffprobe
from rich.console import Console
from rich.table import Table
from rich.prompt import Confirm, IntPrompt
from rich.panel import Panel
from rich import box


class Config:
    """Load and manage configuration from TOML file"""
    
    def __init__(self, config_path: str = 'config.toml'):
        self._script_dir = os.path.dirname(os.path.abspath(__file__))
        self.config_path = os.path.join(self._script_dir, config_path) if not os.path.isabs(config_path) else config_path
        self.config = self._load_config()
        self.user_cache_path = os.path.join(self._script_dir, 'z_user.cache')
        self.files_cache_path = os.path.join(self._script_dir, 'z_files.cache')
    
    def _load_config(self) -> Dict:
        """Load configuration from TOML file"""
        if not os.path.exists(self.config_path):
            print(f"Error: Config file not found: {self.config_path}")
            print("Creating example config file...")
            self._create_example_config()
            sys.exit(1)
        
        try:
            with open(self.config_path, 'rb') as f:
                return tomli.load(f)
        except Exception as e:
            print(f"Error loading config file: {e}")
            sys.exit(1)
    
    def load_user_cache(self) -> Dict:
        """Load user decisions cache"""
        if os.path.exists(self.user_cache_path):
            try:
                with open(self.user_cache_path, 'r') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Warning: Could not load user cache: {e}")
                return {}
        return {}
    
    def save_user_cache(self, cache: Dict):
        """Save user decisions cache"""
        try:
            with open(self.user_cache_path, 'w') as f:
                json.dump(cache, f, indent=2)
        except Exception as e:
            print(f"Warning: Could not save user cache: {e}")
    
    def load_files_cache(self) -> Dict:
        """Load good files cache"""
        if os.path.exists(self.files_cache_path):
            try:
                with open(self.files_cache_path, 'r') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Warning: Could not load files cache: {e}")
                return {}
        return {}
    
    def save_files_cache(self, cache: Dict):
        """Save good files cache"""
        try:
            with open(self.files_cache_path, 'w') as f:
                json.dump(cache, f, indent=2)
        except Exception as e:
            print(f"Warning: Could not save files cache: {e}")
    
    def _create_example_config(self):
        """Create an example config file by copying config.example.toml."""
        example_path = os.path.join(self._script_dir, 'config.example.toml')
        try:
            if os.path.exists(example_path):
                import shutil
                shutil.copy2(example_path, self.config_path)
            else:
                example_config = """# Media Quality Checker Configuration

[sonarr]
url = "http://localhost:8989"
api_key = "your-sonarr-api-key-here"
enabled = true
# http_basic_auth_username = ""
# http_basic_auth_password = ""

[radarr]
url = "http://localhost:7878"
api_key = "your-radarr-api-key-here"
enabled = true
# http_basic_auth_username = ""
# http_basic_auth_password = ""

[settings]
# Run in dry-run mode (no actual changes)
dry_run = false

# Interactive mode - ask for confirmation and show alternatives
interactive = true

# Require English audio stream
require_english_audio = true

# Require English subtitles
require_english_subs = true

# Language codes to consider as English
english_language_codes = ["eng", "en", "english"]

# Highlight episodes missing subtitles (light red background in UI)
# The value is a label; all codes from english_language_codes count as a match
# highlight_missing_subs = "english"
"""
                with open(self.config_path, 'w') as f:
                    f.write(example_config)
            print(f"Created example config at: {self.config_path}")
            print("Please edit this file with your Sonarr/Radarr settings.")
        except Exception as e:
            print(f"Error creating example config: {e}")
    
    def get_sonarr_config(self) -> Optional[Dict]:
        """Get Sonarr configuration"""
        sonarr = self.config.get('sonarr', {})
        if not sonarr.get('enabled', True):
            return None
        return sonarr
    
    def get_radarr_config(self) -> Optional[Dict]:
        """Get Radarr configuration"""
        radarr = self.config.get('radarr', {})
        if not radarr.get('enabled', True):
            return None
        return radarr
    
    def get_settings(self) -> Dict:
        """Get general settings"""
        return self.config.get('settings', {})


class MediaQualityChecker:
    def __init__(self, sonarr_url: str, sonarr_api: str, radarr_url: str, radarr_api: str,
                 require_audio: bool = True, require_subs: bool = True,
                 english_codes: List[str] = None, interactive: bool = False,
                 config: 'Config' = None,
                 sonarr_http_auth: tuple = None, radarr_http_auth: tuple = None,
                 ffprobe_path: str = 'ffprobe'):
        self.sonarr_url = sonarr_url.rstrip('/')
        self.sonarr_api = sonarr_api
        self.radarr_url = radarr_url.rstrip('/')
        self.radarr_api = radarr_api
        self.sonarr_http_auth = sonarr_http_auth
        self.radarr_http_auth = radarr_http_auth
        self.require_audio = require_audio
        self.require_subs = require_subs
        self.english_codes = [code.lower() for code in (english_codes or ['eng', 'en', 'english'])]
        self.interactive = interactive
        self.console = Console()
        self.config = config
        self.ffprobe_path = ffprobe_path
        
        # Load caches
        self.user_cache = config.load_user_cache() if config else {}
        self.files_cache = config.load_files_cache() if config else {}

        # Build sets for O(1) lookups (JSON stores lists, we use sets in memory)
        self._good_files_set = set(self.files_cache.get('good_files', []))
        self._skipped_files_set = set(self.user_cache.get('skipped_files', []))
    
    def _add_good_file(self, file_path: str):
        """Add a file to the good files cache."""
        if file_path not in self._good_files_set:
            self._good_files_set.add(file_path)
            self.files_cache.setdefault('good_files', []).append(file_path)

    def _add_skipped_file(self, file_path: str):
        """Add a file to the skipped files cache."""
        if file_path not in self._skipped_files_set:
            self._skipped_files_set.add(file_path)
            self.user_cache.setdefault('skipped_files', []).append(file_path)

    def save_caches(self):
        """Save user cache and files cache to disk"""
        if self.config:
            self.config.save_user_cache(self.user_cache)
            self.config.save_files_cache(self.files_cache)
        
    def _make_request(self, url: str, api_key: str, endpoint: str, method: str = 'GET', data: Dict = None) -> Optional[Dict]:
        """Make API request to Sonarr/Radarr"""
        headers = {'X-Api-Key': api_key}
        full_url = f"{url}/api/v3/{endpoint}"
        auth = self.sonarr_http_auth if url == self.sonarr_url else self.radarr_http_auth

        try:
            if method == 'GET':
                response = requests.get(full_url, headers=headers, auth=auth, timeout=300)
            elif method == 'PUT':
                response = requests.put(full_url, headers=headers, auth=auth, json=data, timeout=300)
            elif method == 'DELETE':
                response = requests.delete(full_url, headers=headers, auth=auth, timeout=300)
            else:
                response = requests.post(full_url, headers=headers, auth=auth, json=data, timeout=300)
            
            response.raise_for_status()
            return response.json() if response.text else {}
        except requests.exceptions.RequestException as e:
            print(f"Error making request to {full_url}: {e}")
            return None

    def check_file_streams(self, file_path: str) -> Tuple[bool, bool]:
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
                '-v', 'quiet',
                '-print_format', 'json',
                '-show_streams',
                file_path
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
            
            if result.returncode != 0:
                print(f"ffprobe error for {file_path}: {result.stderr}")
                return False, False
            
            data = json.loads(result.stdout)
            streams = data.get('streams', [])
            
            has_eng_audio = False
            has_eng_subs = False
            
            for stream in streams:
                codec_type = stream.get('codec_type', '')
                tags = stream.get('tags', {})
                language = tags.get('language', '').lower()
                
                # Check for English audio
                if codec_type == 'audio':
                    if language in self.english_codes:
                        has_eng_audio = True
                
                # Check for English subtitles
                if codec_type == 'subtitle':
                    if language in self.english_codes:
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
        if self.require_subs and not has_eng_subs:
            return True
        return False
    
    def get_episode_releases(self, episode_id: int, quality_profile_id: int = None) -> List[Dict]:
        """Get available releases for an episode from Sonarr"""
        endpoint = f'release?episodeId={episode_id}'
        if quality_profile_id:
            endpoint += f'&qualityProfileId={quality_profile_id}'
        
        releases = self._make_request(
            self.sonarr_url,
            self.sonarr_api,
            endpoint
        )
        return releases or []
    
    def get_episodes_for_file(self, series_id: int, episode_file_id: int) -> List[int]:
        """Get episode IDs associated with an episode file"""
        episodes = self._make_request(
            self.sonarr_url,
            self.sonarr_api,
            f'episode?seriesId={series_id}'
        )
        
        if not episodes:
            return []
        
        # Find episodes that use this file
        episode_ids = []
        for ep in episodes:
            if ep.get('episodeFileId') == episode_file_id:
                episode_ids.append(ep.get('id'))
        
        return episode_ids
    
    def get_movie_releases(self, movie_id: int, quality_profile_id: int = None) -> List[Dict]:
        """Get available releases for a movie from Radarr"""
        endpoint = f'release?movieId={movie_id}'
        if quality_profile_id:
            endpoint += f'&qualityProfileId={quality_profile_id}'
        
        releases = self._make_request(
            self.radarr_url,
            self.radarr_api,
            endpoint
        )
        return releases or []
    
    def display_releases_and_select(self, releases: List[Dict], title: str, file_path: str) -> Optional[Dict]:
        """Display releases in a table and let user select one"""
        if not releases:
            self.console.print("[yellow]No alternative releases found[/yellow]")
            return None
        
        # Sort releases by quality profile match (preferred first) then size descending
        def sort_key(release):
            # Get quality score - higher is better
            quality = release.get('quality', {}).get('quality', {})
            quality_score = quality.get('id', 0)
            
            # Get size - larger is better for our sort
            size = release.get('size', 0)
            
            # Return tuple - sort by quality score desc, then size desc
            return (-quality_score, -size)
        
        releases = sorted(releases, key=sort_key)
        
        # Calculate dynamic width for title column
        max_title_len = max(len(r.get('title', '')) for r in releases) if releases else 80
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
                    r for r in releases 
                    if search_term.lower() in r.get('title', '').lower()
                ]
            else:
                filtered_releases = releases
            
            if not filtered_releases:
                self.console.print(f"[yellow]No releases match filter: '{search_term}'[/yellow]")
                self.console.print("[cyan]Press Enter to clear filter[/cyan]")
                input()
                search_term = ""
                continue
            
            for idx, release in enumerate(filtered_releases, 1):
                title_text = release.get('title', 'Unknown')
                # Truncate only if longer than calculated width
                if len(title_text) > title_width:
                    title_text = title_text[:title_width]
                
                size = release.get('size', 0)
                size_gb = f"{size / (1024**3):.2f} GB" if size > 0 else "Unknown"
                quality = release.get('quality', {}).get('quality', {}).get('name', 'Unknown')
                indexer = release.get('indexer', 'Unknown')[:6]  # Truncate to 6 chars
                
                table.add_row(
                    str(idx),
                    title_text,
                    size_gb,
                    quality,
                    indexer
                )
            
            self.console.print(table)
            self.console.print(f"\n[dim]Showing {len(filtered_releases)} of {len(releases)} releases[/dim]")
            
            # Ask user to select
            self.console.print("\n[bold cyan]Options:[/bold cyan]")
            self.console.print("  Enter release number to download")
            self.console.print("  Enter 's' to search/filter releases")
            self.console.print("  Enter 'c' to clear filter")
            self.console.print("  Enter 0 to skip (and remember permanently)")
            self.console.print("  Enter -1 to keep current file")
            
            try:
                choice_input = input("\n[Your choice]: ").strip()
                
                if choice_input.lower() == 's':
                    # Search/filter mode
                    search_term = input("Enter search term: ").strip()
                    continue
                elif choice_input.lower() == 'c':
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
                    self.console.print(f"[yellow]Marked to skip permanently[/yellow]")
                    return None
                elif 1 <= choice <= len(filtered_releases):
                    return filtered_releases[choice - 1]
                else:
                    self.console.print("[red]Invalid choice[/red]")
                    continue
            except ValueError:
                self.console.print("[red]Invalid input. Please enter a number, 's', 'c', 0, or -1[/red]")
                continue
            except KeyboardInterrupt:
                self.console.print("\n[yellow]Skipped[/yellow]")
                return None
    
    def download_release(self, url: str, api_key: str, release: Dict, is_sonarr: bool = True) -> bool:
        """Download a specific release"""
        try:
            guid = release.get('guid')
            indexer_id = release.get('indexerId')
            
            if not guid or indexer_id is None:
                self.console.print("[red]Invalid release data[/red]")
                return False
            
            data = {
                'guid': guid,
                'indexerId': indexer_id
            }
            
            result = self._make_request(
                url,
                api_key,
                'release',
                method='POST',
                data=data
            )
            
            if result:
                self.console.print("[green]+ Download queued successfully[/green]")
                return True
            else:
                self.console.print("[red]- Failed to queue download[/red]")
                return False
                
        except Exception as e:
            self.console.print(f"[red]Error downloading release: {e}[/red]")
            return False

    def process_sonarr(self, dry_run: bool = False):
        """Process all Sonarr series and check episode files"""
        if self.interactive:
            self.console.print(Panel("[bold cyan]Processing Sonarr[/bold cyan]", box=box.DOUBLE))
        else:
            print("\n=== Processing Sonarr ===")
        
        # Get all series
        series_list = self._make_request(self.sonarr_url, self.sonarr_api, 'series')
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
            series_id = series['id']
            series_title = series['title']
            quality_profile_id = series.get('qualityProfileId')
            
            # Get episode files for this series
            episode_files = self._make_request(
                self.sonarr_url, 
                self.sonarr_api, 
                f'episodefile?seriesId={series_id}'
            )
            
            if not episode_files:
                continue
            
            if self.interactive:
                self.console.print(f"\n[bold white]Checking series:[/bold white] {series_title} [dim]({len(episode_files)} files)[/dim]")
            else:
                print(f"\nChecking series: {series_title} ({len(episode_files)} files)")
            
            for ep_file in episode_files:
                file_path = ep_file.get('path')
                file_id = ep_file.get('id')
                
                if not file_path or not file_id:
                    continue
                
                # Check if file is already in good files cache
                if file_path in self._good_files_set:
                    if self.interactive:
                        self.console.print(f"[dim]Skipping (already verified as OK): {Path(file_path).name}[/dim]")
                    continue

                # Check for English streams
                has_eng_audio, has_eng_subs = self.check_file_streams(file_path)

                if self.should_redownload(has_eng_audio, has_eng_subs):
                    filename = Path(file_path).name

                    # Check if this file was previously skipped BEFORE showing anything
                    if file_path in self._skipped_files_set:
                        if self.interactive:
                            self.console.print(f"[dim]Skipping (previously marked to skip): {filename}[/dim]")
                        continue

                    if self.interactive:
                        # Interactive mode - show details and ask for confirmation
                        self.console.print(f"\n[red]X Issue found:[/red] {filename}")
                        self.console.print(f"  [yellow]English audio:[/yellow] {'YES' if has_eng_audio else 'NO'}")
                        self.console.print(f"  [yellow]English subs:[/yellow] {'YES' if has_eng_subs else 'NO'}")

                        # Get episode IDs for this file
                        episode_ids = self.get_episodes_for_file(series_id, file_id)

                        if episode_ids:
                            # Ask user what to do (default is No when pressing Enter)
                            view_alternatives = Confirm.ask("\n[bold cyan]View alternative releases?[/bold cyan]", default=False)

                            if view_alternatives:
                                # Get releases for the first episode (they should be the same for multi-episode files)
                                releases = self.get_episode_releases(episode_ids[0], quality_profile_id)
                                selected_release = self.display_releases_and_select(releases, filename, file_path)

                                if selected_release and not dry_run:
                                    # Delete current file
                                    self.console.print("[yellow]Deleting current file...[/yellow]")
                                    self._make_request(
                                        self.sonarr_url,
                                        self.sonarr_api,
                                        f'episodefile/{file_id}',
                                        method='DELETE'
                                    )

                                    # Download selected release
                                    self.download_release(
                                        self.sonarr_url,
                                        self.sonarr_api,
                                        selected_release,
                                        is_sonarr=True
                                    )
                                elif selected_release and dry_run:
                                    self.console.print("[dim]DRY RUN: Would download selected release[/dim]")
                            else:
                                # User said No - save to cache to skip this file next time
                                self._add_skipped_file(file_path)
                                self.save_caches()
                                self.console.print(f"[yellow]Marked to skip permanently[/yellow]")
                        else:
                            self.console.print("[yellow]No episode IDs found for this file[/yellow]")
                    else:
                        # Non-interactive mode - original behavior
                        print(f"  X {filename}")
                        print(f"     English audio: {has_eng_audio}, English subs: {has_eng_subs}")

                        if not dry_run:
                            # Delete the episode file to trigger re-download
                            print(f"     Deleting file to trigger re-download...")
                            self._make_request(
                                self.sonarr_url,
                                self.sonarr_api,
                                f'episodefile/{file_id}',
                                method='DELETE'
                            )

                            # Trigger search for the episodes
                            episode_ids = self.get_episodes_for_file(series_id, file_id)
                            if episode_ids:
                                print(f"     Triggering episode search...")
                                self._make_request(
                                    self.sonarr_url,
                                    self.sonarr_api,
                                    'command',
                                    method='POST',
                                    data={'name': 'EpisodeSearch', 'episodeIds': episode_ids}
                                )
                        else:
                            print(f"     [DRY RUN] Would delete and re-download")
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
            self.console.print(Panel("[bold cyan]Processing Radarr[/bold cyan]", box=box.DOUBLE))
        else:
            print("\n=== Processing Radarr ===")
        
        # Get all movies
        movies = self._make_request(self.radarr_url, self.radarr_api, 'movie')
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
            if not movie.get('hasFile'):
                continue
            
            movie_id = movie['id']
            movie_title = movie['title']
            quality_profile_id = movie.get('qualityProfileId')
            movie_file = movie.get('movieFile', {})
            file_path = movie_file.get('path')
            file_id = movie_file.get('id')
            
            if not file_path or not file_id:
                continue
            
            # Check if file is already in good files cache
            if file_path in self._good_files_set:
                if self.interactive:
                    self.console.print(f"[dim]Skipping (already verified as OK): {movie_title}[/dim]")
                continue

            # Check for English streams
            has_eng_audio, has_eng_subs = self.check_file_streams(file_path)

            if self.should_redownload(has_eng_audio, has_eng_subs):
                filename = Path(file_path).name

                # Check if this file was previously skipped BEFORE showing anything
                if file_path in self._skipped_files_set:
                    if self.interactive:
                        self.console.print(f"[dim]Skipping (previously marked to skip): {movie_title}[/dim]")
                    continue

                if self.interactive:
                    # Interactive mode - show details and ask for confirmation
                    self.console.print(f"\n[red]X Issue found:[/red] {movie_title}")
                    self.console.print(f"  [dim]File:[/dim] {filename}")
                    self.console.print(f"  [yellow]English audio:[/yellow] {'YES' if has_eng_audio else 'NO'}")
                    self.console.print(f"  [yellow]English subs:[/yellow] {'YES' if has_eng_subs else 'NO'}")

                    # Ask user what to do (default is No when pressing Enter)
                    view_alternatives = Confirm.ask("\n[bold cyan]View alternative releases?[/bold cyan]", default=False)

                    if view_alternatives:
                        # Get releases
                        releases = self.get_movie_releases(movie_id, quality_profile_id)
                        selected_release = self.display_releases_and_select(releases, movie_title, file_path)

                        if selected_release and not dry_run:
                            # Delete current file
                            self.console.print("[yellow]Deleting current file...[/yellow]")
                            self._make_request(
                                self.radarr_url,
                                self.radarr_api,
                                f'moviefile/{file_id}',
                                method='DELETE'
                            )

                            # Download selected release
                            self.download_release(
                                self.radarr_url,
                                self.radarr_api,
                                selected_release,
                                is_sonarr=False
                            )
                        elif selected_release and dry_run:
                            self.console.print("[dim]DRY RUN: Would download selected release[/dim]")
                    else:
                        # User said No - save to cache to skip this file next time
                        self._add_skipped_file(file_path)
                        self.save_caches()
                        self.console.print(f"[yellow]Marked to skip permanently[/yellow]")
                else:
                    # Non-interactive mode - original behavior
                    print(f"  X {movie_title}")
                    print(f"     File: {filename}")
                    print(f"     English audio: {has_eng_audio}, English subs: {has_eng_subs}")

                    if not dry_run:
                        # Delete the movie file to trigger re-download
                        print(f"     Deleting file to trigger re-download...")
                        self._make_request(
                            self.radarr_url,
                            self.radarr_api,
                            f'moviefile/{file_id}',
                            method='DELETE'
                        )

                        # Trigger movie search
                        print(f"     Triggering movie search...")
                        self._make_request(
                            self.radarr_url,
                            self.radarr_api,
                            'command',
                            method='POST',
                            data={'name': 'MoviesSearch', 'movieIds': [movie_id]}
                        )
                    else:
                        print(f"     [DRY RUN] Would delete and re-download")
            else:
                # File is OK - add to cache
                self._add_good_file(file_path)
                self.save_caches()

                if self.interactive:
                    self.console.print(f"[green]OK[/green] {movie_title}")
                else:
                    print(f"  OK {movie_title}")


def main():
    # Load configuration
    config = Config('config.toml')
    
    # Check if ffprobe is available
    ffprobe_path = find_ffprobe()
    if not ffprobe_path:
        print("Error: ffprobe not found in PATH or common install locations.")
        print("Please install ffmpeg/ffprobe: https://ffmpeg.org/download.html")
        sys.exit(1)
    
    # Get settings
    settings = config.get_settings()
    dry_run = settings.get('dry_run', False)
    interactive = settings.get('interactive', False)
    require_audio = settings.get('require_english_audio', True)
    require_subs = settings.get('require_english_subs', True)
    english_codes = settings.get('english_language_codes', ['eng', 'en', 'english'])
    
    # Get Sonarr and Radarr configs
    sonarr_config = config.get_sonarr_config()
    radarr_config = config.get_radarr_config()
    
    if not sonarr_config and not radarr_config:
        print("Error: Both Sonarr and Radarr are disabled in config.")
        print("Please enable at least one in config.toml")
        sys.exit(1)
    
    # Create checker instance
    def _http_auth(cfg):
        u = cfg.get('http_basic_auth_username', '') if cfg else ''
        p = cfg.get('http_basic_auth_password', '') if cfg else ''
        return (u, p) if u else None

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
            checker.console.print(Panel(
                "[bold yellow]DRY RUN MODE[/bold yellow] - No changes will be made",
                box=box.DOUBLE
            ))
        else:
            checker.console.print("")
    else:
        if dry_run:
            print("=== DRY RUN MODE - No changes will be made ===\n")

    # Process Sonarr if enabled
    if sonarr_config:
        checker.process_sonarr(dry_run)

    # Process Radarr if enabled
    if radarr_config:
        checker.process_radarr(dry_run)

    if interactive:
        checker.console.print(Panel("[bold green]Complete[/bold green]", box=box.DOUBLE))
    else:
        print("\n=== Complete ===")


if __name__ == '__main__':
    main()
