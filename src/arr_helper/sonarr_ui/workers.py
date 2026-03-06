"""Background workers for Sonarr UI."""

from collections.abc import Callable
from pathlib import Path
from typing import Any

from PySide6.QtCore import QThread, Signal

from .api import SonarrAPI
from .probe_cache import _save_probe_cache, probe_file


class LoadWorker(QThread):
    """Fetches series + episode data from Sonarr, probes files, emits results."""
    progress = Signal(str)
    series_ready = Signal(list)

    def __init__(self, api: SonarrAPI):
        super().__init__()
        self.api = api

    def run(self):
        try:
            self.progress.emit('Fetching series list…')
            try:
                all_series = self.api.get_series()
            except Exception as e:
                self.progress.emit(f'Error: {e}')
                self.series_ready.emit([])
                return

            result = []
            series_error_count = 0
            for idx, series in enumerate(all_series):
                if self.isInterruptionRequested():
                    return
                series_id = series.get('id')
                if series_id is None:
                    series_error_count += 1
                    self.progress.emit('Warning: skipped a series with missing id')
                    continue
                title = series.get('title', '?')
                year = series.get('year', '')
                series_path = series.get('path', '')

                self.progress.emit(f'Loading {idx + 1}/{len(all_series)}: {title}')

                series_failed = False
                try:
                    ep_files = self.api.get_episode_files(series_id)
                except Exception as e:
                    self.progress.emit(f'Warning: failed episode files for "{title}": {e}')
                    series_failed = True
                    ep_files = []

                try:
                    episodes = self.api.get_episodes(series_id)
                except Exception as e:
                    self.progress.emit(f'Warning: failed episodes for "{title}": {e}')
                    series_failed = True
                    episodes = []

                if series_failed:
                    series_error_count += 1
                    continue

                # map episode_file_id -> episode metadata
                file_to_eps: dict[int, list] = {}
                for ep in episodes:
                    fid = ep.get('episodeFileId', 0)
                    if fid:
                        file_to_eps.setdefault(fid, []).append(ep)

                # group downloaded episodes by season
                seasons: dict[int, list] = {}
                for ef in ep_files:
                    if self.isInterruptionRequested():
                        return
                    file_path = ef.get('path', '')
                    file_id = ef.get('id')
                    eps_for_file = file_to_eps.get(file_id, [])
                    season_num = eps_for_file[0].get('seasonNumber', 0) if eps_for_file else 0
                    ep_num = eps_for_file[0].get('episodeNumber', 0) if eps_for_file else 0

                    probe = probe_file(file_path)

                    ep_entry = {
                        'episode_number': ep_num,
                        'file_name': Path(file_path).name if file_path else '',
                        'file_path': file_path,
                        'file_id': file_id,
                        'size_bytes': probe['size_bytes'],
                        'video_resolution': probe['video_resolution'],
                        'video_bitrate': probe['video_bitrate'],
                        'video_codec': probe['video_codec'],
                        'hdr': probe['hdr'],
                        'audio_codec': probe['audio_codec'],
                        'audio_bitrate': probe['audio_bitrate'],
                        'audio_langs': ', '.join(dict.fromkeys(probe['audio_langs'])),
                        'sub_langs': ', '.join(dict.fromkeys(probe['sub_langs'])),
                        'episode_data': eps_for_file,
                    }
                    seasons.setdefault(season_num, []).append(ep_entry)

                # sort episodes inside each season
                for sn in seasons:
                    seasons[sn].sort(key=lambda e: e['episode_number'])

                # collect missing episodes per season (no file)
                missing_seasons: dict[int, list] = {}
                for ep in episodes:
                    if ep.get('episodeFileId', 0) == 0:
                        sn = ep.get('seasonNumber', 0)
                        missing_seasons.setdefault(sn, []).append(ep)
                for sn in missing_seasons:
                    missing_seasons[sn].sort(key=lambda e: e.get('episodeNumber', 0))

                # collect all season numbers from the series metadata
                all_season_nums = set()
                for s_info in series.get('seasons', []):
                    all_season_nums.add(s_info.get('seasonNumber', 0))
                # also include seasons from episodes
                for ep in episodes:
                    all_season_nums.add(ep.get('seasonNumber', 0))

                total_size = sum(e['size_bytes'] for s in seasons.values() for e in s)

                result.append({
                    'series_id': series_id,
                    'title': title,
                    'year': year,
                    'path': series_path,
                    'total_size': total_size,
                    'seasons': seasons,
                    'missing_seasons': missing_seasons,
                    'all_season_nums': sorted(all_season_nums),
                    'series_data': series,
                })

            result.sort(key=lambda s: str(s.get('title', '')).lower())
            _save_probe_cache()
            if series_error_count:
                self.progress.emit(f'Done with warnings: loaded {len(result)} series, skipped {series_error_count}')
            else:
                self.progress.emit('Done')
            self.series_ready.emit(result)
        except Exception as e:
            if self.isInterruptionRequested():
                return
            self.progress.emit(f'Error: unexpected loader failure: {e}')
            self.series_ready.emit([])


class ApiActionWorker(QThread):
    """Runs a Sonarr API action off the UI thread and returns result/error."""
    finished_action = Signal(object, object)  # result, error

    def __init__(self, action: Callable[[], Any]):
        super().__init__()
        self._action = action

    def run(self):
        try:
            result = self._action()
            self.finished_action.emit(result, None)
        except Exception as e:
            self.finished_action.emit(None, e)
