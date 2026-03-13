"""Background workers for Sonarr UI."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, TypedDict, cast

from PySide6.QtCore import QThread, Signal

from .probe_cache import ProbeInfo, probe_file, save_probe_cache

if TYPE_CHECKING:
    from collections.abc import Callable

    from .api import JsonDict, JsonList, SonarrAPI


class EpisodeEntry(TypedDict):
    """Normalized episode-file payload used to populate the tree model."""

    episode_number: int
    file_name: str
    file_path: str
    file_id: int | None
    size_bytes: int
    video_resolution: str
    video_bitrate: str
    video_codec: str
    hdr: str
    audio_codec: str
    audio_bitrate: str
    audio_langs: str
    sub_langs: str
    episode_data: JsonList


class LoadedSeriesEntry(TypedDict):
    """Normalized series payload emitted back to the UI thread."""

    series_id: int
    title: str
    year: int | str
    path: str
    total_size: int
    seasons: dict[int, list[EpisodeEntry]]
    missing_seasons: dict[int, JsonList]
    all_season_nums: list[int]
    series_data: JsonDict


def _as_int(value: object) -> int:
    """Normalize Sonarr payload integers that may arrive as loose JSON values."""

    return int(value) if isinstance(value, int | float | str) else 0


class LoadWorker(QThread):
    """Fetches series + episode data from Sonarr, probes files, emits results."""

    progress = Signal(str)
    series_ready = Signal(list)

    def __init__(self, api: SonarrAPI) -> None:
        super().__init__()
        self.api = api

    def run(self) -> None:
        try:
            self.progress.emit("Fetching series listâ€¦")
            try:
                all_series = self.api.get_series()
            except Exception as exc:
                self.progress.emit(f"Error: {exc}")
                self.series_ready.emit([])
                return

            result: list[LoadedSeriesEntry] = []
            series_error_count = 0
            for idx, series in enumerate(all_series):
                if self.isInterruptionRequested():
                    return
                series_id_obj = series.get("id")
                if not isinstance(series_id_obj, int):
                    series_error_count += 1
                    self.progress.emit("Warning: skipped a series with missing id")
                    continue
                series_id = series_id_obj
                title = str(series.get("title", "?"))
                year_obj = series.get("year", "")
                year = year_obj if isinstance(year_obj, (int, str)) else ""
                series_path = str(series.get("path", ""))

                self.progress.emit(f"Loading {idx + 1}/{len(all_series)}: {title}")

                series_failed = False
                try:
                    ep_files = self.api.get_episode_files(series_id)
                except Exception as exc:
                    self.progress.emit(
                        f'Warning: failed episode files for "{title}": {exc}'
                    )
                    series_failed = True
                    ep_files = []

                try:
                    episodes = self.api.get_episodes(series_id)
                except Exception as exc:
                    self.progress.emit(f'Warning: failed episodes for "{title}": {exc}')
                    series_failed = True
                    episodes = []

                if series_failed:
                    series_error_count += 1
                    continue

                file_to_eps: dict[int, JsonList] = {}
                for episode in episodes:
                    episode_file_id = episode.get("episodeFileId", 0)
                    if isinstance(episode_file_id, int) and episode_file_id:
                        file_to_eps.setdefault(episode_file_id, []).append(episode)

                seasons: dict[int, list[EpisodeEntry]] = {}
                for episode_file in ep_files:
                    if self.isInterruptionRequested():
                        return
                    file_path = str(episode_file.get("path", ""))
                    file_id_obj = episode_file.get("id")
                    file_id = file_id_obj if isinstance(file_id_obj, int) else None
                    eps_for_file = (
                        file_to_eps.get(file_id, []) if file_id is not None else []
                    )
                    season_num = (
                        _as_int(eps_for_file[0].get("seasonNumber", 0))
                        if eps_for_file
                        else 0
                    )
                    ep_num = (
                        _as_int(eps_for_file[0].get("episodeNumber", 0))
                        if eps_for_file
                        else 0
                    )

                    probe: ProbeInfo = probe_file(file_path)
                    audio_langs = probe["audio_langs"]
                    sub_langs = probe["sub_langs"]
                    ep_entry: EpisodeEntry = {
                        "episode_number": ep_num,
                        "file_name": Path(file_path).name if file_path else "",
                        "file_path": file_path,
                        "file_id": file_id,
                        "size_bytes": int(probe["size_bytes"]),
                        "video_resolution": str(probe["video_resolution"]),
                        "video_bitrate": str(probe["video_bitrate"]),
                        "video_codec": str(probe["video_codec"]),
                        "hdr": str(probe["hdr"]),
                        "audio_codec": str(probe["audio_codec"]),
                        "audio_bitrate": str(probe["audio_bitrate"]),
                        "audio_langs": ", ".join(dict.fromkeys(audio_langs)),
                        "sub_langs": ", ".join(dict.fromkeys(sub_langs)),
                        "episode_data": eps_for_file,
                    }
                    seasons.setdefault(season_num, []).append(ep_entry)

                for season_entries in seasons.values():
                    season_entries.sort(key=lambda entry: entry["episode_number"])

                missing_seasons: dict[int, JsonList] = {}
                for episode in episodes:
                    episode_file_id = episode.get("episodeFileId", 0)
                    if _as_int(episode_file_id) == 0:
                        season_num = _as_int(episode.get("seasonNumber", 0))
                        missing_seasons.setdefault(season_num, []).append(episode)
                for season_entries in missing_seasons.values():
                    season_entries.sort(
                        key=lambda entry: _as_int(entry.get("episodeNumber", 0))
                    )

                all_season_nums: set[int] = set()
                season_info_obj = series.get("seasons", [])
                season_info_list = (
                    cast("list[JsonDict]", season_info_obj)
                    if isinstance(season_info_obj, list)
                    else []
                )
                for season_info in season_info_list:
                    all_season_nums.add(_as_int(season_info.get("seasonNumber", 0)))
                for episode in episodes:
                    all_season_nums.add(_as_int(episode.get("seasonNumber", 0)))

                total_size = sum(
                    entry["size_bytes"]
                    for season_entries in seasons.values()
                    for entry in season_entries
                )
                result.append(
                    {
                        "series_id": series_id,
                        "title": title,
                        "year": year,
                        "path": series_path,
                        "total_size": total_size,
                        "seasons": seasons,
                        "missing_seasons": missing_seasons,
                        "all_season_nums": sorted(all_season_nums),
                        "series_data": series,
                    }
                )

            result.sort(key=lambda series: series["title"].lower())
            save_probe_cache()
            if series_error_count:
                self.progress.emit(
                    f"Done with warnings: loaded {len(result)} series, skipped {series_error_count}"
                )
            else:
                self.progress.emit("Done")
            self.series_ready.emit(result)
        except Exception as exc:
            if self.isInterruptionRequested():
                return
            self.progress.emit(f"Error: unexpected loader failure: {exc}")
            self.series_ready.emit([])


class ApiActionWorker(QThread):
    """Runs a Sonarr API action off the UI thread and returns result/error."""

    finished_action = Signal(object, object)

    def __init__(self, action: Callable[[], object]) -> None:
        super().__init__()
        self._action = action

    def run(self) -> None:
        try:
            result = self._action()
            self.finished_action.emit(result, None)
        except Exception as exc:
            self.finished_action.emit(None, exc)
