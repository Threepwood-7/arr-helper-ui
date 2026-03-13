"""Sonarr API client."""

from __future__ import annotations

from typing import cast
from urllib.parse import quote

import requests

JsonDict = dict[str, object]
JsonList = list[JsonDict]
AuthTuple = tuple[str, str]
RequestTimeout = int | tuple[float, float]
DeleteParams = dict[str, str]


class SonarrAPI:
    """Wrap the Sonarr v3 HTTP API with typed convenience helpers."""

    def __init__(
        self,
        url: str,
        api_key: str,
        http_user: str = "",
        http_pass: str = "",
        request_timeout: RequestTimeout = 300,
    ) -> None:
        self.url = url.rstrip("/")
        self.api_key = api_key
        self.headers: dict[str, str] = {"X-Api-Key": api_key}
        self.auth: AuthTuple | None = (http_user, http_pass) if http_user else None
        self.timeout = request_timeout

    def _get(self, endpoint: str) -> JsonDict | JsonList:
        r = requests.get(
            f"{self.url}/api/v3/{endpoint}",
            headers=self.headers,
            auth=self.auth,
            timeout=self.timeout,
        )
        r.raise_for_status()
        return cast("JsonDict | JsonList", r.json())

    def _post(self, endpoint: str, data: JsonDict) -> JsonDict:
        r = requests.post(
            f"{self.url}/api/v3/{endpoint}",
            headers=self.headers,
            auth=self.auth,
            json=data,
            timeout=self.timeout,
        )
        r.raise_for_status()
        return cast("JsonDict", r.json()) if r.text else {}

    def _delete(
        self,
        endpoint: str,
        params: DeleteParams | None = None,
    ) -> requests.Response:
        r = requests.delete(
            f"{self.url}/api/v3/{endpoint}",
            headers=self.headers,
            auth=self.auth,
            params=params or {},
            timeout=self.timeout,
        )
        r.raise_for_status()
        return r

    def _put(self, endpoint: str, data: JsonDict) -> JsonDict:
        r = requests.put(
            f"{self.url}/api/v3/{endpoint}",
            headers=self.headers,
            auth=self.auth,
            json=data,
            timeout=self.timeout,
        )
        r.raise_for_status()
        return cast("JsonDict", r.json())

    def get_series(self) -> JsonList:
        return cast("JsonList", self._get("series"))

    def get_series_by_id(self, series_id: int) -> JsonDict:
        return cast("JsonDict", self._get(f"series/{series_id}"))

    def get_episodes(self, series_id: int) -> JsonList:
        return cast("JsonList", self._get(f"episode?seriesId={series_id}"))

    def get_episode_files(self, series_id: int) -> JsonList:
        return cast("JsonList", self._get(f"episodefile?seriesId={series_id}"))

    def update_series(self, series: JsonDict) -> JsonDict:
        return self._put(f"series/{series['id']}", series)

    def update_episode(self, episode: JsonDict) -> JsonDict:
        return self._put(f"episode/{episode['id']}", episode)

    def delete_series(self, series_id: int, delete_files: bool = True) -> None:
        self._delete(f"series/{series_id}", {"deleteFiles": str(delete_files).lower()})

    def delete_episode_file(self, file_id: int) -> None:
        self._delete(f"episodefile/{file_id}")

    # search / commands
    def command(self, body: JsonDict) -> JsonDict:
        return self._post("command", body)

    def series_search(self, series_id: int) -> JsonDict:
        return self.command({"name": "SeriesSearch", "seriesId": series_id})

    def season_search(self, series_id: int, season_number: int) -> JsonDict:
        return self.command(
            {
                "name": "SeasonSearch",
                "seriesId": series_id,
                "seasonNumber": season_number,
            }
        )

    def episode_search(self, episode_ids: list[int]) -> JsonDict:
        return self.command({"name": "EpisodeSearch", "episodeIds": episode_ids})

    def get_release(self, episode_id: int) -> JsonList:
        return cast("JsonList", self._get(f"release?episodeId={episode_id}"))

    def get_release_by_series(self, series_id: int) -> JsonList:
        return cast("JsonList", self._get(f"release?seriesId={series_id}"))

    def get_release_by_season(self, series_id: int, season_number: int) -> JsonList:
        return cast(
            "JsonList",
            self._get(f"release?seriesId={series_id}&seasonNumber={season_number}"),
        )

    def download_release(self, guid: str, indexer_id: int) -> JsonDict:
        return self._post("release", {"guid": guid, "indexerId": indexer_id})

    # lookup / add
    def lookup_series(self, term: str) -> JsonList:
        return cast(
            "JsonList",
            self._get(f"series/lookup?term={quote(term)}"),
        )

    def add_series(self, series: JsonDict) -> JsonDict:
        return self._post("series", series)

    def get_root_folders(self) -> JsonList:
        return cast("JsonList", self._get("rootfolder"))

    def get_quality_profiles(self) -> JsonList:
        return cast("JsonList", self._get("qualityprofile"))
