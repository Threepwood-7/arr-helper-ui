"""Sonarr API client."""


import requests


class SonarrAPI:
    def __init__(
        self,
        url: str,
        api_key: str,
        http_user: str = '',
        http_pass: str = '',
        request_timeout: int | tuple = 300,
    ):
        self.url = url.rstrip('/')
        self.api_key = api_key
        self.headers = {'X-Api-Key': api_key}
        self.auth = (http_user, http_pass) if http_user else None
        self.timeout = request_timeout

    def _get(self, endpoint: str):
        r = requests.get(f"{self.url}/api/v3/{endpoint}", headers=self.headers, auth=self.auth, timeout=self.timeout)
        r.raise_for_status()
        return r.json()

    def _post(self, endpoint: str, data: dict):
        r = requests.post(f"{self.url}/api/v3/{endpoint}", headers=self.headers, auth=self.auth, json=data, timeout=self.timeout)
        r.raise_for_status()
        return r.json() if r.text else {}

    def _delete(self, endpoint: str, params: dict = None):
        r = requests.delete(f"{self.url}/api/v3/{endpoint}", headers=self.headers, auth=self.auth, params=params or {}, timeout=self.timeout)
        r.raise_for_status()
        return r

    def _put(self, endpoint: str, data: dict):
        r = requests.put(f"{self.url}/api/v3/{endpoint}", headers=self.headers, auth=self.auth, json=data, timeout=self.timeout)
        r.raise_for_status()
        return r.json()

    def get_series(self) -> list[dict]:
        return self._get('series')

    def get_series_by_id(self, series_id: int) -> dict:
        return self._get(f'series/{series_id}')

    def get_episodes(self, series_id: int) -> list[dict]:
        return self._get(f'episode?seriesId={series_id}')

    def get_episode_files(self, series_id: int) -> list[dict]:
        return self._get(f'episodefile?seriesId={series_id}')

    def update_series(self, series: dict) -> dict:
        return self._put(f'series/{series["id"]}', series)

    def update_episode(self, episode: dict) -> dict:
        return self._put(f'episode/{episode["id"]}', episode)

    def delete_series(self, series_id: int, delete_files: bool = True):
        self._delete(f'series/{series_id}', {'deleteFiles': str(delete_files).lower()})

    def delete_episode_file(self, file_id: int):
        self._delete(f'episodefile/{file_id}')

    # search / commands
    def command(self, body: dict) -> dict:
        return self._post('command', body)

    def series_search(self, series_id: int):
        return self.command({'name': 'SeriesSearch', 'seriesId': series_id})

    def season_search(self, series_id: int, season_number: int):
        return self.command({'name': 'SeasonSearch', 'seriesId': series_id, 'seasonNumber': season_number})

    def episode_search(self, episode_ids: list[int]):
        return self.command({'name': 'EpisodeSearch', 'episodeIds': episode_ids})

    def get_release(self, episode_id: int) -> list[dict]:
        return self._get(f'release?episodeId={episode_id}')

    def get_release_by_series(self, series_id: int) -> list[dict]:
        return self._get(f'release?seriesId={series_id}')

    def get_release_by_season(self, series_id: int, season_number: int) -> list[dict]:
        return self._get(f'release?seriesId={series_id}&seasonNumber={season_number}')

    def download_release(self, guid: str, indexer_id: int) -> dict:
        return self._post('release', {'guid': guid, 'indexerId': indexer_id})

    # lookup / add
    def lookup_series(self, term: str) -> list[dict]:
        return self._get(f'series/lookup?term={requests.utils.quote(term)}')

    def add_series(self, series: dict) -> dict:
        return self._post('series', series)

    def get_root_folders(self) -> list[dict]:
        return self._get('rootfolder')

    def get_quality_profiles(self) -> list[dict]:
        return self._get('qualityprofile')
