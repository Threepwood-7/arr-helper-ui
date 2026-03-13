"""Qt model roles used by Sonarr UI tree/dialogs."""

from PySide6.QtCore import Qt

_USER_ROLE = int(Qt.ItemDataRole.UserRole)

ROLE_NODE_TYPE: int = _USER_ROLE + 1  # 'series' | 'season' | 'episode'
ROLE_SERIES_ID: int = _USER_ROLE + 2
ROLE_SERIES_PATH: int = _USER_ROLE + 3
ROLE_SEASON_NUM: int = _USER_ROLE + 4
ROLE_SEASON_PATH: int = _USER_ROLE + 5
ROLE_FILE_PATH: int = _USER_ROLE + 6
ROLE_FILE_ID: int = _USER_ROLE + 7
ROLE_EPISODE_DATA: int = _USER_ROLE + 8
ROLE_IS_MISSING: int = _USER_ROLE + 9

ROLE_RELEASE: int = _USER_ROLE + 20  # stores the release dict on each row
