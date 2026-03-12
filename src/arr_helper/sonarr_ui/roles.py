"""Qt model roles used by Sonarr UI tree/dialogs."""

from PySide6.QtCore import Qt

ROLE_NODE_TYPE = Qt.UserRole + 1  # 'series' | 'season' | 'episode'
ROLE_SERIES_ID = Qt.UserRole + 2
ROLE_SERIES_PATH = Qt.UserRole + 3
ROLE_SEASON_NUM = Qt.UserRole + 4
ROLE_SEASON_PATH = Qt.UserRole + 5
ROLE_FILE_PATH = Qt.UserRole + 6
ROLE_FILE_ID = Qt.UserRole + 7
ROLE_EPISODE_DATA = Qt.UserRole + 8
ROLE_IS_MISSING = Qt.UserRole + 9

ROLE_RELEASE = Qt.UserRole + 20  # stores the release dict on each row
