# Media Quality Checker & Sonarr UI Helper

A toolkit for managing Sonarr and Radarr libraries:
- `arr_helper.sonarr_ui.app`: desktop GUI for browsing/managing Sonarr content.
- `arr_helper.media_checker.app`: CLI checker for required English audio/subtitles.

## Table of Contents

- [Features](#features)
- [UI Walkthrough](#ui-walkthrough)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Configuration](#configuration)
- [Keyboard Shortcuts](#keyboard-shortcuts)
- [Menus](#menus)
- [Project Structure](#project-structure)
- [Architecture](#architecture)
- [Development](#development)
- [Troubleshooting](#troubleshooting)
- [Legal Disclaimer](#legal-disclaimer)

## Features

### Sonarr UI Helper (`arr_helper.sonarr_ui.app`)

PySide6 desktop app for browsing and managing Sonarr series, seasons, and episodes.

- Tree view with series/season/episode hierarchy
- Monitored status editing
- ffprobe metadata columns (resolution, bitrates, codecs, HDR, languages)
- Manual and auto search actions
- File deletion workflows (delete from disk, unmonitor + delete)
- Add new show with root folder and quality profile selection
- Open selected path in system file explorer

### Media Quality Checker (`arr_helper.media_checker.app`)

CLI tool that scans downloaded files in Sonarr/Radarr for required English audio/subtitle streams.

- `dry_run = true`: no destructive changes; reports only
- `dry_run = false`: files failing requirements are deleted and search commands are triggered
- `interactive = true`: lets you view/select alternative releases or skip

## UI Walkthrough

1. Configure sources and review your Sonarr library tree.

   ![Configure sources and review library](docs/images/ui-01-overview.png)

   Overview of series/seasons/episodes with monitored-state and metadata columns.

2. Execute add-show workflow with explicit root/profile choices.

   ![Execute add-show workflow](docs/images/ui-02-workflow.png)

   Add-show dialog state used to onboard new series into your managed library.

3. Review manual-search candidates and execute the selected action.

   ![Review manual search results](docs/images/ui-03-details.png)

   Manual-search dialog state for choosing and sending the best release to Sonarr.

## Requirements

- **Windows** (10 or later)
- **Python 3.10+**
- **ffmpeg/ffprobe** installed and available in PATH
- **Sonarr** and/or **Radarr** with API access

Python dependencies are defined in `pyproject.toml`.

## Installation

### First-time Setup

```bat
python scripts\windows\setup_env.py
```

Creates the `.venv` by running `uv sync --locked` (falls back to `uv sync` if no lockfile).

Manual alternative for development:

```bat
uv sync --all-extras
```

Fallback (pip):

```bat
python -m pip install -e .[dev]
```

Both installs register console scripts: `sonarr-ui-helper` and `media-quality-checker`.

Install ffmpeg if needed (Windows):

```bat
:: download from https://ffmpeg.org/download.html and add to PATH
:: or via package managers:
choco install ffmpeg
scoop install ffmpeg
winget install Gyan.FFmpeg
```

## Usage

### Sonarr UI Helper

#### Recommended (console-less)

```bat
pyw scripts\windows\run_app_gui.pyw
```

Launches the GUI without a console window. Auto-bootstraps the `.venv` via `setup_env.py` if not yet created.

#### With console

```bat
python scripts\windows\run_app.py
```

Runs via `hatch run python -m arr_helper`. Requires `hatch` in PATH.

#### Direct

```bat
python -m arr_helper
:: equivalent explicit GUI entrypoint:
python -m arr_helper.sonarr_ui.app
:: or installed script:
sonarr-ui-helper
```

#### Display Columns

- Name
- Size
- Mon
- Quality Profile
- Resolution
- V.Bitrate
- V.Codec
- HDR
- A.Codec
- A.Bitrate
- Audio Lang
- Sub Lang

#### Manual Search Dialog

Pressing **N** on an episode/season/series opens a Manual Search dialog that lists
available releases from indexers. The dialog supports:

- Text filter with autocomplete from previous searches
- Quality and Indexer dropdown filters
- Sortable columns: Title, Size (GB), Quality, Indexer, Age
- Double-click a release row to download it

### Media Quality Checker

```bat
:: media checker stays a manual CLI command:
python -m arr_helper.media_checker.app
:: or installed script:
media-quality-checker
```

#### Processing Flow

1. Fetch series/movies from enabled services
2. Inspect file streams via ffprobe
3. Check against configured language requirements
4. If failing:
   - non-interactive: delete file + trigger search
   - interactive: let user choose alternative, skip, or keep
5. Cache decisions/results for future runs

#### Safety Notes

- Run with `dry_run = true` first.
- In non-dry-run mode, this tool can delete files.
- Interactive mode can permanently remember skip decisions.

#### Example (Interactive)

```text
X Issue found: Some Show (2024)
  File: Some.Show.S01E01.1080p.BluRay.x264.mkv
  English audio: NO
  English subs: YES

View alternative releases? [Y/n]: y

1) Some.Show.2024.2160p.UHD.BluRay.REMUX-GRP   68.9 GB   Remux-2160p
2) Some.Show.2024.1080p.BluRay.REMUX-GRP       25.3 GB   Remux-1080p
3) Some.Show.2024.1080p.BluRay.x264-GRP        12.5 GB   Bluray-1080

Options: Enter release number, 's' to search, 'c' to clear, 0 to skip, -1 to keep
```

#### Automation (cron example)

```bat
:: Windows Task Scheduler equivalent of cron:
:: Run every day at 3 AM
schtasks /create /tn "MediaChecker" /tr "python -m arr_helper.media_checker.app" /sc daily /st 03:00
```

## Configuration

Both tools use a typed QSettings runtime config store.

- Config store backend: `QSettings(IniFormat, UserScope, "ThreepSoftwz", "arr_helper")`
- Default INI path: `%APPDATA%\ThreepSoftwz\arr_helper.ini`
- Default non-INI data path: `%LOCALAPPDATA%\ThreepSoftwz\arr_helper\`
- OV01 env overrides:
  - `CONFIG_DIR` overrides the QSettings INI root
  - `DATA_DIR` overrides the runtime data root (cache files, probe cache, etc.)
- Secrets are persisted in the same store (plaintext), with env overrides when present:
  - `ARR_HELPER_SONARR_API_KEY`
  - `ARR_HELPER_RADARR_API_KEY`
  - `ARR_HELPER_SONARR_HTTP_BASIC_AUTH_PASSWORD`
  - `ARR_HELPER_RADARR_HTTP_BASIC_AUTH_PASSWORD`
- Namespaces:
  - `config/...` service and checker settings
  - `ui/sonarr_ui/...` Sonarr UI state (column widths, dialog state, etc.)

1. Launch `python -m arr_helper` to open the setup wizard
2. Fill Sonarr/Radarr URLs and API keys
3. Save and restart the tool

Example:

```ini
config/sonarr/url=http://localhost:8989
config/sonarr/api_key=your-sonarr-api-key
config/sonarr/enabled=true
config/radarr/url=http://localhost:7878
config/radarr/api_key=your-radarr-api-key
config/radarr/enabled=true
config/settings/dry_run=false
config/settings/interactive=true
config/settings/require_english_audio=true
config/settings/require_english_subs=true
config/settings/english_language_codes=eng,en,english
config/settings/highlight_missing_subs=
```

### Finding API Keys

1. Open Sonarr/Radarr web UI
2. Go to `Settings -> General`
3. Find `API Key` in the Security section

### HTTP Basic Auth (Optional)

If your services are behind a reverse proxy with HTTP Basic Auth:

```ini
[sonarr]
http_basic_auth_username=myuser
http_basic_auth_password=mypass
```

### Cache Files

Cache/state files are stored under `%LOCALAPPDATA%\ThreepSoftwz\arr_helper\cache\` by default
(or `DATA_DIR\arr_helper\cache\` when overridden).

Files:
- `z_fprobe.cache` - ffprobe metadata cache used by `arr_helper.sonarr_ui.app`
- `z_user.cache` - persistent skip decisions used by `arr_helper.media_checker.app`
- `z_files.cache` - list of files already validated as good by `arr_helper.media_checker.app`

## Keyboard Shortcuts

| Shortcut | Action |
|---|---|
| **General** | |
| Ctrl+N | Add Show |
| F5 | Refresh |
| Ctrl+F5 | Clear cache and refresh |
| Ctrl+Q / Alt+X | Quit |
| **Navigation** | |
| Enter | Open file/folder in Explorer |
| O | Open in Explorer |
| Double-click | Open file/folder |
| Delete | Remove series/season from Sonarr (deletes files) |
| **View** | |
| Ctrl+E | Expand all (except Specials) |
| Ctrl+Shift+E | Expand series only |
| Ctrl+W | Collapse seasons |
| Ctrl+Shift+W | Collapse all |
| Ctrl+M | Toggle missing episodes |
| Ctrl+Shift+R | Reset saved view state |
| **Actions** | |
| M | Monitor selected |
| S | Auto search |
| N | Manual search |
| Q | Change quality profile (series) |
| U | Unmonitor selected |
| D | Delete from disk (keep in Sonarr) |
| Ctrl+Delete | Unmonitor and delete from disk |
| F1 | Shortcut help |

## Menus

- **File** — Add Show, Refresh, Clear Cache & Refresh, Quit
- **View** — Show Missing, Fit Columns, Reset View
- **Actions** — Monitor, Auto Search, Manual Search, Change Quality Profile, Unmonitor, Delete from Disk, Unmonitor & Delete, Open in Explorer
- **Tools** — Edit .ini file (opens the QSettings INI in your default editor)
- **Help** — Keyboard Shortcuts

## Project Structure

```text
arr-helper-ui/
|-- pyproject.toml
|-- uv.lock
|-- src/
|   `-- arr_helper/
|       |-- __init__.py
|       |-- core/
|       |   |-- ffprobe.py
|       |   |-- locking.py
|       |   `-- paths.py
|       |-- sonarr_ui/
|       |   |-- api.py
|       |   |-- app.py                   # Desktop GUI entrypoint
|       |   |-- helpers.py
|       |   |-- main_window.py
|       |   |-- probe_cache.py
|       |   |-- roles.py
|       |   |-- workers.py
|       |   `-- dialogs/
|       |       |-- add_show.py
|       |       `-- manual_search.py
|       `-- media_checker/
|           |-- app.py                   # CLI checker entrypoint
|           |-- checker.py
|           `-- config.py
|-- .pre-commit-config.yaml
|-- .github/workflows/ci.yml
|-- docs/
|   |-- development.md
|   |-- images/
|   |   |-- ui-01-overview.png
|   |   |-- ui-02-workflow.png
|   |   `-- ui-03-details.png
|   `-- release-checklist.md
|-- tests/
|   |-- conftest.py                      # Pytest fixtures (headless Qt setup)
|   |-- test_sonarr_ui_helper_unit.py
|   |-- test_sonarr_ui_helper_qt.py
|   `-- test_media_quality_checker_unit.py
|-- scripts/
|   `-- windows/
|       |-- setup_env.py                # Create/verify .venv via uv sync
|       |-- run_app.py                  # Launch app via hatch run
|       |-- run_app_gui.pyw              # Launch GUI without console window
|       `-- run_tests.py                # Run tests via hatch run test
```

## Architecture

- `src/` layout with two sub-packages: `sonarr_ui` (GUI) and `media_checker` (CLI).
- `core/` contains shared utilities (ffprobe integration, file locking, path resolution).
- GUI uses PySide6 with background workers (`QRunnable` + `QThreadPool`) for API calls and ffprobe scans.
- Media checker runs synchronously in CLI mode, with optional interactive prompts.
- Both tools share the same QSettings config store for Sonarr/Radarr credentials.

## Development

### Windows Helpers

| Script | Description |
|---|---|
| `python scripts\windows\setup_env.py` | Create/verify `.venv` via `uv sync --locked` |
| `pyw scripts\windows\run_app_gui.pyw` | Launch GUI without console window (auto-bootstraps venv) |
| `python scripts\windows\run_app.py` | Launch app via `hatch run` (requires hatch in PATH) |
| `python scripts\windows\run_tests.py` | Run test suite via `hatch run test` |

### Testing

```bat
uv run pytest
```

Tests run headless (`QT_QPA_PLATFORM=offscreen`) and cover API interaction,
cache logic, Qt window lifecycle, and the quality checker's decision logic.

### Quality Checks

```bat
uv run ruff check .
uv run mypy src
uv run python -m build
uv run pip-audit
```

### Pre-commit

```bat
uv run pre-commit install
uv run pre-commit run --all-files
```

## Troubleshooting

### ffprobe not found

On Windows the tools automatically search common install locations
(Chocolatey, Scoop, WinGet, `C:\ffmpeg`, etc.) even if ffprobe is not in PATH.

Verify ffprobe is available:

```bat
ffprobe -version
```

### API connection errors

- Verify URLs include `http://` or `https://`
- Verify API keys are valid
- Confirm Sonarr/Radarr are reachable

### Config issues

- Run the Sonarr UI setup wizard if required keys are missing
- Ensure at least one of `[sonarr].enabled` or `[radarr].enabled` is `true`

### Destructive operations warning

`arr_helper.media_checker.app` can delete and re-download files. Validate settings with `dry_run = true` before real runs.

---

<!-- legal-disclaimer:start -->
## Legal Disclaimer

THIS SOFTWARE IS PROVIDED "AS IS" AND "AS AVAILABLE," WITHOUT WARRANTIES OF ANY KIND, WHETHER EXPRESS, IMPLIED, STATUTORY, OR OTHERWISE, INCLUDING, WITHOUT LIMITATION, ANY IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, TITLE, NON-INFRINGEMENT, ACCURACY, OR QUIET ENJOYMENT. TO THE MAXIMUM EXTENT PERMITTED BY APPLICABLE LAW, THE AUTHORS, CONTRIBUTORS, MAINTAINERS, DISTRIBUTORS, AND AFFILIATED PARTIES SHALL NOT BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, CONSEQUENTIAL, EXEMPLARY, OR PUNITIVE DAMAGES, OR FOR ANY LOSS OF DATA, PROFITS, GOODWILL, BUSINESS OPPORTUNITY, OR SERVICE INTERRUPTION, ARISING OUT OF OR RELATING TO THE USE OF, OR INABILITY TO USE, THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. THIS SOFTWARE HAS BEEN DEVELOPED, IN WHOLE OR IN PART, BY "INTELLIGENT TOOLS"; ACCORDINGLY, OUTPUTS MAY CONTAIN ERRORS OR OMISSIONS, AND YOU ASSUME FULL RESPONSIBILITY FOR INDEPENDENT VALIDATION, TESTING, LEGAL COMPLIANCE, AND SAFE OPERATION PRIOR TO ANY RELIANCE OR DEPLOYMENT.
<!-- legal-disclaimer:end -->
