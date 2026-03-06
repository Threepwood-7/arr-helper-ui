from __future__ import annotations

import re
import subprocess
import sys
import tomllib
from fnmatch import fnmatch
from pathlib import Path

TEXT_SUFFIXES = {
    ".cfg",
    ".ini",
    ".jinja",
    ".json",
    ".md",
    ".ps1",
    ".py",
    ".pyw",
    ".pyi",
    ".toml",
    ".txt",
    ".yaml",
    ".yml",
}

ROOT_REQUIRED = (
    ".python-version",
    ".editorconfig",
    ".gitattributes",
    ".pre-commit-config.yaml",
    "pyproject.toml",
    "README.md",
)

WINDOWS_SCRIPT_REQUIRED = (
    "scripts/windows/setup_env.py",
    "scripts/windows/run_app.py",
    "scripts/windows/run_app_gui.pyw",
    "scripts/windows/run_tests.py",
)

PACKAGE_REQUIRED = ("__init__.py", "__main__.py", "py.typed")
SRC_MODULE_RE = re.compile(r"^[a-z_][a-z0-9_]*\.py$")
TEST_FILE_RE = re.compile(r"^test_[a-z0-9_]*\.py$")
PROJECT_NAME_RE = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")
PACKAGE_NAME_RE = re.compile(r"^[a-z][a-z0-9_]*$")
ALL_EXPORT_RE = re.compile(r"^\s*__all__\s*=", flags=re.MULTILINE)
REQUIRED_TEST_DIRS = ("tests/unit", "tests/integration", "tests/gui")
REQUIRED_PYTEST_MARKERS = {"integration", "gui", "slow"}
WINDOW_FIXTURE_RE = re.compile(
    r"@pytest\.fixture(?:\([^)]*\))?\s*[\r\n]+(?:[^\n]*[\r\n])*?def\s+window\s*\(",
    flags=re.MULTILINE,
)
MAX_SRC_FILE_LINES = 600
LEGACY_ROOT_CONFIG_PATTERNS = (
    "config.toml",
    "config_example.toml",
    "*_config.toml",
    "*_config_example.toml",
    "*_config_totemp.toml",
)
CANONICAL_CONFIG_REQUIRED = ("config/app.defaults.toml", "config/app.example.toml")
CANONICAL_CONFIG_LOCAL = "config/app.local.toml"
REQUIRED_NAMING_RULE = "N"
REQUIRED_QT_NAMING_IGNORES = {
    "activateWindow",
    "closeEvent",
    "eventFilter",
    "filterAcceptsRow",
    "headerData",
    "highlightBlock",
    "isActive",
    "isRunning",
    "keyPressEvent",
    "lessThan",
    "mouseReleaseEvent",
    "paintEvent",
    "requestInterruption",
    "setApplicationDisplayName",
    "setApplicationName",
    "setOrganizationName",
    "showEvent",
    "showNormal",
    "windowState",
}
LEGAL_DISCLAIMER_START = "<!-- legal-disclaimer:start -->"
LEGAL_DISCLAIMER_END = "<!-- legal-disclaimer:end -->"
LEGAL_DISCLAIMER_REQUIRED = """
## Legal Disclaimer

THIS SOFTWARE IS PROVIDED "AS IS" AND "AS AVAILABLE," WITHOUT WARRANTIES OF ANY KIND, WHETHER EXPRESS, IMPLIED, STATUTORY, OR OTHERWISE, INCLUDING, WITHOUT LIMITATION, ANY IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, TITLE, NON-INFRINGEMENT, ACCURACY, OR QUIET ENJOYMENT. TO THE MAXIMUM EXTENT PERMITTED BY APPLICABLE LAW, THE AUTHORS, CONTRIBUTORS, MAINTAINERS, DISTRIBUTORS, AND AFFILIATED PARTIES SHALL NOT BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, CONSEQUENTIAL, EXEMPLARY, OR PUNITIVE DAMAGES, OR FOR ANY LOSS OF DATA, PROFITS, GOODWILL, BUSINESS OPPORTUNITY, OR SERVICE INTERRUPTION, ARISING OUT OF OR RELATING TO THE USE OF, OR INABILITY TO USE, THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. THIS SOFTWARE HAS BEEN DEVELOPED, IN WHOLE OR IN PART, BY "INTELLIGENT TOOLS"; ACCORDINGLY, OUTPUTS MAY CONTAIN ERRORS OR OMISSIONS, AND YOU ASSUME FULL RESPONSIBILITY FOR INDEPENDENT VALIDATION, TESTING, LEGAL COMPLIANCE, AND SAFE OPERATION PRIOR TO ANY RELIANCE OR DEPLOYMENT.
"""


def run_command(repo_root: Path, args: list[str]) -> str:
    proc = subprocess.run(
        args,
        cwd=repo_root,
        check=False,
        capture_output=True,
        text=True,
    )
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or proc.stdout.strip() or "command failed")
    return proc.stdout


def tracked_files(repo_root: Path) -> list[str]:
    output = run_command(repo_root, ["git", "ls-files"])
    return [line.strip() for line in output.splitlines() if line.strip()]


def tracked_eol_lines(repo_root: Path) -> list[str]:
    output = run_command(repo_root, ["git", "ls-files", "--eol"])
    return [line.rstrip() for line in output.splitlines() if line.strip()]


def to_suffix(path_text: str) -> str:
    return Path(path_text).suffix.lower()


def is_legacy_root_config(path_text: str) -> bool:
    path = Path(path_text)
    if len(path.parts) != 1:
        return False
    name = path.name
    return any(fnmatch(name, pattern) for pattern in LEGACY_ROOT_CONFIG_PATTERNS)


def normalize_newlines(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n")


def count_text_lines(text: str) -> int:
    if not text:
        return 0
    return text.count("\n") + 1


def check_legal_disclaimer(readme_content: str, errors: list[str]) -> None:
    pattern = re.compile(
        rf"{re.escape(LEGAL_DISCLAIMER_START)}\s*(.*?)\s*{re.escape(LEGAL_DISCLAIMER_END)}",
        flags=re.DOTALL,
    )
    match = pattern.search(normalize_newlines(readme_content))
    if not match:
        errors.append(
            "README.md must include legal disclaimer markers "
            f"{LEGAL_DISCLAIMER_START} ... {LEGAL_DISCLAIMER_END}."
        )
        return

    actual = normalize_newlines(match.group(1)).strip()
    expected = normalize_newlines(LEGAL_DISCLAIMER_REQUIRED).strip()
    if actual != expected:
        errors.append(
            "README.md legal disclaimer content does not match the canonical required block."
        )


def load_pyproject(repo_root: Path, errors: list[str]) -> dict[str, object]:
    pyproject_path = repo_root / "pyproject.toml"
    if not pyproject_path.exists():
        return {}
    try:
        data = tomllib.loads(pyproject_path.read_text(encoding="utf-8"))
    except (OSError, tomllib.TOMLDecodeError) as exc:
        errors.append(f"invalid pyproject.toml ({exc})")
        return {}
    if not isinstance(data, dict):
        errors.append("pyproject.toml must be a table")
        return {}
    return data


def require_table(
    table: dict[str, object],
    key: str,
    errors: list[str],
    label: str,
) -> dict[str, object]:
    value = table.get(key)
    if not isinstance(value, dict):
        errors.append(f"Missing or invalid [{label}] table in pyproject.toml")
        return {}
    return value


def require_string_list(
    table: dict[str, object],
    key: str,
    errors: list[str],
    label: str,
) -> list[str]:
    value = table.get(key)
    if not isinstance(value, list) or not all(isinstance(item, str) for item in value):
        errors.append(f"Missing or invalid [{label}.{key}] list in pyproject.toml")
        return []
    return [item for item in value if isinstance(item, str)]


def parse_marker_name(marker_entry: str) -> str:
    return marker_entry.split(":", 1)[0].strip()


def main() -> int:
    repo_root = Path(__file__).resolve().parents[2]
    errors: list[str] = []
    warnings: list[str] = []

    try:
        tracked = tracked_files(repo_root)
        eol_lines = tracked_eol_lines(repo_root)
    except RuntimeError as exc:
        print(f"Policy check failed: unable to inspect git-tracked files ({exc}).", file=sys.stderr)
        return 1

    tracked_set = set(tracked)

    src_root = repo_root / "src"
    src_packages = []
    if src_root.exists():
        src_packages = sorted(
            [
                p
                for p in src_root.iterdir()
                if p.is_dir() and not p.name.startswith(".") and p.name != "__pycache__"
            ]
        )

    if len(src_packages) != 1:
        errors.append(f"Expected exactly one package directory under src/, found {len(src_packages)}.")
        package_name = ""
    else:
        package_name = src_packages[0].name
        if not PACKAGE_NAME_RE.fullmatch(package_name):
            errors.append(
                f"Package directory name must be snake_case under src/: {package_name}"
            )

    pyproject = load_pyproject(repo_root, errors)
    if pyproject:
        project_table = require_table(pyproject, "project", errors, "project")
        project_name = project_table.get("name")
        if not isinstance(project_name, str) or not project_name.strip():
            errors.append("project.name must be a non-empty string in pyproject.toml")
        elif not PROJECT_NAME_RE.fullmatch(project_name):
            errors.append(f"project.name must be kebab-case: {project_name}")

        tool_table = require_table(pyproject, "tool", errors, "tool")
        basedpyright_table = require_table(tool_table, "basedpyright", errors, "tool.basedpyright")
        if basedpyright_table and basedpyright_table.get("typeCheckingMode") != "strict":
            errors.append("tool.basedpyright.typeCheckingMode must be strict")

        ruff_table = require_table(tool_table, "ruff", errors, "tool.ruff")
        if ruff_table:
            lint_table = require_table(ruff_table, "lint", errors, "tool.ruff.lint")
            if lint_table:
                select_rules = require_string_list(lint_table, "select", errors, "tool.ruff.lint")
                if select_rules and REQUIRED_NAMING_RULE not in set(select_rules):
                    errors.append("tool.ruff.lint.select must include naming rule 'N'")

                pep8_naming_table = require_table(
                    lint_table,
                    "pep8-naming",
                    errors,
                    "tool.ruff.lint.pep8-naming",
                )
                if pep8_naming_table:
                    ignore_names = require_string_list(
                        pep8_naming_table,
                        "ignore-names",
                        errors,
                        "tool.ruff.lint.pep8-naming",
                    )
                    if ignore_names:
                        missing_ignores = sorted(REQUIRED_QT_NAMING_IGNORES - set(ignore_names))
                        if missing_ignores:
                            errors.append(
                                "tool.ruff.lint.pep8-naming.ignore-names is missing required Qt "
                                f"exceptions: {', '.join(missing_ignores)}"
                            )

        pytest_table = tool_table.get("pytest", {}).get("ini_options", {})
        if not isinstance(pytest_table, dict):
            errors.append("Missing or invalid [tool.pytest.ini_options] table in pyproject.toml")
        else:
            testpaths = pytest_table.get("testpaths")
            if not isinstance(testpaths, list) or "tests" not in testpaths:
                errors.append("tool.pytest.ini_options.testpaths must include 'tests'")

            markers = pytest_table.get("markers")
            if not isinstance(markers, list) or not all(isinstance(item, str) for item in markers):
                errors.append(
                    "tool.pytest.ini_options.markers must be a list containing integration/gui/slow"
                )
            else:
                configured_markers = {parse_marker_name(item) for item in markers}
                missing_markers = sorted(REQUIRED_PYTEST_MARKERS - configured_markers)
                if missing_markers:
                    errors.append(
                        "tool.pytest.ini_options.markers is missing required markers: "
                        + ", ".join(missing_markers)
                    )

    for rel in ROOT_REQUIRED:
        if rel not in tracked_set:
            errors.append(f"Missing required tracked file: {rel}")

    readme_path = repo_root / "README.md"
    if readme_path.exists():
        try:
            readme_content = readme_path.read_text(encoding="utf-8", errors="ignore")
        except OSError as exc:
            errors.append(f"Unable to read README.md for legal disclaimer policy: {exc}")
        else:
            check_legal_disclaimer(readme_content, errors)

    for rel in WINDOWS_SCRIPT_REQUIRED:
        if rel not in tracked_set:
            errors.append(f"Missing required tracked file: {rel}")

    for rel in REQUIRED_TEST_DIRS:
        if rel not in tracked_set and not (repo_root / rel).is_dir():
            errors.append(f"Missing required test directory: {rel}")

    tests_root = repo_root / "tests"
    if tests_root.is_dir():
        for test_dir in tests_root.rglob("*"):
            if not test_dir.is_dir():
                continue
            if test_dir.name.startswith(".") or test_dir.name == "__pycache__":
                continue
            rel_dir = test_dir.relative_to(repo_root).as_posix()
            init_file = test_dir / "__init__.py"
            if not init_file.is_file():
                errors.append(f"Missing __init__.py in test package directory: {rel_dir}")
    else:
        errors.append("Missing tests/ directory")

    conftest_path = repo_root / "tests" / "conftest.py"
    if not conftest_path.is_file():
        errors.append("Missing tests/conftest.py")
    else:
        try:
            conftest_content = conftest_path.read_text(encoding="utf-8", errors="ignore")
        except OSError as exc:
            errors.append(f"Unable to read tests/conftest.py for fixture policy: {exc}")
        else:
            if not WINDOW_FIXTURE_RE.search(conftest_content):
                errors.append("tests/conftest.py must define a shared @pytest.fixture def window(...)")

    if package_name:
        for rel in PACKAGE_REQUIRED:
            package_file = f"src/{package_name}/{rel}"
            if package_file not in tracked_set:
                errors.append(f"Missing required tracked file: {package_file}")
        package_init_file = f"src/{package_name}/__init__.py"
        if package_init_file in tracked_set:
            try:
                package_init_content = (repo_root / package_init_file).read_text(
                    encoding="utf-8",
                    errors="ignore",
                )
            except OSError as exc:
                errors.append(f"Unable to read {package_init_file} for __all__ policy: {exc}")
            else:
                if not ALL_EXPORT_RE.search(package_init_content):
                    errors.append(f"Missing required __all__ export list in {package_init_file}")

    for rel in tracked:
        lower = rel.lower()
        if lower.endswith(".cmd") or lower.endswith(".ps1"):
            errors.append(f"Legacy .cmd/.ps1 scripts are not allowed: {rel}")

    legacy_root_configs = sorted(rel for rel in tracked if is_legacy_root_config(rel))
    for rel in legacy_root_configs:
        errors.append(f"Legacy root config filename is not allowed: {rel}")

    if CANONICAL_CONFIG_LOCAL in tracked_set:
        errors.append(
            f"Local override config must be untracked: {CANONICAL_CONFIG_LOCAL}"
        )

    has_toml_app_config = bool(
        legacy_root_configs
        or any(rel in tracked_set for rel in CANONICAL_CONFIG_REQUIRED)
        or CANONICAL_CONFIG_LOCAL in tracked_set
    )
    if has_toml_app_config:
        for rel in CANONICAL_CONFIG_REQUIRED:
            if rel not in tracked_set:
                errors.append(f"Missing canonical app config file: {rel}")

    for rel in tracked:
        if not rel.endswith(".py"):
            continue

        if package_name and rel.startswith(f"src/{package_name}/"):
            filename = Path(rel).name
            if filename in PACKAGE_REQUIRED:
                continue
            if not SRC_MODULE_RE.match(filename):
                errors.append(f"Non-standard Python module filename in src/: {rel}")

        if rel.startswith("tests/"):
            filename = Path(rel).name
            if filename in {"__init__.py", "conftest.py"}:
                continue
            if not TEST_FILE_RE.match(filename):
                errors.append(f"Non-standard test filename in tests/: {rel}")

        if rel.startswith("src/"):
            abs_path = repo_root / rel
            try:
                line_count = count_text_lines(
                    abs_path.read_text(encoding="utf-8", errors="ignore")
                )
            except OSError:
                continue
            if line_count > MAX_SRC_FILE_LINES:
                warnings.append(
                    f"Python source file exceeds {MAX_SRC_FILE_LINES} lines "
                    f"({line_count}): {rel}"
                )

    if package_name:
        banned_launch_ref = f"python -m {package_name}.main"
        for rel in tracked:
            if to_suffix(rel) not in TEXT_SUFFIXES:
                continue
            abs_path = repo_root / rel
            try:
                content = abs_path.read_text(encoding="utf-8", errors="ignore")
            except OSError:
                continue
            if banned_launch_ref in content:
                errors.append(f"Found deprecated launch reference '{banned_launch_ref}' in {rel}")

    gitattributes_path = repo_root / ".gitattributes"
    if gitattributes_path.exists():
        content = gitattributes_path.read_text(encoding="utf-8", errors="ignore")
        if not re.search(r"^\*\s+text=auto\s+eol=lf\s*$", content, flags=re.MULTILINE):
            errors.append("Missing canonical LF rule in .gitattributes: '* text=auto eol=lf'")
    else:
        errors.append("Missing .gitattributes")

    for line in eol_lines:
        meta, _, rel = line.partition("\t")
        if not rel:
            continue
        if to_suffix(rel) not in TEXT_SUFFIXES:
            continue
        if "w/crlf" in meta:
            errors.append(f"CRLF line ending detected in tracked text file: {rel}")

    if errors:
        print("Project policy check failed with the following issues:", file=sys.stderr)
        for issue in errors:
            print(f"- {issue}", file=sys.stderr)
        if warnings:
            print("Project policy check warnings:", file=sys.stderr)
            for issue in warnings:
                print(f"- {issue}", file=sys.stderr)
        return 1

    if warnings:
        print("Project policy check warnings:")
        for issue in warnings:
            print(f"- {issue}")
        print("Project policy check passed with warnings.")
        return 0

    print("Project policy check passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
