from __future__ import annotations

import ast
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
    "uv.lock",
    "README.md",
)

QT_APP_WINDOWS_SCRIPT_REQUIRED = (
    "scripts/windows/setup_env.py",
    "scripts/windows/run_app.py",
    "scripts/windows/run_app_gui.pyw",
    "scripts/windows/run_tests.py",
)
SHARED_LIB_WINDOWS_SCRIPT_REQUIRED = (
    "scripts/windows/setup_env.py",
    "scripts/windows/run_tests.py",
)

QT_APP_PACKAGE_REQUIRED = (
    "__init__.py",
    "__main__.py",
    "constants.py",
    "py.typed",
)
SHARED_LIB_PACKAGE_REQUIRED = (
    "__init__.py",
    "__main__.py",
    "py.typed",
)
SRC_MODULE_RE = re.compile(r"^[a-z_][a-z0-9_]*\.py$")
TEST_FILE_RE = re.compile(r"^test_[a-z0-9_]*\.py$")
PROJECT_NAME_RE = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")
PACKAGE_NAME_RE = re.compile(r"^[a-z][a-z0-9_]*$")
ALL_EXPORT_RE = re.compile(r"^\s*__all__\s*=", flags=re.MULTILINE)
QT_APP_REQUIRED_TEST_DIRS = ("tests/unit", "tests/integration", "tests/gui")
SHARED_LIB_REQUIRED_TEST_DIRS = ("tests/unit", "tests/integration")
QT_APP_REQUIRED_PYTEST_MARKERS = {"integration", "gui", "slow"}
SHARED_LIB_REQUIRED_PYTEST_MARKERS = {"integration", "slow"}
WINDOW_FIXTURE_RE = re.compile(
    r"@pytest\.fixture(?:\([^)]*\))?\s*[\r\n]+(?:[^\n]*[\r\n])*?def\s+window\s*\(",
    flags=re.MULTILINE,
)
PROJECT_KIND_RE = re.compile(
    r"^\s*project_kind:\s*['\"]?([a-z_]+)['\"]?\s*$",
    flags=re.MULTILINE,
)
MAX_SRC_FILE_LINES = 600
MAX_SRC_CLASS_LINES_WARNING = 450
MAX_SRC_FUNCTION_LINES_WARNING = 120
MAX_SIZE_GUIDANCE_WARNINGS = 30
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
ALLOWED_NON_PACKAGE_SRC_DIRS = {"c", "vbs"}
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
SILENT_BROAD_EXCEPT_RE = re.compile(
    r"except\s+Exception(?:\s+as\s+[A-Za-z_][A-Za-z0-9_]*)?\s*:\s*(?:#.*\n\s*)?pass\b",
    flags=re.MULTILINE,
)
MAX_DOCSTRING_WARNINGS = 20
LEGAL_DISCLAIMER_REQUIRED = (
    "## Legal Disclaimer\n\n"
    'THIS SOFTWARE IS PROVIDED "AS IS" AND "AS AVAILABLE," WITHOUT WARRANTIES OF ANY '
    "KIND, WHETHER EXPRESS, IMPLIED, STATUTORY, OR OTHERWISE, INCLUDING, WITHOUT "
    "LIMITATION, ANY IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR "
    "PURPOSE, TITLE, NON-INFRINGEMENT, ACCURACY, OR QUIET ENJOYMENT. TO THE MAXIMUM "
    "EXTENT PERMITTED BY APPLICABLE LAW, THE AUTHORS, CONTRIBUTORS, MAINTAINERS, "
    "DISTRIBUTORS, AND AFFILIATED PARTIES SHALL NOT BE LIABLE FOR ANY DIRECT, "
    "INDIRECT, INCIDENTAL, SPECIAL, CONSEQUENTIAL, EXEMPLARY, OR PUNITIVE DAMAGES, OR "
    "FOR ANY LOSS OF DATA, PROFITS, GOODWILL, BUSINESS OPPORTUNITY, OR SERVICE "
    "INTERRUPTION, ARISING OUT OF OR RELATING TO THE USE OF, OR INABILITY TO USE, THIS "
    "SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. THIS SOFTWARE HAS "
    'BEEN DEVELOPED, IN WHOLE OR IN PART, BY "INTELLIGENT TOOLS"; ACCORDINGLY, OUTPUTS '
    "MAY CONTAIN ERRORS OR OMISSIONS, AND YOU ASSUME FULL RESPONSIBILITY FOR "
    "INDEPENDENT VALIDATION, TESTING, LEGAL COMPLIANCE, AND SAFE OPERATION PRIOR TO "
    "ANY RELIANCE OR DEPLOYMENT.\n"
)


def run_command(repo_root: Path, args: list[str]) -> str:
    proc = subprocess.run(
        args,
        cwd=repo_root,
        check=False,
        capture_output=True,
        text=True,
    )
    if proc.returncode != 0:
        raise RuntimeError(
            proc.stderr.strip() or proc.stdout.strip() or "command failed"
        )
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


def has_utf8_bom(path: Path) -> bool:
    try:
        with path.open("rb") as handle:
            return handle.read(3) == b"\xef\xbb\xbf"
    except OSError:
        return False


def has_module_package_name_collisions(tracked: list[str], package_name: str) -> bool:
    prefix = f"src/{package_name}/"
    module_names: set[str] = set()
    package_names: set[str] = set()
    for rel in tracked:
        if not rel.startswith(prefix):
            continue
        parts = Path(rel).parts
        if len(parts) < 4:
            continue
        tail = parts[3:]
        if len(tail) == 1 and tail[0].endswith(".py"):
            stem = Path(tail[0]).stem
            if stem not in {"__init__", "__main__"}:
                module_names.add(stem)
        if len(tail) >= 2 and tail[1] == "__init__.py":
            package_names.add(tail[0])
    return bool(module_names & package_names)


def validate_main_entrypoint_contract(
    repo_root: Path,
    package_name: str,
    tracked_set: set[str],
    errors: list[str],
) -> None:
    main_rel = f"src/{package_name}/__main__.py"
    if main_rel not in tracked_set:
        return
    main_path = repo_root / main_rel
    try:
        content = main_path.read_text(encoding="utf-8", errors="ignore")
    except OSError as exc:
        errors.append(f"Unable to read {main_rel} for entrypoint contract: {exc}")
        return

    if "from __future__ import annotations" not in content:
        errors.append(f"{main_rel} must include 'from __future__ import annotations'.")

    if not re.search(r"if __name__\s*==\s*['\"]__main__['\"]\s*:", content):
        errors.append(
            f'{main_rel} must guard execution with if __name__ == "__main__":'
        )

    if not re.search(r"raise\s+SystemExit\(\s*main\(\)\s*\)", content):
        errors.append(
            f"{main_rel} must end with 'raise SystemExit(main())' in the guard block."
        )


def check_legal_disclaimer(readme_content: str, errors: list[str]) -> None:
    pattern = re.compile(
        rf"{re.escape(LEGAL_DISCLAIMER_START)}\s*(.*?)\s*"
        rf"{re.escape(LEGAL_DISCLAIMER_END)}",
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
            "README.md legal disclaimer content does not match the canonical "
            "required block."
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


def load_project_kind(repo_root: Path) -> str:
    answers_path = repo_root / ".copier-answers.yml"
    if not answers_path.exists():
        return "qt_app"
    try:
        content = answers_path.read_text(encoding="utf-8", errors="ignore")
    except OSError:
        return "qt_app"
    match = PROJECT_KIND_RE.search(content)
    if not match:
        return "qt_app"
    project_kind = match.group(1).strip()
    if project_kind not in {"qt_app", "shared_lib"}:
        return "qt_app"
    return project_kind


def required_windows_scripts(project_kind: str) -> tuple[str, ...]:
    if project_kind == "shared_lib":
        return SHARED_LIB_WINDOWS_SCRIPT_REQUIRED
    return QT_APP_WINDOWS_SCRIPT_REQUIRED


def required_package_files(project_kind: str) -> tuple[str, ...]:
    if project_kind == "shared_lib":
        return SHARED_LIB_PACKAGE_REQUIRED
    return QT_APP_PACKAGE_REQUIRED


def required_test_dirs(project_kind: str) -> tuple[str, ...]:
    if project_kind == "shared_lib":
        return SHARED_LIB_REQUIRED_TEST_DIRS
    return QT_APP_REQUIRED_TEST_DIRS


def required_pytest_markers(project_kind: str) -> set[str]:
    if project_kind == "shared_lib":
        return SHARED_LIB_REQUIRED_PYTEST_MARKERS
    return QT_APP_REQUIRED_PYTEST_MARKERS


def collect_silent_broad_exception_warnings(
    repo_root: Path,
    tracked: list[str],
    warnings: list[str],
) -> None:
    for rel in tracked:
        if not rel.startswith("src/") or not rel.endswith(".py"):
            continue
        abs_path = repo_root / rel
        try:
            content = abs_path.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue
        if SILENT_BROAD_EXCEPT_RE.search(content):
            warnings.append(
                "Silent broad exception guidance: replace 'except Exception: "
                f"pass' in {rel} "
                "with logging or signal emission."
            )


def _is_public_name(name: str) -> bool:
    return not name.startswith("_")


def collect_docstring_guidance(
    repo_root: Path,
    tracked: list[str],
    warnings: list[str],
) -> None:
    warning_count = 0
    truncated = False

    for rel in tracked:
        if not rel.startswith("src/") or not rel.endswith(".py"):
            continue

        abs_path = repo_root / rel
        try:
            content = abs_path.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue

        try:
            module = ast.parse(content)
        except SyntaxError:
            continue

        if ast.get_docstring(module) is None:
            warnings.append(f"Docstring guidance: add a module docstring in {rel}.")
            warning_count += 1

        for node in module.body:
            if isinstance(node, ast.ClassDef):
                if _is_public_name(node.name) and ast.get_docstring(node) is None:
                    warnings.append(
                        "Docstring guidance: add class docstring for "
                        f"'{node.name}' in {rel}:{node.lineno}."
                    )
                    warning_count += 1
            elif (
                isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
                and _is_public_name(node.name)
                and ast.get_docstring(node) is None
            ):
                warnings.append(
                    "Docstring guidance: add function docstring for "
                    f"'{node.name}' in {rel}:{node.lineno}."
                )
                warning_count += 1

            if warning_count >= MAX_DOCSTRING_WARNINGS:
                truncated = True
                break

        if truncated:
            break

    if truncated:
        warnings.append(
            f"Docstring guidance output truncated at {MAX_DOCSTRING_WARNINGS} warnings."
        )


def _node_line_count(node: ast.AST) -> int:
    lineno = int(getattr(node, "lineno", 1))
    end_lineno = int(getattr(node, "end_lineno", lineno))
    return max(1, end_lineno - lineno + 1)


def _append_large_class_guidance(
    module: ast.Module,
    *,
    rel: str,
    warnings: list[str],
    warning_count: int,
) -> tuple[int, bool]:
    for node in module.body:
        if not isinstance(node, ast.ClassDef):
            continue
        class_lines = _node_line_count(node)
        if class_lines > MAX_SRC_CLASS_LINES_WARNING:
            warnings.append(
                "Structure guidance: oversized class "
                f"'{node.name}' has {class_lines} lines in {rel}:{node.lineno}. "
                "Prefer feature coordinators or helper components."
            )
            warning_count += 1
        if warning_count >= MAX_SIZE_GUIDANCE_WARNINGS:
            return warning_count, True
    return warning_count, False


def _append_large_function_guidance(
    module: ast.Module,
    *,
    rel: str,
    warnings: list[str],
    warning_count: int,
) -> tuple[int, bool]:
    for node in ast.walk(module):
        if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        function_lines = _node_line_count(node)
        if function_lines > MAX_SRC_FUNCTION_LINES_WARNING:
            warnings.append(
                "Structure guidance: oversized function "
                f"'{node.name}' has {function_lines} lines in {rel}:{node.lineno}. "
                "Prefer extracting focused helper routines."
            )
            warning_count += 1
        if warning_count >= MAX_SIZE_GUIDANCE_WARNINGS:
            return warning_count, True
    return warning_count, False


def collect_structure_size_guidance(
    repo_root: Path,
    tracked: list[str],
    warnings: list[str],
) -> None:
    warning_count = 0
    truncated = False

    for rel in tracked:
        if not rel.startswith("src/") or not rel.endswith(".py"):
            continue

        abs_path = repo_root / rel
        try:
            content = abs_path.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue

        try:
            module = ast.parse(content)
        except SyntaxError:
            continue

        warning_count, truncated = _append_large_class_guidance(
            module,
            rel=rel,
            warnings=warnings,
            warning_count=warning_count,
        )
        if truncated:
            break

        warning_count, truncated = _append_large_function_guidance(
            module,
            rel=rel,
            warnings=warnings,
            warning_count=warning_count,
        )
        if truncated:
            break

    if truncated:
        warnings.append(
            "Structure guidance output truncated at "
            f"{MAX_SIZE_GUIDANCE_WARNINGS} warnings."
        )


def collect_src_packages(src_root: Path) -> list[Path]:
    if not src_root.exists():
        return []
    return sorted(
        [
            path
            for path in src_root.iterdir()
            if path.is_dir()
            and not path.name.startswith(".")
            and path.name != "__pycache__"
            and path.name not in ALLOWED_NON_PACKAGE_SRC_DIRS
        ]
    )


def resolve_package_name(src_packages: list[Path], errors: list[str]) -> str:
    if len(src_packages) != 1:
        errors.append(
            "Expected exactly one package directory under src/, found "
            f"{len(src_packages)}."
        )
        return ""
    package_name = src_packages[0].name
    if not PACKAGE_NAME_RE.fullmatch(package_name):
        errors.append(
            f"Package directory name must be snake_case under src/: {package_name}"
        )
    return package_name


def validate_pyproject_policy(
    pyproject: dict[str, object],
    *,
    project_kind: str,
    errors: list[str],
) -> None:
    project_table = require_table(pyproject, "project", errors, "project")
    project_name = project_table.get("name")
    if not isinstance(project_name, str) or not project_name.strip():
        errors.append("project.name must be a non-empty string in pyproject.toml")
    elif not PROJECT_NAME_RE.fullmatch(project_name):
        errors.append(f"project.name must be kebab-case: {project_name}")

    tool_table = require_table(pyproject, "tool", errors, "tool")
    basedpyright_table = require_table(
        tool_table,
        "basedpyright",
        errors,
        "tool.basedpyright",
    )
    if basedpyright_table and basedpyright_table.get("typeCheckingMode") != "strict":
        errors.append("tool.basedpyright.typeCheckingMode must be strict")

    ruff_table = require_table(tool_table, "ruff", errors, "tool.ruff")
    if ruff_table:
        lint_table = require_table(ruff_table, "lint", errors, "tool.ruff.lint")
        if lint_table:
            select_rules = require_string_list(
                lint_table,
                "select",
                errors,
                "tool.ruff.lint",
            )
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
                    missing_ignores = sorted(
                        REQUIRED_QT_NAMING_IGNORES - set(ignore_names)
                    )
                    if missing_ignores:
                        errors.append(
                            "tool.ruff.lint.pep8-naming.ignore-names is "
                            "missing required Qt exceptions: "
                            f"{', '.join(missing_ignores)}"
                        )

    pytest_table = tool_table.get("pytest", {}).get("ini_options", {})
    if not isinstance(pytest_table, dict):
        errors.append(
            "Missing or invalid [tool.pytest.ini_options] table in pyproject.toml"
        )
        return

    testpaths = pytest_table.get("testpaths")
    if not isinstance(testpaths, list) or "tests" not in testpaths:
        errors.append("tool.pytest.ini_options.testpaths must include 'tests'")

    markers = pytest_table.get("markers")
    if not isinstance(markers, list) or not all(
        isinstance(item, str) for item in markers
    ):
        required_marker_names = sorted(required_pytest_markers(project_kind))
        errors.append(
            "tool.pytest.ini_options.markers must be a list containing "
            + "/".join(required_marker_names)
        )
        return

    configured_markers = {parse_marker_name(item) for item in markers}
    missing_markers = sorted(required_pytest_markers(project_kind) - configured_markers)
    if missing_markers:
        errors.append(
            "tool.pytest.ini_options.markers is missing required markers: "
            + ", ".join(missing_markers)
        )


def validate_required_paths(
    repo_root: Path,
    *,
    tracked_set: set[str],
    project_kind: str,
    errors: list[str],
) -> None:
    for rel in ROOT_REQUIRED:
        if rel not in tracked_set:
            errors.append(f"Missing required tracked file: {rel}")

    readme_path = repo_root / "README.md"
    if readme_path.exists():
        try:
            readme_content = readme_path.read_text(encoding="utf-8", errors="ignore")
        except OSError as exc:
            errors.append(
                f"Unable to read README.md for legal disclaimer policy: {exc}"
            )
        else:
            check_legal_disclaimer(readme_content, errors)

    for rel in required_windows_scripts(project_kind):
        if rel not in tracked_set:
            errors.append(f"Missing required tracked file: {rel}")

    for rel in required_test_dirs(project_kind):
        if rel not in tracked_set and not (repo_root / rel).is_dir():
            errors.append(f"Missing required test directory: {rel}")


def validate_test_layout(
    repo_root: Path,
    *,
    project_kind: str,
    errors: list[str],
) -> None:
    tests_root = repo_root / "tests"
    if not tests_root.is_dir():
        errors.append("Missing tests/ directory")
        return

    for test_dir in tests_root.rglob("*"):
        if not test_dir.is_dir():
            continue
        if test_dir.name.startswith(".") or test_dir.name == "__pycache__":
            continue
        rel_dir = test_dir.relative_to(repo_root).as_posix()
        if not (test_dir / "__init__.py").is_file():
            errors.append(f"Missing __init__.py in test package directory: {rel_dir}")

    if project_kind != "qt_app":
        return

    conftest_path = tests_root / "conftest.py"
    if not conftest_path.is_file():
        errors.append("Missing tests/conftest.py")
        return
    try:
        conftest_content = conftest_path.read_text(
            encoding="utf-8",
            errors="ignore",
        )
    except OSError as exc:
        errors.append(f"Unable to read tests/conftest.py for fixture policy: {exc}")
        return
    if not WINDOW_FIXTURE_RE.search(conftest_content):
        errors.append(
            "tests/conftest.py must define a shared @pytest.fixture def window(...)"
        )


def validate_package_contracts(
    repo_root: Path,
    *,
    tracked: list[str],
    tracked_set: set[str],
    package_name: str,
    project_kind: str,
    errors: list[str],
) -> None:
    if not package_name:
        return

    for rel in required_package_files(project_kind):
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
            errors.append(
                f"Unable to read {package_init_file} for __all__ policy: {exc}"
            )
        else:
            if not ALL_EXPORT_RE.search(package_init_content):
                errors.append(
                    f"Missing required __all__ export list in {package_init_file}"
                )

    validate_main_entrypoint_contract(repo_root, package_name, tracked_set, errors)
    if has_module_package_name_collisions(tracked, package_name):
        errors.append(
            "Module/package collision detected in "
            f"src/{package_name}: avoid name pairs like "
            "'utils.py' and 'utils/__init__.py'."
        )


def validate_legacy_configs(
    tracked: list[str],
    *,
    tracked_set: set[str],
    errors: list[str],
) -> None:
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
    if not has_toml_app_config:
        return
    for rel in CANONICAL_CONFIG_REQUIRED:
        if rel not in tracked_set:
            errors.append(f"Missing canonical app config file: {rel}")


def validate_python_paths(
    repo_root: Path,
    *,
    tracked: list[str],
    package_name: str,
    project_kind: str,
    warnings: list[str],
    errors: list[str],
) -> None:
    required_files = set(required_package_files(project_kind))
    package_prefix = f"src/{package_name}/" if package_name else ""

    for rel in tracked:
        if not rel.endswith(".py"):
            continue

        if package_prefix and rel.startswith(package_prefix):
            filename = Path(rel).name
            if filename not in required_files and not SRC_MODULE_RE.match(filename):
                errors.append(f"Non-standard Python module filename in src/: {rel}")

        if rel.startswith("tests/"):
            filename = Path(rel).name
            if filename not in {
                "__init__.py",
                "conftest.py",
            } and not TEST_FILE_RE.match(filename):
                errors.append(f"Non-standard test filename in tests/: {rel}")

        if not rel.startswith("src/"):
            continue
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


def validate_launch_references(
    repo_root: Path,
    *,
    tracked: list[str],
    package_name: str,
    errors: list[str],
) -> None:
    if not package_name:
        return

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
            errors.append(
                f"Found deprecated launch reference '{banned_launch_ref}' in {rel}"
            )


def validate_text_file_policies(
    repo_root: Path,
    *,
    tracked: list[str],
    eol_lines: list[str],
    errors: list[str],
) -> None:
    gitattributes_path = repo_root / ".gitattributes"
    if gitattributes_path.exists():
        content = gitattributes_path.read_text(encoding="utf-8", errors="ignore")
        if not re.search(r"^\*\s+text=auto\s+eol=lf\s*$", content, flags=re.MULTILINE):
            errors.append(
                "Missing canonical LF rule in .gitattributes: '* text=auto eol=lf'"
            )
    else:
        errors.append("Missing .gitattributes")

    for line in eol_lines:
        meta, _, rel = line.partition("\t")
        if rel and to_suffix(rel) in TEXT_SUFFIXES and "w/crlf" in meta:
            errors.append(f"CRLF line ending detected in tracked text file: {rel}")

    for rel in tracked:
        if to_suffix(rel) in TEXT_SUFFIXES and has_utf8_bom(repo_root / rel):
            errors.append(f"UTF-8 BOM is not allowed in tracked text file: {rel}")


def print_policy_results(errors: list[str], warnings: list[str]) -> int:
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


def main() -> int:
    repo_root = Path(__file__).resolve().parents[2]
    errors: list[str] = []
    warnings: list[str] = []
    project_kind = load_project_kind(repo_root)

    try:
        tracked = tracked_files(repo_root)
        eol_lines = tracked_eol_lines(repo_root)
    except RuntimeError as exc:
        print(
            f"Policy check failed: unable to inspect git-tracked files ({exc}).",
            file=sys.stderr,
        )
        return 1

    tracked_set = set(tracked)
    package_name = resolve_package_name(collect_src_packages(repo_root / "src"), errors)

    pyproject = load_pyproject(repo_root, errors)
    if pyproject:
        validate_pyproject_policy(
            pyproject,
            project_kind=project_kind,
            errors=errors,
        )

    validate_required_paths(
        repo_root,
        tracked_set=tracked_set,
        project_kind=project_kind,
        errors=errors,
    )
    validate_test_layout(
        repo_root,
        project_kind=project_kind,
        errors=errors,
    )
    validate_package_contracts(
        repo_root,
        tracked=tracked,
        tracked_set=tracked_set,
        package_name=package_name,
        project_kind=project_kind,
        errors=errors,
    )
    validate_legacy_configs(
        tracked,
        tracked_set=tracked_set,
        errors=errors,
    )
    validate_python_paths(
        repo_root,
        tracked=tracked,
        package_name=package_name,
        project_kind=project_kind,
        warnings=warnings,
        errors=errors,
    )
    validate_launch_references(
        repo_root,
        tracked=tracked,
        package_name=package_name,
        errors=errors,
    )
    validate_text_file_policies(
        repo_root,
        tracked=tracked,
        eol_lines=eol_lines,
        errors=errors,
    )

    collect_silent_broad_exception_warnings(repo_root, tracked, warnings)
    collect_structure_size_guidance(repo_root, tracked, warnings)
    collect_docstring_guidance(repo_root, tracked, warnings)
    return print_policy_results(errors, warnings)


if __name__ == "__main__":
    raise SystemExit(main())
