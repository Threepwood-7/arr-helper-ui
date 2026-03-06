from __future__ import annotations

import re
import subprocess
import sys
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
LEGACY_ROOT_CONFIG_PATTERNS = (
    "config.toml",
    "config_example.toml",
    "*_config.toml",
    "*_config_example.toml",
    "*_config_totemp.toml",
)
CANONICAL_CONFIG_REQUIRED = ("config/app.defaults.toml", "config/app.example.toml")
CANONICAL_CONFIG_LOCAL = "config/app.local.toml"


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


def main() -> int:
    repo_root = Path(__file__).resolve().parents[2]
    errors: list[str] = []

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

    for rel in ROOT_REQUIRED:
        if rel not in tracked_set:
            errors.append(f"Missing required tracked file: {rel}")

    for rel in WINDOWS_SCRIPT_REQUIRED:
        if rel not in tracked_set:
            errors.append(f"Missing required tracked file: {rel}")

    if package_name:
        for rel in PACKAGE_REQUIRED:
            package_file = f"src/{package_name}/{rel}"
            if package_file not in tracked_set:
                errors.append(f"Missing required tracked file: {package_file}")

    for rel in tracked:
        if rel.lower().endswith(".cmd"):
            errors.append(f"Legacy .cmd scripts are not allowed: {rel}")

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
        return 1

    print("Project policy check passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
