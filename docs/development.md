# Development Workflow

## Primary Setup (uv)

```bash
uv sync --all-extras
```

## Fallback Setup (pip)

```bash
python -m pip install -e .[dev]
```

## Quality Gates

```bash
uv run ruff check .
uv run mypy src
uv run pytest
uv run python -m build
uv run pip-audit
```

## Pre-commit

```bash
uv run pre-commit install
uv run pre-commit run --all-files
```

## Notes

- Hatchling is the only build backend.
- uv is the primary lock/sync/command runner.
- Nox is intentionally not used.
