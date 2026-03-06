# Release Checklist

1. `uv sync --all-extras --frozen`
2. `uv run ruff check .`
3. `uv run mypy src`
4. `uv run pytest`
5. `uv run python -m build`
6. `uv run pip-audit`
7. Install wheel locally and smoke test scripts:
   - `uv run python -m pip install dist/*.whl`
   - `media-quality-checker`
   - `sonarr-ui-helper`
8. Verify docs reflect current command paths and dependency workflow.
9. Tag and publish.
