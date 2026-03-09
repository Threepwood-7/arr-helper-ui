## Qt Composition Playbook

This playbook defines reusable structure guidance for large `qt_app` repositories.

### Goals

- Keep top-level window/dialog classes thin and readable.
- Isolate behavior by feature domain.
- Preserve import compatibility while internals evolve.
- Keep GUI tests aligned with feature boundaries.

### Composition Boundaries

Use the main `QMainWindow`/`QDialog` class as a facade and delegate to focused collaborators.

- `actions`:
  action catalog, shortcut registration, menu wiring.
- `layout`:
  split/tree orchestration, active-panel targeting, rebuild flows.
- `operations`:
  operation request construction and dispatch orchestration.
- `persistence`:
  state codecs, settings roundtrip, session restore/apply.
- `status`:
  status-bar text, role visuals, source/target context labels.

### Suggested Package Shape

```text
src/<package>/
|-- ui/
|   `-- window/
|       |-- __init__.py
|       |-- actions.py
|       |-- layout.py
|       |-- operations.py
|       |-- persistence.py
|       `-- status.py
```

Names can vary, but keep one module per concern.

### Public API Stability

- Preserve existing top-level imports (`window.py`, `dialogs/settings_dialog.py`, etc.).
- Keep facade method names stable where tests or integrations rely on them.
- Treat collaborator modules as internal implementation details.

### Test Slicing Conventions

Avoid single massive GUI test files. Split by behavior domain.

- `test_window_layout_actions.py`
- `test_window_operations.py`
- `test_window_persistence.py`
- `test_settings_dialog_sections.py`
- `test_settings_dialog_live_preview.py`

Prefer shared fakes/fixtures in `tests/gui/_helpers.py` and `tests/conftest.py`.

### Refactor Workflow

1. Add characterization tests around current behavior.
2. Extract one concern at a time into collaborator modules.
3. Keep behavior and widget identity contract unchanged.
4. Run `hatch run lint:check`, `hatch run lint:types`, `hatch run lint:policy`, `hatch run test`.
5. Repeat for the next concern.
