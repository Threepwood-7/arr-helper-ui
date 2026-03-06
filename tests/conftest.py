from __future__ import annotations

import pytest
from PySide6.QtWidgets import QWidget


@pytest.fixture
def window(qtbot: object) -> QWidget:
    widget = QWidget()
    add_widget = getattr(qtbot, "addWidget")
    add_widget(widget)
    widget.show()
    yield widget
    widget.close()
