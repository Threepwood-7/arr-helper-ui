import os
import sys
from pathlib import Path

# Ensure Qt tests run headless in CI/local terminals without a display server.
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

# Ensure src-layout package modules are importable when tests are executed from tests/.
ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))
