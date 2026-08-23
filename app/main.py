from __future__ import annotations

import sys
from pathlib import Path


def _bootstrap_path() -> None:
    if getattr(sys, "frozen", False):
        return
    root = Path(__file__).resolve().parents[1]
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))


_bootstrap_path()

from app.ui.main_window import run  # noqa: E402


if __name__ == "__main__":
    run()
