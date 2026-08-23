from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def main() -> None:
    from app.services import storage
    from app.ui.main_window import App

    storage.init_db()
    app = App()
    print("App created. profiles:", len(storage.list_profiles()))
    app.destroy()
    print("GUI OK")


if __name__ == "__main__":
    main()
