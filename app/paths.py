from __future__ import annotations

import sys
from pathlib import Path


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def resource_root() -> Path:
    """Read-only assets bundled with the app (templates, schema)."""
    if is_frozen():
        return Path(getattr(sys, "_MEIPASS"))
    return Path(__file__).resolve().parents[1]


def app_root() -> Path:
    """Writable base next to the .exe (or project root in dev)."""
    if is_frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def templates_dir() -> Path:
    if is_frozen():
        return resource_root() / "templates"
    return resource_root() / "app" / "templates"


def schema_path() -> Path:
    if is_frozen():
        return resource_root() / "db" / "schema.sql"
    return resource_root() / "app" / "db" / "schema.sql"


def data_dir() -> Path:
    path = app_root() / "data"
    path.mkdir(parents=True, exist_ok=True)
    return path


def exports_dir() -> Path:
    path = data_dir() / "exports"
    path.mkdir(parents=True, exist_ok=True)
    return path


def db_path() -> Path:
    return data_dir() / "app.db"
