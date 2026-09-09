from __future__ import annotations

import json
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from app.paths import db_path, schema_path


def _connect() -> sqlite3.Connection:
    path = db_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db() -> None:
    schema = schema_path().read_text(encoding="utf-8")
    with _connect() as conn:
        conn.executescript(schema)
        # Ensure profiles exists even if DB was created before this feature
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS profiles (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              name TEXT NOT NULL UNIQUE,
              nombre TEXT NOT NULL DEFAULT '',
              cargo TEXT NOT NULL DEFAULT '',
              cargo_firma TEXT NOT NULL DEFAULT '',
              supervisor TEXT NOT NULL DEFAULT '',
              cargo_supervisor TEXT NOT NULL DEFAULT '',
              empresa TEXT NOT NULL DEFAULT '',
              equipo TEXT NOT NULL DEFAULT '',
              ciudad TEXT NOT NULL DEFAULT '',
              is_default INTEGER NOT NULL DEFAULT 0,
              created_at TEXT NOT NULL,
              updated_at TEXT NOT NULL
            )
            """
        )
        conn.commit()


def list_profiles() -> list[dict[str, Any]]:
    with _connect() as conn:
        rows = conn.execute(
            """
            SELECT * FROM profiles
            ORDER BY is_default DESC, name COLLATE NOCASE ASC
            """
        ).fetchall()
        return [dict(r) for r in rows]


def get_profile(profile_id: int) -> dict[str, Any] | None:
    with _connect() as conn:
        row = conn.execute(
            "SELECT * FROM profiles WHERE id = ?",
            (profile_id,),
        ).fetchone()
        return dict(row) if row else None


def get_default_profile() -> dict[str, Any] | None:
    with _connect() as conn:
        row = conn.execute(
            """
            SELECT * FROM profiles
            WHERE is_default = 1
            ORDER BY id ASC
            LIMIT 1
            """
        ).fetchone()
        if row:
            return dict(row)
        row = conn.execute(
            "SELECT * FROM profiles ORDER BY id ASC LIMIT 1"
        ).fetchone()
        return dict(row) if row else None


def create_profile(
    *,
    name: str,
    nombre: str = "",
    cargo: str = "",
    cargo_firma: str = "",
    supervisor: str = "",
    cargo_supervisor: str = "",
    empresa: str = "",
    equipo: str = "",
    ciudad: str = "",
    is_default: bool = False,
) -> int:
    name = (name or "").strip()
    if not name:
        raise ValueError("El nombre del perfil es obligatorio")
    now = datetime.now(timezone.utc).isoformat()
    with _connect() as conn:
        if is_default:
            conn.execute("UPDATE profiles SET is_default = 0")
        try:
            cur = conn.execute(
                """
                INSERT INTO profiles (
                    name, nombre, cargo, cargo_firma, supervisor, cargo_supervisor,
                    empresa, equipo, ciudad, is_default, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    name,
                    nombre.strip(),
                    cargo.strip(),
                    cargo_firma.strip(),
                    supervisor.strip(),
                    cargo_supervisor.strip(),
                    empresa.strip(),
                    equipo.strip(),
                    ciudad.strip(),
                    1 if is_default else 0,
                    now,
                    now,
                ),
            )
        except sqlite3.IntegrityError as exc:
            raise ValueError(f"Ya existe un perfil llamado '{name}'") from exc
        conn.commit()
        return int(cur.lastrowid)


def update_profile(
    profile_id: int,
    *,
    name: str | None = None,
    nombre: str | None = None,
    cargo: str | None = None,
    cargo_firma: str | None = None,
    supervisor: str | None = None,
    cargo_supervisor: str | None = None,
    empresa: str | None = None,
    equipo: str | None = None,
    ciudad: str | None = None,
    is_default: bool | None = None,
) -> None:
    current = get_profile(profile_id)
    if not current:
        raise ValueError("Perfil no encontrado")

    data = {
        "name": (name if name is not None else current["name"]).strip(),
        "nombre": (nombre if nombre is not None else current["nombre"] or "").strip(),
        "cargo": (cargo if cargo is not None else current["cargo"] or "").strip(),
        "cargo_firma": (
            cargo_firma if cargo_firma is not None else current["cargo_firma"] or ""
        ).strip(),
        "supervisor": (
            supervisor if supervisor is not None else current["supervisor"] or ""
        ).strip(),
        "cargo_supervisor": (
            cargo_supervisor
            if cargo_supervisor is not None
            else current["cargo_supervisor"] or ""
        ).strip(),
        "empresa": (empresa if empresa is not None else current["empresa"] or "").strip(),
        "equipo": (equipo if equipo is not None else current["equipo"] or "").strip(),
        "ciudad": (ciudad if ciudad is not None else current["ciudad"] or "").strip(),
        "is_default": (
            int(bool(is_default))
            if is_default is not None
            else int(current.get("is_default") or 0)
        ),
    }
    if not data["name"]:
        raise ValueError("El nombre del perfil es obligatorio")

    now = datetime.now(timezone.utc).isoformat()
    with _connect() as conn:
        if data["is_default"]:
            conn.execute(
                "UPDATE profiles SET is_default = 0 WHERE id != ?",
                (profile_id,),
            )
        try:
            conn.execute(
                """
                UPDATE profiles SET
                    name = ?, nombre = ?, cargo = ?, cargo_firma = ?,
                    supervisor = ?, cargo_supervisor = ?, empresa = ?,
                    equipo = ?, ciudad = ?, is_default = ?, updated_at = ?
                WHERE id = ?
                """,
                (
                    data["name"],
                    data["nombre"],
                    data["cargo"],
                    data["cargo_firma"],
                    data["supervisor"],
                    data["cargo_supervisor"],
                    data["empresa"],
                    data["equipo"],
                    data["ciudad"],
                    data["is_default"],
                    now,
                    profile_id,
                ),
            )
        except sqlite3.IntegrityError as exc:
            raise ValueError(f"Ya existe un perfil llamado '{data['name']}'") from exc
        conn.commit()


def delete_profile(profile_id: int) -> None:
    with _connect() as conn:
        conn.execute("DELETE FROM profiles WHERE id = ?", (profile_id,))
        conn.commit()


def set_default_profile(profile_id: int) -> None:
    if not get_profile(profile_id):
        raise ValueError("Perfil no encontrado")
    now = datetime.now(timezone.utc).isoformat()
    with _connect() as conn:
        conn.execute("UPDATE profiles SET is_default = 0")
        conn.execute(
            "UPDATE profiles SET is_default = 1, updated_at = ? WHERE id = ?",
            (now, profile_id),
        )
        conn.commit()


def save_import_batch(filename: str, rows: list[dict[str, Any]]) -> int:
    now = datetime.now(timezone.utc).isoformat()
    with _connect() as conn:
        cur = conn.execute(
            "INSERT INTO import_batches (filename, imported_at, row_count) VALUES (?, ?, ?)",
            (filename, now, len(rows)),
        )
        batch_id = int(cur.lastrowid)
        for row in rows:
            search = " ".join(
                str(row.get(k) or "")
                for k in (
                    "actividad",
                    "cliente_proyecto",
                    "descripcion",
                    "estatus",
                )
            ).lower()
            conn.execute(
                """
                INSERT INTO activities (
                    batch_id, numero, actividad, cliente_proyecto,
                    fecha_desde, fecha_hasta, descripcion, estatus,
                    raw_json, search_text
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    batch_id,
                    row.get("numero"),
                    row.get("actividad"),
                    row.get("cliente_proyecto"),
                    row.get("fecha_desde"),
                    row.get("fecha_hasta"),
                    row.get("descripcion"),
                    row.get("estatus"),
                    json.dumps(row, ensure_ascii=False),
                    search,
                ),
            )
        conn.commit()
        return batch_id


def list_batches() -> list[dict[str, Any]]:
    with _connect() as conn:
        rows = conn.execute(
            "SELECT * FROM import_batches ORDER BY id DESC"
        ).fetchall()
        return [dict(r) for r in rows]


def delete_import_batch(batch_id: int) -> None:
    with _connect() as conn:
        cur = conn.execute("DELETE FROM import_batches WHERE id = ?", (batch_id,))
        if cur.rowcount == 0:
            raise ValueError("Lote de importación no encontrado")
        conn.commit()


def update_activity_text(
    activity_id: int,
    *,
    actividad: str,
    cliente_proyecto: str,
    descripcion: str,
    estatus: str,
) -> None:
    actividad = actividad.strip()
    if not actividad:
        raise ValueError("La actividad no puede quedar vacía")

    with _connect() as conn:
        row = conn.execute(
            "SELECT * FROM activities WHERE id = ?", (activity_id,)
        ).fetchone()
        if not row:
            raise ValueError("Actividad no encontrada")

        raw = dict(row)
        raw.update(
            {
                "actividad": actividad,
                "cliente_proyecto": cliente_proyecto,
                "descripcion": descripcion,
                "estatus": estatus,
            }
        )
        search = " ".join(
            str(raw.get(k) or "")
            for k in ("actividad", "cliente_proyecto", "descripcion", "estatus")
        ).lower()
        conn.execute(
            """
            UPDATE activities
            SET actividad = ?, cliente_proyecto = ?, descripcion = ?, estatus = ?,
                raw_json = ?, search_text = ?
            WHERE id = ?
            """,
            (
                actividad,
                cliente_proyecto,
                descripcion,
                estatus,
                json.dumps(raw, ensure_ascii=False),
                search,
                activity_id,
            ),
        )
        conn.commit()


def get_activities(batch_id: int | None = None) -> list[dict[str, Any]]:
    with _connect() as conn:
        if batch_id is None:
            row = conn.execute(
                "SELECT id FROM import_batches ORDER BY id DESC LIMIT 1"
            ).fetchone()
            if not row:
                return []
            batch_id = int(row["id"])
        rows = conn.execute(
            """
            SELECT * FROM activities
            WHERE batch_id = ?
            ORDER BY
              CASE WHEN fecha_desde IS NULL OR fecha_desde = '' THEN 1 ELSE 0 END,
              fecha_desde,
              id
            """,
            (batch_id,),
        ).fetchall()
        return [dict(r) for r in rows]


def save_report(
    *,
    codigo_informe: str,
    periodo_mes: int,
    periodo_anio: int,
    nombre: str,
    cargo: str,
    supervisor: str,
    header: dict[str, Any],
    filters: dict[str, Any],
    activity_ids: list[int],
    path_docx: str | None,
    path_xlsx: str | None,
) -> int:
    now = datetime.now(timezone.utc).isoformat()
    with _connect() as conn:
        cur = conn.execute(
            """
            INSERT INTO reports (
                created_at, codigo_informe, periodo_mes, periodo_anio,
                nombre, cargo, supervisor, header_json, filter_json,
                activity_count, path_docx, path_xlsx
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                now,
                codigo_informe,
                periodo_mes,
                periodo_anio,
                nombre,
                cargo,
                supervisor,
                json.dumps(header, ensure_ascii=False),
                json.dumps(filters, ensure_ascii=False),
                len(activity_ids),
                path_docx,
                path_xlsx,
            ),
        )
        report_id = int(cur.lastrowid)
        for aid in activity_ids:
            conn.execute(
                "INSERT INTO report_activities (report_id, activity_id) VALUES (?, ?)",
                (report_id, aid),
            )
        conn.commit()
        return report_id


def list_reports(limit: int = 100) -> list[dict[str, Any]]:
    with _connect() as conn:
        rows = conn.execute(
            "SELECT * FROM reports ORDER BY id DESC LIMIT ?",
            (limit,),
        ).fetchall()
        return [dict(r) for r in rows]
