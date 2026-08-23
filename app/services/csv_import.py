from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
from dateutil import parser as date_parser

# Canonical internal keys -> accepted CSV header aliases (lowercase)
COLUMN_ALIASES: dict[str, list[str]] = {
    "numero": ["numero", "no", "no.", "n", "nº", "n°", "#"],
    "actividad": ["actividad", "actividad/tarea", "tarea", "actividades"],
    "cliente_proyecto": [
        "cliente / proyecto / servicio / aplicación",
        "cliente / proyecto / servicio / aplicacion",
        "cliente_proyecto",
        "cliente",
        "proyecto",
        "cliente/proyecto",
        "cliente / proyecto",
    ],
    "fecha_desde": [
        "fecha desde",
        "fecha_desde",
        "desde",
        "fecha inicio",
        "fecha_inicio",
        "inicio",
    ],
    "fecha_hasta": [
        "fecha hasta",
        "fecha_hasta",
        "hasta",
        "fecha fin",
        "fecha_fin",
        "fin",
    ],
    "descripcion": ["descripcion", "descripción", "detalle", "descripcion actividad"],
    "estatus": ["estatus", "estatus *", "estado", "status"],
}

REQUIRED = ["actividad", "fecha_desde"]


def _norm_header(h: str) -> str:
    return " ".join(str(h).strip().lower().split())


def _build_rename_map(columns: list[str]) -> dict[str, str]:
    alias_to_canon: dict[str, str] = {}
    for canon, aliases in COLUMN_ALIASES.items():
        for a in aliases:
            alias_to_canon[_norm_header(a)] = canon

    rename: dict[str, str] = {}
    for col in columns:
        key = _norm_header(col)
        if key in alias_to_canon:
            rename[col] = alias_to_canon[key]
    return rename


def _parse_date(value: Any) -> str | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, pd.Timestamp):
        return value.date().isoformat()
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "nat"}:
        return None
    # Excel serial number as string
    try:
        as_float = float(text)
        if as_float > 20000:
            dt = pd.to_datetime(as_float, unit="D", origin="1899-12-30")
            return dt.date().isoformat()
    except (ValueError, OverflowError):
        pass
    try:
        dt = date_parser.parse(text, dayfirst=True, fuzzy=True)
        return dt.date().isoformat()
    except (ValueError, OverflowError, TypeError):
        return text


def load_csv(path: str | Path) -> tuple[list[dict[str, Any]], list[str]]:
    """Load CSV into normalized activity dicts.

    Returns (rows, warnings).
    Raises ValueError if required columns are missing.
    """
    path = Path(path)
    last_err: Exception | None = None
    df: pd.DataFrame | None = None
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            df = pd.read_csv(path, encoding=enc, dtype=str, keep_default_na=False)
            break
        except Exception as exc:  # noqa: BLE001
            last_err = exc
    if df is None:
        raise ValueError(f"No se pudo leer el CSV: {last_err}")

    rename = _build_rename_map(list(df.columns))
    df = df.rename(columns=rename)

    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(
            "Faltan columnas obligatorias en el CSV: "
            + ", ".join(missing)
            + f". Columnas encontradas: {list(df.columns)}"
        )

    warnings: list[str] = []
    optional = ["numero", "cliente_proyecto", "fecha_hasta", "descripcion", "estatus"]
    for col in optional:
        if col not in df.columns:
            df[col] = ""
            warnings.append(f"Columna opcional ausente, se usará vacía: {col}")

    rows: list[dict[str, Any]] = []
    for i, rec in enumerate(df.to_dict(orient="records"), start=1):
        fecha_desde = _parse_date(rec.get("fecha_desde"))
        fecha_hasta = _parse_date(rec.get("fecha_hasta")) or fecha_desde
        numero_raw = str(rec.get("numero") or "").strip()
        try:
            numero = int(float(numero_raw)) if numero_raw else i
        except ValueError:
            numero = i
        row = {
            "numero": numero,
            "actividad": str(rec.get("actividad") or "").strip(),
            "cliente_proyecto": str(rec.get("cliente_proyecto") or "").strip(),
            "fecha_desde": fecha_desde or "",
            "fecha_hasta": fecha_hasta or "",
            "descripcion": str(rec.get("descripcion") or "").strip(),
            "estatus": str(rec.get("estatus") or "").strip() or "Entregado",
        }
        if not row["actividad"]:
            warnings.append(f"Fila {i}: actividad vacía, se omitió")
            continue
        rows.append(row)

    if not rows:
        raise ValueError("El CSV no contiene filas de actividad válidas")

    return rows, warnings
