from __future__ import annotations

from datetime import date, datetime
from typing import Any


def _to_date(value: str | None) -> date | None:
    if not value:
        return None
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(text[:10], fmt).date()
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(text[:10]).date()
    except ValueError:
        return None


def apply_filters(
    rows: list[dict[str, Any]],
    *,
    mes: int | None = None,
    anio: int | None = None,
    fecha_desde: str | None = None,
    fecha_hasta: str | None = None,
    query: str | None = None,
) -> list[dict[str, Any]]:
    d_from = _to_date(fecha_desde)
    d_to = _to_date(fecha_hasta)
    q = (query or "").strip().lower()

    result: list[dict[str, Any]] = []
    for row in rows:
        fd = _to_date(row.get("fecha_desde"))
        fh = _to_date(row.get("fecha_hasta")) or fd

        if mes is not None or anio is not None:
            # include if either bound falls in period, or spans the month
            matched = False
            for d in (fd, fh):
                if d is None:
                    continue
                if mes is not None and d.month != mes:
                    continue
                if anio is not None and d.year != anio:
                    continue
                matched = True
                break
            if not matched and fd and fh and mes is not None and anio is not None:
                # activity spans across the month
                month_start = date(anio, mes, 1)
                if mes == 12:
                    month_end = date(anio + 1, 1, 1)
                else:
                    month_end = date(anio, mes + 1, 1)
                if fd < month_end and fh >= month_start:
                    matched = True
            if not matched:
                continue

        if d_from and fd and fd < d_from:
            continue
        if d_to and fd and fd > d_to:
            continue

        if q:
            hay = " ".join(
                str(row.get(k) or "")
                for k in (
                    "actividad",
                    "cliente_proyecto",
                    "descripcion",
                    "estatus",
                    "numero",
                )
            ).lower()
            if q not in hay:
                continue

        result.append(row)

    return result
