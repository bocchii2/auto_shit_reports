from __future__ import annotations

import shutil
from copy import copy
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

from app.services.report_code import build_periodo_corto

# Layout from reference matrix
HEADER_CELLS = {
    "nombre": "E3",
    "equipo": "E4",
    "periodo": "E5",
    "ciudad": "E6",
}
TABLE_START_ROW = 8  # first data row
TABLE_HEADER_ROW = 7
COLS = {
    "numero": "B",
    "actividad": "C",
    "cliente_proyecto": "D",
    "fecha_desde": "E",
    "fecha_hasta": "F",
    "descripcion": "G",
    "estatus": "H",
}


def _copy_style(src_cell, dst_cell) -> None:
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.border = copy(src_cell.border)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.number_format = src_cell.number_format
        dst_cell.protection = copy(src_cell.protection)
        dst_cell.alignment = copy(src_cell.alignment)


def _to_excel_date(iso: str):
    if not iso:
        return None
    text = str(iso).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text[:10], fmt).date()
        except ValueError:
            continue
    return text


def _clear_old_rows(ws, start_row: int, max_scan: int = 500) -> None:
    end = start_row
    for r in range(start_row, start_row + max_scan):
        val = ws[f"B{r}"].value
        if val is None or str(val).strip() == "":
            # stop after a few empty rows if we already saw data
            if r > start_row:
                break
        end = r
    # Also clear note row area loosely; keep note if far below
    for r in range(start_row, end + 1):
        for col in "BCDEFGH":
            cell = ws[f"{col}{r}"]
            cell.value = None


def generate_excel(
    template_path: str | Path,
    output_path: str | Path,
    *,
    header: dict[str, Any],
    activities: list[dict[str, Any]],
) -> Path:
    template_path = Path(template_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_path, output_path)

    wb = load_workbook(str(output_path))
    ws = wb.active

    mes = int(header["periodo_mes"])
    anio = int(header["periodo_anio"])
    nombre = header.get("nombre") or ""
    equipo = header.get("equipo") or ""
    ciudad = header.get("ciudad") or ""
    periodo = header.get("periodo_corto") or build_periodo_corto(mes, anio)

    ws[HEADER_CELLS["nombre"]] = nombre
    ws[HEADER_CELLS["equipo"]] = equipo
    ws[HEADER_CELLS["periodo"]] = periodo
    ws[HEADER_CELLS["ciudad"]] = ciudad

    # Capture template style from first data row if present
    style_row = TABLE_START_ROW
    template_cells = {col: ws[f"{col}{style_row}"] for col in "BCDEFGH"}

    _clear_old_rows(ws, TABLE_START_ROW)

    thin = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    wrap = Alignment(wrap_text=True, vertical="top")

    for i, act in enumerate(activities):
        r = TABLE_START_ROW + i
        # ensure row height reasonable
        ws.row_dimensions[r].height = max(ws.row_dimensions[r].height or 30, 45)

        values = {
            "B": i + 1,
            "C": act.get("actividad") or "",
            "D": act.get("cliente_proyecto") or "",
            "E": _to_excel_date(str(act.get("fecha_desde") or "")),
            "F": _to_excel_date(str(act.get("fecha_hasta") or "")),
            "G": act.get("descripcion") or "",
            "H": act.get("estatus") or "",
        }
        for col, val in values.items():
            cell = ws[f"{col}{r}"]
            src = template_cells[col]
            _copy_style(src, cell)
            if cell.border is None or cell.border.left is None:
                cell.border = thin
            if col in {"C", "G"}:
                cell.alignment = wrap
            cell.value = val
            if col in {"E", "F"} and hasattr(val, "year"):
                cell.number_format = "D/M/YYYY"

    # Move / rewrite note below data
    note_row = TABLE_START_ROW + max(len(activities), 1) + 2
    note = "*Nota: Campo solo obligatorio para personal bajo Servicios profesionales"
    # Clear old note around row 29 if present
    for r in range(TABLE_START_ROW + 1, 80):
        for col in ("B", "C", "D"):
            cell = ws[f"{col}{r}"]
            if cell.value and "Nota" in str(cell.value):
                cell.value = None
    ws[f"C{note_row}"] = note

    wb.save(str(output_path))
    return output_path
