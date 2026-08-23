from __future__ import annotations

import copy
import re
import shutil
from datetime import date
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.table import Table, _Cell

from app.services.report_code import (
    build_codigo,
    build_periodo_label,
    format_fecha_informe,
)

ACTIVITY_HEADERS = [
    "No.",
    "Actividad",
    "Cliente / Proyecto / Servicio / Aplicación",
    "Fecha desde",
    "Fecha Hasta",
    "Descripción",
    "Estatus *",
]


def _set_run_text(paragraph, text: str) -> None:
    if paragraph.runs:
        paragraph.runs[0].text = text
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(text)


def _set_cell_text(cell, text: str) -> None:
    lines = (text or "").split("\n")
    paragraphs = cell.paragraphs
    if not paragraphs:
        cell.text = text
        return
    _set_run_text(paragraphs[0], lines[0] if lines else "")
    for i, line in enumerate(lines[1:], start=1):
        if i < len(paragraphs):
            _set_run_text(paragraphs[i], line)
        else:
            p = cell.add_paragraph()
            p.add_run(line)
    for p in list(paragraphs)[len(lines) :]:
        for run in p.runs:
            run.text = ""


def _format_date_display(iso: str) -> str:
    if not iso:
        return ""
    try:
        y, m, d = iso[:10].split("-")
        return f"{int(d)}/{int(m)}/{y}"
    except ValueError:
        return iso


def _join_text(element) -> str:
    return "".join(t.text or "" for t in element.iter(qn("w:t")))


def _unique_cells(row):
    seen = set()
    for cell in row.cells:
        key = id(cell._tc)
        if key in seen:
            continue
        seen.add(key)
        yield cell


def _fill_meta_table(table: Table, mapping: dict[str, str]) -> None:
    for row in table.rows:
        cells = list(_unique_cells(row))
        if len(cells) < 2:
            continue
        label = cells[0].text.strip().lower()
        if not label:
            continue
        for key, val in mapping.items():
            if key in label:
                _set_cell_text(cells[1], val)
                break


def _fill_signature_table(
    table: Table,
    *,
    nombre: str,
    cargo_firma: str,
    supervisor: str,
    cargo_supervisor: str,
    empresa: str,
) -> None:
    if len(table.rows) < 2:
        return
    try:
        if len(table.rows[1].cells) >= 3:
            _set_cell_text(table.rows[1].cells[0], f"{nombre}\n{cargo_firma}")
            _set_cell_text(table.rows[1].cells[2], f"{supervisor}\n{cargo_supervisor}")
        if len(table.rows) >= 3 and len(table.rows[2].cells) >= 3:
            _set_cell_text(table.rows[2].cells[0], empresa)
            _set_cell_text(table.rows[2].cells[2], empresa)
    except Exception:
        pass


def _set_header_codigo(doc: Document, codigo: str) -> None:
    """Update report code in header table without wiping INFORME labels."""
    for section in doc.sections:
        header = section.header
        for table in header.tables:
            for row in table.rows:
                for cell in _unique_cells(row):
                    text = cell.text or ""
                    flat = text.replace("\r", "")
                    if "ALT-INF-ACT-PERS" not in flat.replace("–", "-") and "ALT–INF–ACT–PERS" not in flat:
                        # try joined runs
                        joined = _join_text(cell._tc)
                        if "ALT-INF-ACT-PERS" not in joined.replace("–", "-"):
                            continue
                        text = joined
                    # Preserve lines when possible
                    lines = [ln.strip() for ln in re.split(r"[\n\r]+", text) if ln.strip()]
                    if not lines:
                        # single blob
                        new_blob = re.sub(
                            r"ALT[\-–]INF[\-–]ACT[\-–]PERS.*",
                            codigo,
                            text.replace("\n", ""),
                            flags=re.I,
                        )
                        if "INFORME" in text.upper() and "INFORME DE ACTIVIDADES" not in new_blob.upper():
                            new_blob = f"INFORME\nINFORME DE ACTIVIDADES\n{codigo}"
                        _set_cell_text(cell, new_blob)
                        continue
                    new_lines = []
                    replaced = False
                    for ln in lines:
                        if "ALT-INF-ACT-PERS" in ln.replace("–", "-"):
                            new_lines.append(codigo)
                            replaced = True
                        else:
                            new_lines.append(ln)
                    if not replaced:
                        new_lines.append(codigo)
                    # Ensure standard 3-line header block if labels exist
                    labels = [ln for ln in new_lines if "ALT-INF-ACT-PERS" not in ln.replace("–", "-")]
                    if any("INFORME" in ln.upper() for ln in labels):
                        keep = []
                        for ln in labels:
                            if ln.upper() == "INFORME" or "INFORME DE ACTIVIDADES" in ln.upper():
                                keep.append(ln)
                        if not keep:
                            keep = labels[:2]
                        _set_cell_text(cell, "\n".join(keep + [codigo]))
                    else:
                        _set_cell_text(cell, "\n".join(new_lines))


def _find_sdt_activity_table(doc: Document):
    """Return (sdt_element, tbl_element) for the activities content control."""
    body = doc.element.body
    for child in body:
        if child.tag != qn("w:sdt"):
            continue
        tables = list(child.iter(qn("w:tbl")))
        if not tables:
            continue
        joined = _join_text(tables[0]).lower()
        if "actividad" in joined or "estatus" in joined or "no." in joined:
            return child, tables[0]
    return None, None


def _clear_table_rows_keep_header(tbl_elm, header_rows: int = 1):
    rows = list(tbl_elm.iterchildren(qn("w:tr")))
    for tr in rows[header_rows:]:
        tbl_elm.remove(tr)
    return rows[0] if rows else None


def _row_cells(tr_elm):
    """Return tc elements in row order (handles nested wrappers)."""
    direct = list(tr_elm.iterchildren(qn("w:tc")))
    if direct:
        return direct
    # fallback: document-order unique tc under this tr
    return list(tr_elm.iter(qn("w:tc")))


def _fill_row_cells(tr_elm, values: list[str]) -> None:
    tcs = _row_cells(tr_elm)
    for i, val in enumerate(values):
        if i >= len(tcs):
            break
        tc = tcs[i]
        paragraphs = list(tc.iter(qn("w:p")))
        # Prefer direct child paragraphs when present
        direct_ps = list(tc.iterchildren(qn("w:p")))
        if direct_ps:
            paragraphs = direct_ps
        if not paragraphs:
            p = OxmlElement("w:p")
            r = OxmlElement("w:r")
            t = OxmlElement("w:t")
            t.text = val
            t.set(qn("xml:space"), "preserve")
            r.append(t)
            p.append(r)
            tc.append(p)
            continue
        texts = list(paragraphs[0].iter(qn("w:t")))
        if texts:
            texts[0].text = val
            texts[0].set(qn("xml:space"), "preserve")
            for t in texts[1:]:
                t.text = ""
        else:
            r = OxmlElement("w:r")
            t = OxmlElement("w:t")
            t.text = val
            t.set(qn("xml:space"), "preserve")
            r.append(t)
            paragraphs[0].append(r)
        for p in paragraphs[1:]:
            for t in p.iter(qn("w:t")):
                t.text = ""


def _populate_sdt_table(tbl_elm, activities: list[dict[str, Any]]) -> None:
    rows = list(tbl_elm.iterchildren(qn("w:tr")))
    if not rows:
        return
    header_tr = rows[0]
    template_tr = rows[1] if len(rows) > 1 else copy.deepcopy(header_tr)

    # remove data rows
    for tr in rows[1:]:
        tbl_elm.remove(tr)

    for i, act in enumerate(activities, start=1):
        new_tr = copy.deepcopy(template_tr)
        values = [
            str(i),
            str(act.get("actividad") or ""),
            str(act.get("cliente_proyecto") or ""),
            _format_date_display(str(act.get("fecha_desde") or "")),
            _format_date_display(str(act.get("fecha_hasta") or "")),
            str(act.get("descripcion") or ""),
            str(act.get("estatus") or ""),
        ]
        _fill_row_cells(new_tr, values)
        tbl_elm.append(new_tr)


def _set_table_borders(table: Table) -> None:
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tblPr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = borders.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "666666")


def _style_header_cell(cell) -> None:
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), "1F4E79")
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)
    for p in cell.paragraphs:
        for run in p.runs:
            run.bold = True
            run.font.size = Pt(9)
            try:
                run.font.color.rgb = RGBColor(255, 255, 255)
            except Exception:
                pass


def _append_activities_fallback(doc: Document, activities: list[dict[str, Any]]) -> None:
    """If no SDT table, append a simple table before the signature block."""
    title = doc.add_paragraph()
    run = title.add_run("RESUMEN DE ACTIVIDADES")
    run.bold = True
    run.font.size = Pt(12)
    title_elm = title._element

    table = doc.add_table(rows=1, cols=len(ACTIVITY_HEADERS))
    try:
        table.style = "Table Grid"
    except KeyError:
        _set_table_borders(table)
    for i, name in enumerate(ACTIVITY_HEADERS):
        cell = table.rows[0].cells[i]
        _set_cell_text(cell, name)
        _style_header_cell(cell)
    for i, act in enumerate(activities, start=1):
        row = table.add_row().cells
        values = [
            str(i),
            str(act.get("actividad") or ""),
            str(act.get("cliente_proyecto") or ""),
            _format_date_display(str(act.get("fecha_desde") or "")),
            _format_date_display(str(act.get("fecha_hasta") or "")),
            str(act.get("descripcion") or ""),
            str(act.get("estatus") or ""),
        ]
        for col, val in enumerate(values):
            _set_cell_text(row[col], val)

    body = doc.element.body
    tbls = [c for c in body if c.tag == qn("w:tbl")]
    if len(tbls) >= 2:
        act_tbl = tbls[-1]
        # place before last original-looking signature: second-to-last before append was signature
        # after append, signature is second last if only 2 main tables + new
        # Find signature by 3 columns empty first row heuristic
        sig = None
        for t in reversed(tbls[:-1]):
            rows = list(t.iterchildren(qn("w:tr")))
            if len(rows) == 3:
                sig = t
                break
        if sig is not None:
            body.remove(act_tbl)
            sig.addprevious(act_tbl)
            if title_elm.getparent() is not None:
                body.remove(title_elm)
            act_tbl.addprevious(title_elm)


def _replace_sample_names(doc: Document, replacements: list[tuple[str, str]]) -> None:
    """Targeted paragraph-level replace without flattening whole body."""
    def do_paragraph(p) -> None:
        text = p.text
        if not text:
            return
        new = text
        for old, val in replacements:
            if old and old in new:
                new = new.replace(old, val)
        if new != text:
            _set_run_text(p, new)

    for p in doc.paragraphs:
        do_paragraph(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in _unique_cells(row):
                for p in cell.paragraphs:
                    do_paragraph(p)
    for section in doc.sections:
        for p in section.header.paragraphs:
            do_paragraph(p)
        for table in section.header.tables:
            for row in table.rows:
                for cell in _unique_cells(row):
                    for p in cell.paragraphs:
                        do_paragraph(p)


def generate_word(
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

    doc = Document(str(output_path))

    mes = int(header["periodo_mes"])
    anio = int(header["periodo_anio"])
    codigo = header.get("codigo_informe") or build_codigo(mes, anio)
    nombre = header.get("nombre") or ""
    cargo = header.get("cargo") or ""
    supervisor = header.get("supervisor") or ""
    empresa = header.get("empresa") or "Altura S.A."
    cargo_firma = header.get("cargo_firma") or cargo
    cargo_supervisor = header.get("cargo_supervisor") or "Supervisor Inmediato"
    if cargo_firma and not str(cargo_firma).endswith("."):
        cargo_firma = f"{cargo_firma}."
    cargo_dot = cargo if not cargo or str(cargo).endswith(".") else f"{cargo}."

    fecha_inf = header.get("fecha_informe")
    if not fecha_inf:
        today = date.today()
        fecha_inf = format_fecha_informe(today.day, today.month, today.year)
    periodo = header.get("periodo") or build_periodo_label(mes, anio)

    # 1) Header code
    _set_header_codigo(doc, codigo)

    # 2) Meta table
    if doc.tables:
        _fill_meta_table(
            doc.tables[0],
            {
                "fecha de informe": fecha_inf,
                "empleado": nombre,
                "cargo": cargo_dot,
                "periodo": periodo,
            },
        )

    # 3) Signature table (last python-docx table = signature; SDT table is hidden)
    if len(doc.tables) >= 2:
        _fill_signature_table(
            doc.tables[-1],
            nombre=nombre,
            cargo_firma=cargo_firma,
            supervisor=supervisor,
            cargo_supervisor=cargo_supervisor,
            empresa=empresa,
        )
    elif doc.tables:
        _fill_signature_table(
            doc.tables[0],
            nombre=nombre,
            cargo_firma=cargo_firma,
            supervisor=supervisor,
            cargo_supervisor=cargo_supervisor,
            empresa=empresa,
        )

    # 4) Light text replacements for leftovers (do not flatten body)
    _replace_sample_names(
        doc,
        [
            ("Del 01 al 31 de julio de 2026", periodo),
            ("7  de agosto de 2026", fecha_inf),
            ("Programador Junior Software.", cargo_dot),
            ("Parametrizador Jr.", cargo_firma),
            ("Ing. Patricio Guzman", supervisor),
            ("Ing. Patricio Guzmán", supervisor),
            ("Jostin Abraham Cedeño Mosquera", nombre),
            ("Altura S.A.", empresa),
        ],
    )

    # Re-assert meta/signature after paragraph replaces
    if doc.tables:
        _fill_meta_table(
            doc.tables[0],
            {
                "fecha de informe": fecha_inf,
                "empleado": nombre,
                "cargo": cargo_dot,
                "periodo": periodo,
            },
        )
        if len(doc.tables) >= 2:
            _fill_signature_table(
                doc.tables[-1],
                nombre=nombre,
                cargo_firma=cargo_firma,
                supervisor=supervisor,
                cargo_supervisor=cargo_supervisor,
                empresa=empresa,
            )

    # 5) Activities: prefer SDT table in template
    _sdt, sdt_tbl = _find_sdt_activity_table(doc)
    if sdt_tbl is not None:
        _populate_sdt_table(sdt_tbl, activities)
    else:
        _append_activities_fallback(doc, activities)

    doc.save(str(output_path))
    return output_path
