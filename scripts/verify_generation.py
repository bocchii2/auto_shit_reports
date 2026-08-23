from __future__ import annotations

import sys
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from app.services import storage
from app.services.csv_import import load_csv
from app.services.excel_generator import generate_excel
from app.services.filters import apply_filters
from app.services.report_code import build_codigo, build_periodo_label
from app.services.word_generator import generate_word


def main() -> None:
    storage.init_db()
    rows, warns = load_csv(ROOT / "sample_data" / "actividades_ejemplo.csv")
    print("rows", len(rows), "warns", warns)
    bid = storage.save_import_batch("actividades_ejemplo.csv", rows)
    acts = storage.get_activities(bid)
    jul = apply_filters(acts, mes=7, anio=2026)
    ago = apply_filters(acts, mes=8, anio=2026)
    print("jul", len(jul), "ago", len(ago))
    print("codigo", build_codigo(7, 2026), build_periodo_label(7, 2026))

    header = {
        "periodo_mes": 7,
        "periodo_anio": 2026,
        "codigo_informe": build_codigo(7, 2026),
        "nombre": "Jostin Abraham Cedeño Mosquera",
        "cargo": "Programador Junior Software",
        "cargo_firma": "Parametrizador Jr.",
        "supervisor": "Ing. Patricio Guzman",
        "cargo_supervisor": "Supervisor Inmediato",
        "empresa": "Altura S.A.",
        "equipo": "IDS",
        "ciudad": "Manta",
        "periodo": build_periodo_label(7, 2026),
        "periodo_corto": "Julio 2026",
        "fecha_informe": "20  de agosto de 2026",
    }
    out = ROOT / "data" / "exports" / header["codigo_informe"]
    out.mkdir(parents=True, exist_ok=True)
    w = generate_word(
        ROOT / "app" / "templates" / "informe.docx",
        out / f"{header['codigo_informe']}.docx",
        header=header,
        activities=jul,
    )
    x = generate_excel(
        ROOT / "app" / "templates" / "matriz.xlsx",
        out / f"Matriz-{header['codigo_informe']}.xlsx",
        header=header,
        activities=jul,
    )
    print("word", w, w.exists(), w.stat().st_size)
    print("excel", x, x.exists(), x.stat().st_size)

    rid = storage.save_report(
        codigo_informe=header["codigo_informe"],
        periodo_mes=7,
        periodo_anio=2026,
        nombre=header["nombre"],
        cargo=header["cargo"],
        supervisor=header["supervisor"],
        header=header,
        filters={"mes": 7, "anio": 2026},
        activity_ids=[int(a["id"]) for a in jul],
        path_docx=str(w),
        path_xlsx=str(x),
    )
    print("report_id", rid, "reports", len(storage.list_reports()))

    ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    with ZipFile(w) as z:
        root = ET.fromstring(z.read("word/document.xml"))
        texts = [t.text for t in root.iter(ns + "t") if t.text]
        blob = "\n".join(texts)
        print("has_nombre", "Jostin" in blob)
        print("has_periodo", "julio de 2026" in blob.lower())
        print("activity_count_in_doc", blob.count("Entregado"))
        for line in texts[:20]:
            print(" ", line[:120])
    print("OK")


if __name__ == "__main__":
    main()
