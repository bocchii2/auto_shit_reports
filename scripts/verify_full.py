"""Verificación integral: perfiles + generación de Word/Excel."""
import os, sys
DB = os.path.join("data", "app.db")
if os.path.exists(DB):
    os.remove(DB)

from app.services import storage
storage.init_db()

p = storage.create_profile(
    name="Carlos Mendoza Rojas", nombre="Carlos Mendoza Rojas", cargo="Analista de Procesos",
    supervisor="Laura Pérez", cargo_supervisor="Jefa de Área",
    empresa="Tech Solutions SAC", equipo="Equipo Operativo", ciudad="Lima"
)
p_id = p
prof = storage.get_profile(p_id)
print("perfil creado id:", p_id, "nombre:", prof["nombre"])
profiles = storage.list_profiles()
print("list profiles:", [x.get("name") for x in profiles])

prof = storage.get_profile(p_id)
print("get by id:", prof["cargo"], prof["ciudad"])

storage.update_profile(p_id, cargo="Analista Senior", ciudad="Arequipa")
print("update:", storage.get_profile(p_id)["ciudad"])
storage.delete_profile(p_id)
print("delete -> list:", storage.list_profiles())

# Generación de informes con el perfil restaurado
mid = storage.create_profile(
    name="María Elena Vásquez Torres", nombre="María Elena Vásquez Torres", cargo="Analista de Procesos",
    supervisor="Laura Pérez", cargo_supervisor="Jefa de Área",
    empresa="Tech Solutions SAC", equipo="Equipo Operativo", ciudad="Lima"
)
prof = storage.get_profile(mid)
print("perfil 2:", prof["nombre"], prof["cargo"], prof["ciudad"])

# test get_default_profile
print("default:", storage.get_default_profile())

from app.services.word_generator import generate_word
from app.services.excel_generator import generate_excel
from app.services.report_code import build_periodo_label
from app.paths import templates_dir
from datetime import date

tpl_dir = templates_dir()
wp = tpl_dir / "informe.docx"
ep = tpl_dir / "matriz.xlsx"

rows = [
    {"numero": "1", "Actividad": "Revisión de tickets",
     "Cliente / Proyecto / Servicio / Aplicación": "Servicio X",
     "Fecha desde": "2026-07-02", "Fecha Hasta": "2026-07-04",
     "Descripción": "Atendí 15 tickets críticos de la base de datos.", "Estatus *": "Entregado"},
    {"numero": "2", "Actividad": "Actualización documentos",
     "Cliente / Proyecto / Servicio / Aplicación": "Proyecto Y",
     "Fecha desde": "2026-07-08", "Fecha Hasta": "2026-07-09",
     "Descripción": "Documentación API actualizada v2.", "Estatus *": "Entregado"},
]

# Map CSV-style rows -> generator activity keys (No. is auto-1-based)
activities = [
    {
        "actividad": r["Actividad"],
        "cliente_proyecto": r["Cliente / Proyecto / Servicio / Aplicación"],
        "fecha_desde": r["Fecha desde"],
        "fecha_hasta": r["Fecha Hasta"],
        "descripcion": r["Descripción"],
        "estatus": r["Estatus *"],
    }
    for r in rows
]

header = {
    "periodo_mes": 7,
    "periodo_anio": 2026,
    "nombre": prof["nombre"],
    "cargo": prof["cargo"],
    "supervisor": prof["supervisor"],
    "empresa": prof["empresa"],
    "cargo_firma": prof.get("cargo_firma") or prof["cargo"],
    "cargo_supervisor": prof.get("cargo_supervisor") or "Supervisor Inmediato",
    "fecha_informe": "7 de agosto de 2026",
    "periodo": build_periodo_label(7, 2026),
}
os.makedirs("data/exports", exist_ok=True)
generate_word(wp, "data/exports/test_word.docx", header=header, activities=activities)
generate_excel(ep, "data/exports/test_excel.xlsx", header=header, activities=activities)

# Verify Word SDT table No. column is 1-based and renumbered
doc = __import__("docx").Document("data/exports/test_word.docx")
act_tbl = None
from docx.oxml.ns import qn
for child in doc.element.body:
    if child.tag == qn("w:sdt") and (tbl := child.find(qn("w:tbl"))):
        joined = "".join(t.text or "" for t in tbl.iter(qn("w:t"))).lower()
        if "actividad" in joined:
            act_tbl = tbl
            break
nos = []
if act_tbl:
    data_rows = list(act_tbl.iterchildren(qn("w:tr")))[1:]
    for tr in data_rows:
        first_tc = list(tr.iterchildren(qn("w:tc")))[0]
        nos.append("".join(t.text for t in first_tc.iter(qn("w:t"))))
print("word No. column:", nos)

# Verify Excel No. column (B) is 1-based from row 8
wb = __import__("openpyxl").load_workbook("data/exports/test_excel.xlsx")
ws = wb.active
ex_nos = [ws[f"B{r}"].value for r in range(8, 8 + len(activities))]
print("excel No. column:", ex_nos)
print("FINAL OK")
