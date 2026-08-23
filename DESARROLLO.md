# Guía de desarrollo

Notas para trabajar en **Generador de Informes de Actividades** (`auto_shit_reports`).

## Stack

| Pieza | Tecnología |
|-------|------------|
| UI | CustomTkinter |
| Datos tabulares | pandas |
| Word | python-docx + plantilla `app/templates/informe.docx` |
| Excel | openpyxl + plantilla `app/templates/matriz.xlsx` |
| Persistencia | SQLite (`data/app.db`) — rutas/metadatos, no blobs de informes |
| Empaquetado | PyInstaller one-file, windowed (`build_exe.spec`) |

Python **3.10+**. Entorno principal de prueba: **Windows**.

## Setup local

```powershell
cd <ruta-del-repo>
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Dependencias de build del `.exe` (solo si empaquetas):

```powershell
pip install pyinstaller
```

## Ejecutar en desarrollo

```powershell
python -m app.main
# o
python app/main.py
```

Al arrancar, la app usa `data/` junto al proyecto (DB + exports).

## Estructura relevante

```
app/
  main.py              # entrypoint
  paths.py             # rutas dev vs frozen (PyInstaller)
  ui/main_window.py    # CustomTkinter
  services/            # CSV, filtros, word, excel, sqlite, código de informe
  templates/           # informe.docx, matriz.xlsx (editables)
  db/schema.sql
  assets/
scripts/               # build, smoke tests, verificación
sample_data/           # CSV de ejemplo
data/                  # local (gitignored salvo placeholders)
build_exe.spec
```

## Flujo de la app (para no romperlo)

1. Importar CSV → filas a SQLite + tabla filtrable.
2. Filtro por periodo (mes/año) y búsqueda de texto.
3. Formulario de cabecera (nombre, cargo, supervisor, equipo, ciudad).
4. Código auto: `ALT-INF-ACT-PERS-{MES}-{AÑO}`.
5. Generar Word + Excel en `data/exports/{codigo}/`.
6. Historial lee metadatos/rutas desde SQLite.

### Reglas de producto a respetar

- Columna **No.** en tablas generadas: renumerar desde **1** en el orden filtrado (no reutilizar el `No.` del CSV).
- Plantillas con marcadores/placeholders; no reconstruir layout desde cero.
- En DB: paths + metadata de informes; filas del CSV sí pueden persistirse.
- Empaquetado portable: un `.exe`, sin consola; crea `data\` junto al binario.

## CSV de entrada

Columnas esperadas (ver también `README.md`):

| Columna | Obligatoria |
|--------|-------------|
| `Actividad` | sí |
| `Fecha desde` | sí (`D/M/AAAA` o `AAAA-MM-DD`) |
| `No.`, `Cliente / Proyecto...`, `Fecha Hasta`, `Descripción`, `Estatus *` | no |

Hay alias en español (`Estado`, `Proyecto`, `Fecha inicio`, etc.) en el servicio de CSV.

CSV de prueba: `sample_data/actividades_ejemplo.csv`.

## Plantillas

Editar en Word/Excel y dejar en `app/templates/`:

- `informe.docx`
- `matriz.xlsx`

La lógica rellena cabecera y reescribe la tabla de actividades; logo/estilos base se conservan.

Si cambias nombres de placeholders o la forma de la tabla, actualiza el servicio correspondiente en `app/services/`.

## Base de datos

- Schema: `app/db/schema.sql`
- Runtime: `data/app.db` (no versionar)
- Exports: `data/exports/` (no versionar)

Tras cambiar el schema en dev, borra `data/app.db` o migra a mano; no hay migraciones automáticas formales todavía.

## Scripts útiles

| Script | Uso |
|--------|-----|
| `scripts/build_exe.ps1` | Build del `.exe` |
| `scripts/do_build.py` | Build auxiliar |
| `scripts/smoke_gui.py` | Smoke de GUI |
| `scripts/verify_generation.py` | Verifica generación Word/Excel |
| `scripts/verify_full.py` | Verificación más completa |
| `scripts/check.py` / `clean_test.py` | Checks puntuales |
| `scripts/inspect_word.py` | Inspección de docx |

Build Windows:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\build_exe.ps1
```

Salida: `dist\GeneradorInformes.exe`.

## Git

```bash
git remote -v
# origin  https://github.com/bocchii2/auto_shit_reports.git
```

No commitear:

- `.venv/`, `dist/`, `build/`
- `data/app.db`, `data/exports/`
- basura de debug (`err*.txt`, `out*.txt`, etc.)
- secretos (`.env`)

Sí versionar:

- código en `app/`
- plantillas
- `sample_data/`
- `requirements.txt`, `build_exe.spec`, scripts de build
- `README.md`, este archivo

## Checklist al tocar generación de informes

1. Probar con `sample_data/actividades_ejemplo.csv`.
2. Filtrar un mes con filas y uno vacío.
3. Revisar cabecera Word/Excel (nombre, código, periodo).
4. Confirmar renumeración `No.` 1..N.
5. Confirmar paths en historial / SQLite (no blobs).
6. Si tocaste `paths.py` o datas de PyInstaller, rebuild del `.exe` y prueba en carpeta limpia.

## Roadmap / fuera de scope actual

- Sin scraping AppSheet: el CSV se arma a mano u otra fuente.
- Sin multi-usuario ni servidor: app de escritorio local.
