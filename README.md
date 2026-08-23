# Generador de Informes de Actividades

App de escritorio (Python + CustomTkinter) para armar el informe mensual de actividades a partir de un **CSV**:

- **Word** `.docx` — `ALT-INF-ACT-PERS-{MES}-{AÑO}`
- **Excel** `.xlsx` — matriz de detalle
- **SQLite** — guarda filas del CSV y las **rutas** de los informes generados

## Requisitos

- Python 3.10+
- Windows (probado en este entorno)

## Instalación

```bash
cd C:\Users\USUARIO\Desktop\xd
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Ejecutar

```bash
python -m app.main
```

o:

```bash
python app/main.py
```

## Ejecutable Windows (.exe)

```powershell
powershell -ExecutionPolicy Bypass -File scripts\build_exe.ps1
```

El binario queda en:

`dist\GeneradorInformes.exe`

- Es portable (un solo archivo).
- Al abrirlo crea `data\` junto al `.exe` (SQLite + exports).
- Las plantillas van embebidas dentro del ejecutable.


## Flujo de uso

1. **Importar CSV** (pestaña Actividades). Puedes usar `sample_data/actividades_ejemplo.csv`.
2. Elige **mes/año** del periodo y, si quieres, busca por texto.
3. En **Cabecera y generar** completa:
   - Nombre, cargo, supervisor
   - Equipo y ciudad (van al Excel)
   - El código se calcula solo: `ALT-INF-ACT-PERS-JUL-2026`
4. Pulsa **Generar Word + Excel**.
5. Los archivos salen en `data/exports/{codigo}/`.
6. El **Historial** muestra generaciones previas y permite abrir la carpeta.

## Columnas del CSV

| Columna CSV | Obligatoria | Notas |
|-------------|-------------|--------|
| `No.` | No | Si falta se numera en orden |
| `Actividad` | **Sí** | |
| `Cliente / Proyecto / Servicio / Aplicación` | No | |
| `Fecha desde` | **Sí** | `D/M/AAAA` o `AAAA-MM-DD` |
| `Fecha Hasta` | No | Por defecto = fecha desde |
| `Descripción` | No | |
| `Estatus *` | No | Default `Entregado` |

Se aceptan alias en español (p. ej. `Estado`, `Proyecto`, `Fecha inicio`).

## Plantillas

Están en `app/templates/`:

- `informe.docx` — copia del informe de referencia
- `matriz.xlsx` — copia de la matriz de referencia

La app rellena cabecera y reescribe la tabla de actividades **sin redibujar** logo ni estilos base.

## Datos locales

| Ruta | Contenido |
|------|-----------|
| `data/app.db` | SQLite (batches, activities, reports) |
| `data/exports/` | Informes generados |

No se guardan los binarios del informe en la base: solo **rutas** y metadatos. Las filas del CSV sí se persisten.

## Sin AppSheet (por ahora)

Esta versión no hace scraping. Exporta/arma un CSV con las columnas de arriba (manual desde AppSheet u otra fuente) y genera los informes.

## Estructura

```
app/
  main.py
  ui/main_window.py
  services/   # csv, filtros, word, excel, sqlite, código
  templates/  # plantillas docx/xlsx
  db/schema.sql
sample_data/
data/
```
