# Cambios pendientes

Resumen de los cambios existentes en el repositorio al 9 de septiembre de 2026.

## Interfaz y flujo de trabajo

- Se añadió la edición de una actividad seleccionada desde el botón, doble clic o la tecla `F2`.
- Se añadió la opción de eliminar la importación cargada, con confirmación previa.
- Se incorporó la selección del formato del informe: `DOCX` u `ODT`.
- Se actualizaron los textos de la interfaz para referirse al resultado como informe, manteniendo la exportación combinada con Excel.

## Servicios y generación de informes

- Se añadió `app/services/odt_generator.py`, que convierte informes DOCX a ODT usando LibreOffice en modo headless.
- La conversión valida la instalación de LibreOffice, la existencia del documento de entrada y el resultado generado.
- Se añadió `storage.delete_import_batch` para eliminar lotes de importación.
- Se añadió `storage.update_activity_text` para actualizar los campos editables y reconstruir el texto de búsqueda y el JSON original.

## Importación y extracción de actividades

- Las filas cargadas desde CSV ahora se ordenan cronológicamente por `fecha_desde`.
- El extractor de historial agrupa todas las líneas de cada bloque de fecha como una única actividad.
- El nombre de la actividad del extractor ahora se puede configurar con `--actividad-desc`; por defecto es `Administrador AThi`.
- Se ajustó el uso del extractor y su documentación de línea de comandos.

## Archivos de datos

- Se añadió `agosto.csv` con las actividades exportadas para importar en la aplicación.
- Se añadió `data/imports/actividades_codecommit_athi_crosscheck.txt` con el resultado del cruce entre el historial de AThi y los commits del repositorio.

