CREATE TABLE IF NOT EXISTS import_batches (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  filename TEXT NOT NULL,
  imported_at TEXT NOT NULL,
  row_count INTEGER NOT NULL
);

CREATE TABLE IF NOT EXISTS activities (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  batch_id INTEGER NOT NULL REFERENCES import_batches(id) ON DELETE CASCADE,
  numero INTEGER,
  actividad TEXT,
  cliente_proyecto TEXT,
  fecha_desde TEXT,
  fecha_hasta TEXT,
  descripcion TEXT,
  estatus TEXT,
  raw_json TEXT NOT NULL,
  search_text TEXT
);

CREATE INDEX IF NOT EXISTS idx_activities_batch ON activities(batch_id);
CREATE INDEX IF NOT EXISTS idx_activities_fecha ON activities(fecha_desde);
CREATE INDEX IF NOT EXISTS idx_activities_search ON activities(search_text);

CREATE TABLE IF NOT EXISTS reports (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  created_at TEXT NOT NULL,
  codigo_informe TEXT NOT NULL,
  periodo_mes INTEGER NOT NULL,
  periodo_anio INTEGER NOT NULL,
  nombre TEXT,
  cargo TEXT,
  supervisor TEXT,
  header_json TEXT,
  filter_json TEXT,
  activity_count INTEGER,
  path_docx TEXT,
  path_xlsx TEXT
);

CREATE TABLE IF NOT EXISTS report_activities (
  report_id INTEGER NOT NULL REFERENCES reports(id) ON DELETE CASCADE,
  activity_id INTEGER NOT NULL REFERENCES activities(id) ON DELETE CASCADE,
  PRIMARY KEY (report_id, activity_id)
);

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
);

CREATE INDEX IF NOT EXISTS idx_profiles_name ON profiles(name);
