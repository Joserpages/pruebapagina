-- Tabla de estudiantes
CREATE TABLE IF NOT EXISTS estudiantes (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  token   TEXT UNIQUE NOT NULL,
  nombre  TEXT NOT NULL,
  curso   TEXT NOT NULL,
  nota    REAL NOT NULL,
  estado  TEXT CHECK(estado IN ('Aprobado','Reprobado')) NOT NULL,
  creado_en TEXT NOT NULL
);

-- Tabla de usuarios (para login)
CREATE TABLE IF NOT EXISTS usuarios (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  username TEXT UNIQUE NOT NULL,
  password_hash TEXT NOT NULL,
  creado_en TEXT NOT NULL
);
