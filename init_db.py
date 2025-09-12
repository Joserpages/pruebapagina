import os, sqlite3
from datetime import datetime
from werkzeug.security import generate_password_hash

DB_PATH = "db.sqlite3"

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    # Crear tablas
    c.executescript("""
    CREATE TABLE IF NOT EXISTS estudiantes (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      nombre  TEXT NOT NULL,
      curso   TEXT NOT NULL,
      nota    REAL NOT NULL,
      estado  TEXT CHECK(estado IN ('Aprobado','Reprobado')) NOT NULL,
      creado_en TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS usuarios (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT UNIQUE NOT NULL,
      password_hash TEXT NOT NULL,
      creado_en TEXT NOT NULL
    );
    """)

    # Crear admin si no existe
    row = c.execute("SELECT COUNT(*) FROM usuarios").fetchone()
    if row[0] == 0:
        c.execute(
            "INSERT INTO usuarios(username,password_hash,creado_en) VALUES(?,?,?)",
            ("admin", generate_password_hash("admin123"), datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        )
        print("✅ Usuario admin creado: admin / admin123")

    conn.commit()
    conn.close()
    print("✅ Base de datos inicializada en", DB_PATH)

if __name__ == "__main__":
    init_db()
