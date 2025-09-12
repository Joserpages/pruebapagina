import sqlite3, uuid
from datetime import datetime

DB_PATH = "db.sqlite3"
PASSING_GRADE = 60.0

def seed_data():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    estudiantes = [
        ("Ana Pérez", "Matemática", 85),
        ("Luis Gómez", "Historia", 72),
        ("María López", "Lenguaje", 55),
        ("Carlos Ramírez", "Física", 92),
        ("Sofía Hernández", "Química", 40),
        ("Pedro Castillo", "Artes", 78),
        ("Laura Martínez", "Biología", 61),
        ("Jorge Díaz", "Informática", 100),
        ("Elena Torres", "Música", 47),
        ("Andrés Morales", "Inglés", 70),
    ]

    for nombre, curso, nota in estudiantes:
        token = uuid.uuid4().hex
        estado = "Aprobado" if nota >= PASSING_GRADE else "Reprobado"
        creado = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        c.execute("""
            INSERT INTO estudiantes (token, nombre, curso, nota, estado, creado_en)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (token, nombre, curso, nota, estado, creado))

    conn.commit()
    conn.close()
    print("✅ Se insertaron 10 estudiantes de prueba.")

if __name__ == "__main__":
    seed_data()
