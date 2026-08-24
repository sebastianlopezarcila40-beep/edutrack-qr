import sqlite3
from datetime import datetime

DB_TENANTS = "database/tenants.db"

def crear_institucion(nombre,
                      municipio,
                      departamento,
                      logo=""):

    conn = sqlite3.connect(DB_TENANTS)

    conn.execute("""
    INSERT INTO instituciones
    (
        nombre,
        municipio,
        departamento,
        logo,
        fecha_creacion
    )
    VALUES
    (?,?,?,?,?)
    """,
    (
        nombre,
        municipio,
        departamento,
        logo,
        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ))

    conn.commit()
    conn.close()