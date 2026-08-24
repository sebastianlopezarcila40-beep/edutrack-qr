"""
Modulo legacy de tenants (SQLite aparte).
La fuente de verdad del multi-inquilino esta en app.py -> modelo Institucion (SQLAlchemy).
Este modulo se mantiene por compatibilidad con el panel de soporte antiguo.
"""
import sqlite3
from datetime import datetime

DB_TENANTS = "database/tenants.db"


def get_conn():
    return sqlite3.connect(DB_TENANTS)


def init_tenants_db():
    conn = get_conn()
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS instituciones(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT UNIQUE,
            nombre TEXT,
            municipio TEXT,
            departamento TEXT,
            logo TEXT,
            estado TEXT DEFAULT 'ACTIVA',
            fecha_creacion TEXT
        )
        """
    )
    cols = [r[1] for r in conn.execute("PRAGMA table_info(instituciones)").fetchall()]
    for col, ddl in [
        ("codigo", "ALTER TABLE instituciones ADD COLUMN codigo TEXT"),
        ("logo", "ALTER TABLE instituciones ADD COLUMN logo TEXT"),
        ("estado", "ALTER TABLE instituciones ADD COLUMN estado TEXT DEFAULT 'ACTIVA'"),
    ]:
        if col not in cols:
            try:
                conn.execute(ddl)
            except Exception:
                pass
    conn.commit()
    conn.close()


def crear_institucion(codigo, nombre, municipio, departamento, estado="ACTIVA"):
    init_tenants_db()
    conn = get_conn()
    conn.execute(
        """
        INSERT INTO instituciones(
            codigo, nombre, municipio, departamento, estado, fecha_creacion
        ) VALUES(?,?,?,?,?,?)
        """,
        (
            codigo,
            nombre,
            municipio,
            departamento,
            estado,
            datetime.now().strftime("%Y-%m-%d"),
        ),
    )
    conn.commit()
    conn.close()


def listar_instituciones():
    init_tenants_db()
    conn = get_conn()
    datos = conn.execute(
        "SELECT * FROM instituciones ORDER BY id DESC"
    ).fetchall()
    conn.close()
    return datos


def total_instituciones():
    init_tenants_db()
    conn = get_conn()
    total = conn.execute("SELECT COUNT(*) FROM instituciones").fetchone()[0]
    conn.close()
    return total
