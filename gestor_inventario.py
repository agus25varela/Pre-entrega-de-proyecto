import csv
import sqlite3

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from colorama import init, Fore, Style
init(autoreset=True)

DB_NAME = "inventario.db"


def conectar_db() -> sqlite3.Connection:
    """Crea y devuelve una conexión a la base de datos SQLite."""
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn


def inicializar_bd() -> None:
    """Crea la tabla 'productos' si no existe con la estructura requerida."""
    with conectar_db() as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS productos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL,
                descripcion TEXT,
                cantidad INTEGER NOT NULL,
                precio REAL NOT NULL,
                categoria TEXT
            )
            """
        )
        conn.commit()