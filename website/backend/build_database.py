"""
build_database.py
Minimal and safe DB initializer.
"""

import sqlite3
import sys
from pathlib import Path
import os

# ------------------------------------------------------------
# Match EXACTLY how db_ops.py decides DB location
# ------------------------------------------------------------
def _base_dir():
    if getattr(sys, "frozen", False):
        # Installed EXE: always store DB in LocalAppData\BrewInsPOS
        return Path(os.environ["LOCALAPPDATA"]) / "BrewInsPOS"
    else:
        # Running from source
        return Path(__file__).resolve().parents[1]

BASE_DIR = _base_dir()

# Use same structure as shipping app
DB_FILE  = BASE_DIR / "backend" / "pos.db"
SQL_FILE = BASE_DIR / "backend" / "models.sql"


def connect():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def create_schema(conn):
    if not SQL_FILE.exists():
        raise FileNotFoundError(f"Schema file not found: {SQL_FILE}")
    script = SQL_FILE.read_text(encoding="utf-8")
    conn.executescript(script)
    conn.commit()
    print("Tables created or already existed.")


def main():
    print("Initializing database...")

    # Ensure backend/ folder exists
    (BASE_DIR / "backend").mkdir(parents=True, exist_ok=True)

    # Create/overwrite database
    with connect() as conn:
        create_schema(conn)

    print(f"Database created at:\n{DB_FILE.resolve()}")


if __name__ == "__main__":
    main()
