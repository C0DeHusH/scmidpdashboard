from __future__ import annotations

import sqlite3
import os
import shutil
from contextlib import contextmanager
from pathlib import Path
from typing import Iterable, Sequence

PROJECT_ROOT = Path(__file__).resolve().parents[2]
STATE_ROOT = Path(os.environ.get("SCM_DATA_DIR", str(PROJECT_ROOT / "uploads")))
DB_PATH = STATE_ROOT / "aging" / "aging.db"
SEED_DB_PATH = PROJECT_ROOT / "data" / "aging" / "aging_seed.db"
NO_DATA_MARKER = STATE_ROOT / ".scm_no_data"

SCHEMA = """
PRAGMA journal_mode=WAL;
PRAGMA foreign_keys=ON;

CREATE TABLE IF NOT EXISTS imports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    filename TEXT NOT NULL,
    imported_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    as_of_date TEXT NOT NULL,
    row_count INTEGER NOT NULL DEFAULT 0,
    import_mode TEXT NOT NULL DEFAULT 'replace'
);

-- Kept for backward compatibility with v1.0/v1.1 databases.  In v1.2 the
-- imported AREA column is authoritative; this table is only a fallback for
-- older files that do not contain AREA.
CREATE TABLE IF NOT EXISTS branch_area (
    branch_key TEXT PRIMARY KEY,
    branch_name TEXT NOT NULL,
    area TEXT NOT NULL DEFAULT 'UNMAPPED',
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS units (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    import_id INTEGER,
    source_row INTEGER,
    branch_key TEXT,
    branch_original TEXT,
    area TEXT,
    incoming_date TEXT,
    created_on TEXT,
    barcode TEXT,
    description TEXT,
    qty REAL NOT NULL DEFAULT 1,
    standard_description TEXT,
    amount REAL NOT NULL DEFAULT 0,
    company TEXT,
    engine_no TEXT,
    chassis TEXT,
    location TEXT,
    brand TEXT,
    color TEXT,
    imported_as_of TEXT,
    FOREIGN KEY(import_id) REFERENCES imports(id) ON DELETE SET NULL
);

CREATE INDEX IF NOT EXISTS idx_units_branch ON units(branch_key);
CREATE INDEX IF NOT EXISTS idx_units_area ON units(area);
CREATE INDEX IF NOT EXISTS idx_units_brand ON units(brand);
CREATE INDEX IF NOT EXISTS idx_units_std_desc ON units(standard_description);
CREATE INDEX IF NOT EXISTS idx_units_engine ON units(engine_no);
CREATE INDEX IF NOT EXISTS idx_units_chassis ON units(chassis);
"""


def connect() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    # First-run seed only. A global Clear Data writes NO_DATA_MARKER so the
    # supplied seed cannot silently repopulate after an intentional reset.
    if not DB_PATH.exists() and SEED_DB_PATH.exists() and not NO_DATA_MARKER.exists():
        shutil.copy2(SEED_DB_PATH, DB_PATH)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def _column_names(conn: sqlite3.Connection, table: str) -> set[str]:
    return {row[1] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()}


def init_db() -> None:
    with connect() as conn:
        # Upgrade older v1.0/v1.1 databases without forcing the user to delete
        # their local database.
        conn.executescript(SCHEMA)
        cols = _column_names(conn, "units")
        if "area" not in cols:
            conn.execute("ALTER TABLE units ADD COLUMN area TEXT")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_units_area ON units(area)")
        conn.commit()


@contextmanager
def transaction():
    conn = connect()
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def executemany(sql: str, rows: Iterable[Sequence]) -> None:
    with transaction() as conn:
        conn.executemany(sql, rows)


def clear_data() -> None:
    """Remove all imported Motorcycle Aging data while preserving the schema."""
    init_db()
    with transaction() as conn:
        conn.execute("DELETE FROM units")
        conn.execute("DELETE FROM imports")
        conn.execute("DELETE FROM branch_area")


def record_count() -> int:
    init_db()
    with connect() as conn:
        return int(conn.execute("SELECT COUNT(*) FROM units").fetchone()[0])


def backup_database(target: str | Path) -> None:
    """Create a consistent SQLite snapshot, including any WAL-resident pages."""
    target = Path(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.unlink(missing_ok=True)
    init_db()
    with connect() as src, sqlite3.connect(target) as dst:
        src.backup(dst)
        dst.commit()


def restore_database(source: str | Path) -> None:
    """Restore a previously created SQLite snapshot into the live Aging DB."""
    source = Path(source)
    if not source.exists():
        raise FileNotFoundError(source)
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(source) as src, sqlite3.connect(DB_PATH) as dst:
        src.backup(dst)
        dst.commit()
