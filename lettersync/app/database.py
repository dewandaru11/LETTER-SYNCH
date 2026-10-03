
from pathlib import Path
import sqlite3
import sys

if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH = BASE_DIR / "data" / "surat_menyurat.db"

def get_connection():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_connection()
    cur = conn.cursor()
    cur.executescript("""
    CREATE TABLE IF NOT EXISTS pegawai (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nama TEXT NOT NULL,
        nip TEXT,
        jabatan TEXT,
        unit_kerja TEXT,
        masa_kerja TEXT,
        telepon TEXT,
        alamat TEXT
    );

    CREATE TABLE IF NOT EXISTS pejabat (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nama TEXT NOT NULL,
        nip TEXT,
        jabatan TEXT,
        jenis TEXT
    );

    CREATE TABLE IF NOT EXISTS surat (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        jenis TEXT NOT NULL,
        nomor TEXT,
        pegawai_id INTEGER,
        tanggal_surat TEXT,
        status TEXT DEFAULT 'Draft',
        file_path TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(pegawai_id) REFERENCES pegawai(id)
    );
    """)
    conn.commit()
    conn.close()



