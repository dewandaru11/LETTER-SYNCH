from pathlib import Path

p = Path("main.py")
s = p.read_text(encoding="utf-8")

# ============================================================
# 1. Tambahkan tanggal_surat ke tabel lembur / migrasi database
# ============================================================

old = '''        cur.execute("""
            CREATE TABLE IF NOT EXISTS lembur (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pegawai_id INTEGER,
                tanggal TEXT NOT NULL,
                jam_mulai TEXT,
                jam_selesai TEXT,
                rentang_jam TEXT,
                total_jam TEXT,
                laporan_lembur TEXT DEFAULT '',
                status TEXT DEFAULT 'Draft',
                file_path TEXT,
                pdf_path TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(pegawai_id) REFERENCES pegawai(id)
            )
        """)

        conn.commit()
        conn.close()
'''

new = '''        cur.execute("""
            CREATE TABLE IF NOT EXISTS lembur (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pegawai_id INTEGER,
                tanggal TEXT NOT NULL,
                tanggal_surat TEXT,
                jam_mulai TEXT,
                jam_selesai TEXT,
                rentang_jam TEXT,
                total_jam TEXT,
                laporan_lembur TEXT DEFAULT '',
                status TEXT DEFAULT 'Draft',
                file_path TEXT,
                pdf_path TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(pegawai_id) REFERENCES pegawai(id)
            )
        """)

        # Migrasi database lama.
        cur.execute("PRAGMA table_info(lembur)")
        columns = {row[1] for row in cur.fetchall()}

        if "tanggal_surat" not in columns:
            cur.execute(
                "ALTER TABLE lembur ADD COLUMN tanggal_surat TEXT"
            )

        conn.commit()
        conn.close()
'''

if old not in s:
    raise SystemExit("ERROR: blok ensure_lembur_table tidak ditemukan.")

s = s.replace(old, new, 1)


# ============================================================
# 2. Tambahkan field Tanggal Pembuatan Surat setelah Tanggal Lembur
# ============================================================

old = '''        form.addRow("Tanggal Lembur", self.lembur_date)

        # Jam mulai
'''

new = '''        form.addRow("Tanggal Lembur", self.lembur_date)

        # Tanggal pembuatan surat
        self.lembur_letter_date = QDateEdit()
        self.lembur_letter_date.setCalendarPopup(True)
        self.lembur_letter_date.setDate(QDate.currentDate())
        self.lembur_letter_date.setDisplayFormat("dd/MM/yyyy")
        self.lembur_letter_date.setMinimumHeight(38)
        self.lembur_letter_date.setMinimumWidth(180)
        self.lembur_letter_date.setKeyboardTracking(False)
        self.lembur_letter_date.setStyleSheet("""
            QDateEdit {
                background-color: #ffffff;
                color: #1f2937;
                border: 1px solid #b8c7cc;
                border-radius: 6px;
                padding: 7px 10px;
                font-size: 10pt;
            }
            QDateEdit:focus {
                border: 2px solid #0f6b78;
                background-color: #ffffff;
                color: #1f2937;
            }
            QDateEdit::drop-down {
                width: 30px;
                border: none;
            }
        """)
        form.addRow(
            "Tanggal Pembuatan Surat",
            self.lembur_letter_date
        )

        # Jam mulai
'''

if old not in s:
    raise SystemExit("ERROR: lokasi Tanggal Lembur tidak ditemukan.")

s = s.replace(old, new, 1)


# ============================================================
# 3. Ganti save_lembur()
# ============================================================

start = s.find("    def save_lembur(self):")
end = s.find("\n    def start_lembur_worker(", start)

if start == -1 or end == -1:
    raise SystemExit("ERROR: fungsi save_lembur tidak ditemukan.")

new_save = '''    def save_lembur(self):
        """Membuat dokumen daftar hadir kerja lembur."""

        employee_id = self.lembur_employee_combo.currentData()

        if not employee_id:
            QMessageBox.warning(
                self,
                "Data Belum Lengkap",
                "Silakan pilih pegawai terlebih dahulu."
            )
            return

        # ----------------------------------------------------
        # Tanggal lembur
        # ----------------------------------------------------
        tanggal = self.lembur_date.date()
        tanggal_iso = tanggal.toString("yyyy-MM-dd")

        hari_id = [
            "Senin",
            "Selasa",
            "Rabu",
            "Kamis",
            "Jumat",
            "Sabtu",
            "Minggu",
        ]

        bulan_id = [
            "Januari",
            "Februari",
            "Maret",
            "April",
            "Mei",
            "Juni",
            "Juli",
            "Agustus",
            "September",
            "Oktober",
            "November",
            "Desember",
        ]

        tanggal_text = (
            f"{hari_id[tanggal.dayOfWeek() - 1]}, "
            f"{tanggal.day()} "
            f"{bulan_id[tanggal.month() - 1]} "
            f"{tanggal.year()}"
        )

        # ----------------------------------------------------
        # Tanggal pembuatan surat
        # ----------------------------------------------------
        tanggal_surat = self.lembur_letter_date.date()
        tanggal_surat_iso = tanggal_surat.toString("yyyy-MM-dd")

        tanggal_surat_text = (
            f"{tanggal_surat.day()} "
            f"{bulan_id[tanggal_surat.month() - 1]} "
            f"{tanggal_surat.year()}"
        )

        jam_mulai = self.lembur_start.time().toString("HH:mm")
        jam_selesai = self.lembur_end.time().toString("HH:mm")
        rentang = self.lembur_range.text()
        total = self.lembur_total.text()

        laporan = self.lembur_report.toPlainText().strip()

        conn = get_connection()

        row = conn.execute("""
            SELECT *
            FROM pegawai
            WHERE id = ?
        """, (employee_id,)).fetchone()

        conn.close()

        if not row:
            QMessageBox.warning(
                self,
                "Data Tidak Ditemukan",
                "Data pegawai tidak ditemukan."
            )
            return

        nama = row["nama"] or ""
        nip = row["nip"] or ""
        jabatan = row["jabatan"] or ""
        unit_kerja = row["unit_kerja"] or ""

        template_name = "DAFTAR HADIR KERJA LEMBUR.docx"
        template_path = TEMPLATE_DIR / template_name

        if not template_path.exists():
            QMessageBox.information(
                self,
                "Template Belum Ada",
                "Form Lembur sudah siap digunakan.\\n\\n"
                f"Silakan masukkan template:\\n"
                f"{template_name}\\n\\n"
                f"ke folder:\\n{TEMPLATE_DIR}"
            )
            return

        values = {
            "NAMA": nama,
            "NIP": nip,
            "JABATAN": jabatan,
            "UNIT_KERJA": unit_kerja,

            "TANGGAL": tanggal_text,
            "TANGGAL_LEMBUR": tanggal_text,
            "TANGGAL_ISO": tanggal_iso,

            "TANGGAL_SURAT": tanggal_surat_text,
            "TANGGAL_PEMBUATAN_SURAT": tanggal_surat_text,
            "TANGGAL_SURAT_ISO": tanggal_surat_iso,

            "JAM_MULAI": jam_mulai,
            "JAM_SELESAI": jam_selesai,

            "RENTANG_JAM": rentang,
            "RENTANG_JAM_LEMBUR": rentang,

            "TOTAL_JAM": total,
            "TOTAL_JAM_LEMBUR": total,

            "LAPORAN_LEMBUR": laporan,
        }

        employee = {
            "id": employee_id,
            "nama": nama,
            "nip": nip,
            "jabatan": jabatan,
            "unit_kerja": unit_kerja,
        }

        self.start_lembur_worker(
            template_name,
            employee,
            tanggal_text,
            tanggal_iso,
            tanggal_surat_text,
            tanggal_surat_iso,
            values,
            employee_id,
            jam_mulai,
            jam_selesai,
            rentang,
            total,
            laporan,
        )
'''

s = s[:start] + new_save + s[end:]


# ============================================================
# 4. Tambahkan parameter tanggal surat pada worker
# ============================================================

old = '''        employee,
        tanggal_text,
        tanggal_iso,
        values,
        employee_id,
'''

new = '''        employee,
        tanggal_text,
        tanggal_iso,
        tanggal_surat_text,
        tanggal_surat_iso,
        values,
        employee_id,
'''

if old not in s:
    raise SystemExit("ERROR: parameter start_lembur_worker tidak ditemukan.")

s = s.replace(old, new, 1)


# ============================================================
# 5. Tambahkan tanggal surat ke database lembur
# ============================================================

old = '''            INSERT INTO lembur (
                pegawai_id,
                tanggal,
                jam_mulai,
                jam_selesai,
                rentang_jam,
                total_jam,
                laporan_lembur,
                status,
                file_path,
                pdf_path
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
'''

new = '''            INSERT INTO lembur (
                pegawai_id,
                tanggal,
                tanggal_surat,
                jam_mulai,
                jam_selesai,
                rentang_jam,
                total_jam,
                laporan_lembur,
                status,
                file_path,
                pdf_path
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
'''

if old not in s:
    raise SystemExit("ERROR: INSERT lembur tidak ditemukan.")

s = s.replace(old, new, 1)

old = '''        """, (
            employee_id,
            tanggal_iso,
            jam_mulai,
            jam_selesai,
            rentang,
            total,
            laporan,
            "Selesai",
            str(docx_path),
            str(pdf_path) if pdf_path else None,
        ))
'''

new = '''        """, (
            employee_id,
            tanggal_iso,
            tanggal_surat_iso,
            jam_mulai,
            jam_selesai,
            rentang,
            total,
            laporan,
            "Selesai",
            str(docx_path),
            str(pdf_path) if pdf_path else None,
        ))
'''

if old not in s:
    raise SystemExit("ERROR: values INSERT lembur tidak ditemukan.")

s = s.replace(old, new, 1)


# ============================================================
# 6. Arsip: gabungkan surat biasa + lembur
# ============================================================

old = '''        cur.execute("""
            SELECT
                surat.id,
                surat.jenis,
                pegawai.nama AS pegawai,
                surat.tanggal_surat,
                surat.status,
                surat.file_path,
                surat.pdf_path,
                surat.created_at
            FROM surat
            LEFT JOIN pegawai
                ON pegawai.id = surat.pegawai_id
            ORDER BY surat.id DESC
        """)
'''

new = '''        cur.execute("""
            SELECT
                surat.id AS id,
                surat.jenis AS jenis,
                pegawai.nama AS pegawai,
                surat.tanggal_surat AS tanggal_surat,
                surat.status AS status,
                surat.file_path AS file_path,
                surat.pdf_path AS pdf_path,
                surat.created_at AS created_at
            FROM surat
            LEFT JOIN pegawai
                ON pegawai.id = surat.pegawai_id

            UNION ALL

            SELECT
                (1000000 + lembur.id) AS id,
                'Lembur' AS jenis,
                pegawai.nama AS pegawai,
                COALESCE(
                    lembur.tanggal_surat,
                    lembur.tanggal
                ) AS tanggal_surat,
                lembur.status AS status,
                lembur.file_path AS file_path,
                lembur.pdf_path AS pdf_path,
                lembur.created_at AS created_at
            FROM lembur
            LEFT JOIN pegawai
                ON pegawai.id = lembur.pegawai_id

            ORDER BY created_at DESC
        """)
'''

if old not in s:
    raise SystemExit("ERROR: query load_archive tidak ditemukan.")

s = s.replace(old, new, 1)


# ============================================================
# 7. Setelah lembur selesai, refresh arsip
# ============================================================

marker = '''    def lembur_finished(self, result):
'''

pos = s.find(marker)

if pos == -1:
    raise SystemExit("ERROR: lembur_finished tidak ditemukan.")

# Tambahkan refresh tepat di awal fungsi jika belum ada.
after = pos + len(marker)

existing = s[after:after + 300]

if "self.load_archive()" not in existing:
    s = s[:after] + '''
        self.load_archive()
''' + s[after:]


p.write_text(s, encoding="utf-8")

print("PATCH MAIN.PY BERHASIL")
