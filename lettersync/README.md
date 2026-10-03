# SURAT-MENYURAT RSHD — Versi 1

Aplikasi desktop Windows untuk pembuatan surat berdasarkan template Word RSUD H. Damanhuri Barabai.

## Fitur versi pertama
- Dashboard
- Data pegawai
- Surat Izin Tidak Masuk Kerja
- Formulir Permintaan dan Pemberian Cuti
- Penyimpanan database SQLite
- Generate dokumen Word dari template
- Arsip dasar

## Struktur
- `main.py` — aplikasi utama
- `app/database.py` — SQLite
- `app/generator.py` — pengisian template Word
- `templates/` — template yang digunakan aplikasi
- `templates/original/` — salinan template asli
- `output/` — hasil dokumen
- `data/` — database

## Menjalankan di Windows

1. Install Python 3.11 atau lebih baru.
2. Buka folder proyek di Visual Studio Code.
3. Buka Terminal.
4. Jalankan:

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

## Membuat EXE

Install PyInstaller:

```powershell
pip install pyinstaller
```

Lalu:

```powershell
pyinstaller --noconfirm --windowed --name Surat-Menyurat-RSHD --add-data "templates;templates" --add-data "assets;assets" main.py
```

File EXE akan berada di:

`dist\Surat-Menyurat-RSHD\Surat-Menyurat-RSHD.exe`

Untuk versi portable satu file, setelah pengujian stabil bisa dibuat:

```powershell
pyinstaller --noconfirm --onefile --windowed --name Surat-Menyurat-RSHD --add-data "templates;templates" --add-data "assets;assets" main.py
```

## Catatan template
Template asli pengguna tidak diubah dan disimpan di `templates/original/`.
Aplikasi memakai salinan placeholder:
- `templates/surat_izin.docx`
- `templates/surat_cuti.docx`

Template kedua `izin_sakit.docx` disimpan sebagai referensi untuk versi berikutnya.

## PDF
Konversi Word ke PDF dapat memakai LibreOffice jika `soffice` tersedia. Versi pertama fokus pada Word terlebih dahulu agar format template tetap terjaga.
