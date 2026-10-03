
from .database import get_connection

def all_pegawai():
    c = get_connection()
    rows = c.execute("SELECT * FROM pegawai ORDER BY nama").fetchall()
    c.close()
    return rows

def add_pegawai(data):
    c = get_connection()
    c.execute("""INSERT INTO pegawai
        (nama,nip,jabatan,unit_kerja,masa_kerja,telepon,alamat)
        VALUES (?,?,?,?,?,?,?)""",
        (data["nama"],data["nip"],data["jabatan"],data["unit_kerja"],
         data["masa_kerja"],data["telepon"],data["alamat"]))
    c.commit()
    c.close()
