@echo off
setlocal
python -m pip install -r requirements.txt
python -m pip install pyinstaller
pyinstaller --noconfirm --windowed --name Surat-Menyurat-RSHD --add-data "templates;templates" --add-data "assets;assets" main.py
echo.
echo SELESAI. EXE ada di dist\Surat-Menyurat-RSHD\
pause
