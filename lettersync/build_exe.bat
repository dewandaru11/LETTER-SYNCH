@echo off
setlocal

echo ==========================================
echo             BUILD LETTERSYNC
echo ==========================================
echo.

python -m pip install -r requirements.txt
if errorlevel 1 goto :error

python -m pip install pyinstaller
if errorlevel 1 goto :error

echo.
echo Membersihkan build sebelumnya...

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo Membuat LetterSync.exe...
echo.

python -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --windowed ^
    --name LetterSync ^
    --add-data "templates;templates" ^
    --add-data "assets;assets" ^
    main.py

if errorlevel 1 goto :error

echo.
echo ==========================================
echo BUILD BERHASIL
echo ==========================================
echo.
echo EXE:
echo dist\LetterSync\LetterSync.exe
echo.

goto :end

:error
echo.
echo ==========================================
echo BUILD GAGAL
echo ==========================================
echo.

:end
pause
