@echo off
title Build Portable EXE — Log Detail & QR Generator
echo ================================================
echo  Membangun Aplikasi Portable (.exe)
echo ================================================
echo.

:: Cek Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python tidak ditemukan.
    pause & exit /b 1
)

:: Install PyInstaller jika belum ada
echo [1/3] Memastikan PyInstaller tersedia...
pip install pyinstaller --quiet

:: Install semua library yang dibutuhkan
echo [2/3] Memastikan semua library tersedia...
pip install pandas openpyxl "qrcode[pil]" Pillow PyQt6 --quiet

:: Build
echo [3/3] Membangun file EXE (mungkin butuh 1-3 menit)...
echo.

pyinstaller ^
    --onefile ^
    --windowed ^
    --name "LogDetailQR" ^
    --add-data "app.py;." ^
    --hidden-import "pandas" ^
    --hidden-import "openpyxl" ^
    --hidden-import "qrcode" ^
    --hidden-import "PIL" ^
    --hidden-import "PyQt6" ^
    --hidden-import "PyQt6.QtPrintSupport" ^
    --hidden-import "qrcode.image.pil" ^
    --collect-all "qrcode" ^
    --collect-all "PIL" ^
    --collect-all "openpyxl" ^
    app.py

if errorlevel 1 (
    echo.
    echo [ERROR] Build gagal. Periksa pesan error di atas.
    pause
    exit /b 1
)

echo.
echo ================================================
echo  BUILD BERHASIL!
echo  File EXE: dist\LogDetailQR.exe
echo  Salin file EXE ke komputer manapun tanpa
echo  perlu install Python atau library apapun.
echo ================================================
echo.

:: Tawarkan buka folder dist
set /p buka="Buka folder dist sekarang? (y/n): "
if /i "%buka%"=="y" explorer dist

pause
