@echo off
title Instalasi Log Detail & QR Generator
echo ============================================
echo  Instalasi Library Python yang Dibutuhkan
echo ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python tidak ditemukan! Silakan install Python 3.10+ terlebih dahulu.
    echo         Download: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/2] Menginstall library...
pip install pandas openpyxl "qrcode[pil]" Pillow PyQt6

echo.
echo [2/2] Verifikasi instalasi...
python -c "import pandas, openpyxl, qrcode, PIL, PyQt6; print('[OK] Semua library berhasil terinstall!')"

echo.
echo ============================================
echo  Instalasi selesai! Jalankan run.bat
echo ============================================
pause
