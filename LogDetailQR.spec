# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('app.py', '.'), ('icon.png', '.')]
binaries = []
hiddenimports = ['pandas', 'openpyxl', 'qrcode', 'PIL', 'PyQt6', 'PyQt6.QtPrintSupport', 'qrcode.image.pil']
tmp_ret = collect_all('qrcode')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('PIL')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('openpyxl')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Scientific / numeric libs tidak dipakai
        'scipy', 'matplotlib', 'pyarrow', 'numba', 'skimage',
        # Database libs tidak dipakai
        'sqlalchemy', 'psycopg2', 'psycopg2_binary', 'pymysql', 'MySQLdb',
        # Security / network libs tidak dipakai
        'bcrypt', 'cryptography', 'urllib3', 'requests', 'certifi',
        'charset_normalizer',
        # Template / markup libs tidak dipakai
        'lxml', 'jinja2', 'fsspec',
        # GUI toolkit lain
        'tkinter', '_tkinter', 'wx',
        # PyQt6 modul tidak dipakai
        'PyQt6.QtNetwork', 'PyQt6.QtBluetooth', 'PyQt6.QtDBus',
        'PyQt6.QtDesigner', 'PyQt6.QtHelp', 'PyQt6.QtMultimedia',
        'PyQt6.QtMultimediaWidgets', 'PyQt6.QtOpenGL', 'PyQt6.QtOpenGLWidgets',
        'PyQt6.QtQml', 'PyQt6.QtQuick', 'PyQt6.QtQuickWidgets',
        'PyQt6.QtSql', 'PyQt6.QtSvg', 'PyQt6.QtSvgWidgets',
        'PyQt6.QtTest', 'PyQt6.QtWebEngine', 'PyQt6.QtWebEngineCore',
        'PyQt6.QtWebEngineWidgets', 'PyQt6.QtXml', 'PyQt6.QtLocation',
        'PyQt6.QtPositioning', 'PyQt6.QtSensors', 'PyQt6.QtSerialPort',
        # Misc tidak dipakai
        'win32com', 'pywintypes', 'pythoncom',
    ],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='LogDetailQR',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)
