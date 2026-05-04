# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['printer_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\ankit\\AppData\\Local\\Programs\\Python\\Python312\\tcl\\tcl8.6', 'tcl\\tcl8.6'), ('C:\\Users\\ankit\\AppData\\Local\\Programs\\Python\\Python312\\tcl\\tk8.6', 'tcl\\tk8.6')],
    hiddenimports=['tkinter', '_tkinter'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['pyi_rth_tkfix.py'],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='SeznikEONPrinterToolkit',
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
)
