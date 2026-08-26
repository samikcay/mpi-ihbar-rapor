# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller yapilandirmasi - TEK DOSYA (onefile) exe uretir.

Derlemek icin:
    pyinstaller --clean --noconfirm ihbar_rapor.spec

Sonuc: dist\\IhbarRapor.exe  (tek dosya, yaninda hicbir sey gerekmez)

Not: tkcalendar, tarih adlarini 'babel' paketinden okur; babel'in yerel
veri dosyalari otomatik toplanmadigi icin elle ekleniyor. Aksi halde
exe acilirken "unknown locale: tr_TR" hatasi verir.
"""

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

datas = []
hiddenimports = []

# babel yerel (locale) verileri - tr_TR takvimi icin sart
try:
    datas += collect_data_files('babel')
    hiddenimports += collect_submodules('babel')
except Exception:
    pass

# tkcalendar
try:
    datas += collect_data_files('tkcalendar')
except Exception:
    pass

# python-docx sablon dosyalari (docx/templates/*.xml)
try:
    datas += collect_data_files('docx')
except Exception:
    pass

# openpyxl / win32com alt modulleri
hiddenimports += [
    'win32com.client',
    'win32timezone',
    'openpyxl',
    'openpyxl.cell._writer',
    'docx',
    'tkcalendar',
    'babel.numbers',
    'babel.dates',
]

a = Analysis(
    ['rapor.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Gereksiz buyuk paketleri disla (exe boyutu kucultur)
    excludes=[
        'matplotlib', 'numpy', 'pandas', 'scipy', 'PIL', 'PyQt5', 'PySide2',
        'IPython', 'jupyter', 'notebook', 'pytest', 'sqlalchemy',
        'torch', 'tensorflow', 'sklearn', 'plotly', 'bokeh',
    ],
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
    name='IhbarRapor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # pencere uygulamasi (konsol acilmasin)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
