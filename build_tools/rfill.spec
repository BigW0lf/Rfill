# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['Main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('glossary_surf.py', '.'),
        ('Rfill.png', '.'),
    ],
    hiddenimports=[
        'openpyxl',
        'openpyxl.cell._writer',
        'xlsxwriter',
        'xlrd',
        'pandas',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Rfill',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon='Rfill.ico',
)
