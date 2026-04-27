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
        'xlsxwriter',
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
    [],
    exclude_binaries=True,
    name='Rfill',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,          # pas de fenêtre console
    icon='Rfill.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Rfill',
)
