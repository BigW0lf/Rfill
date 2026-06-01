# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_all, collect_submodules

ROOT  = os.path.abspath(os.path.join(SPECPATH, '..'))
CONDA = r'C:\Users\JulesFAGUET\anaconda3\envs\venv'

# Collecter openpyxl entièrement (sous-modules + datas)
openpyxl_datas, openpyxl_binaries, openpyxl_hiddenimports = collect_all('openpyxl')

a = Analysis(
    [os.path.join(ROOT, 'Main.py')],
    pathex=[ROOT],
    binaries=[
        (os.path.join(CONDA, 'DLLs',           '_ctypes.pyd'),       '.'),
        (os.path.join(CONDA, 'DLLs',           '_tkinter.pyd'),      '.'),
        (os.path.join(CONDA, 'Library', 'bin', 'ffi.dll'),           '.'),
        (os.path.join(CONDA, 'Library', 'bin', 'tcl86t.dll'),        '.'),
        (os.path.join(CONDA, 'Library', 'bin', 'tk86t.dll'),         '.'),
        (os.path.join(CONDA,                   'python313.dll'),      '.'),
        (os.path.join(CONDA, 'Library', 'bin', 'vcruntime140.dll'),  '.'),
    ] + openpyxl_binaries,
    datas=[
        (os.path.join(ROOT, 'glossary_surf.py'),                      '.'),
        (os.path.join(ROOT, 'Rfill.ico'),                             '.'),
        (os.path.join(ROOT, 'Rfill.png'),                             '.'),
        (os.path.join(ROOT, 'img'),                                   'img'),
        (os.path.join(CONDA, 'Library', 'lib', 'tcl8.6'),            'tcl/tcl8.6'),
        (os.path.join(CONDA, 'Library', 'lib', 'tk8.6'),             'tcl/tk8.6'),
    ] + openpyxl_datas,
    hiddenimports=[
        'xlsxwriter',
        'xlrd',
        'pandas',
    ] + openpyxl_hiddenimports,
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
    icon=os.path.join(ROOT, 'Rfill.ico'),
)
