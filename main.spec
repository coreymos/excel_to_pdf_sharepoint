# -*- mode: python ; coding: utf-8 -*-
import sys
from PyInstaller.utils.hooks import collect_submodules

block_cipher = None

# collect every win32com submodule (this will grab win32timezone too)
pywin32_hidden = collect_submodules('win32com') + ['win32timezone']

a = Analysis(
    ['main.py'],
    pathex=['.'],               # look in the current directory
    binaries=[],
    datas=[],
    hiddenimports=pywin32_hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'PyQt5', 'wx', 'PySide',
    ],
    noarchive=False,
    optimize=2,
)

pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher,
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    exclude_binaries=False,
    name='main',
    debug=False,
    strip=False,
    upx=True,
    console=True,
    bootloader_ignore_signals=False,
    disable_windowed_traceback=False,
)

# No COLLECT step needed for one-file builds
