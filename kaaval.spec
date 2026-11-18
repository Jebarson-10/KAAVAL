# -*- mode: python ; coding: utf-8 -*-
# KAAVAL — PyInstaller Spec File

import os

block_cipher = None
base_dir = os.path.dirname(os.path.abspath(SPEC))

from PyInstaller.utils.hooks import collect_submodules
chromadb_hidden = collect_submodules('chromadb')

a = Analysis(
    [os.path.join(base_dir, 'kaaval_app.py')],
    pathex=[base_dir, os.path.join(base_dir, 'kaaval')],
    binaries=[],
    datas=[
        (os.path.join(base_dir, 'kaaval', 'kaaval.ico'), 'kaaval'),
        (os.path.join(base_dir, 'kaaval', 'logo.png'), 'kaaval'),
    ],
    hiddenimports=[
        'PySide6.QtCore',
        'PySide6.QtGui',
        'PySide6.QtWidgets',
        'openpyxl',
        'PIL',
        'pymupdf',
    ] + chromadb_hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'scipy', 'numpy.testing'],
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='KAAVAL',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window — pure GUI
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(base_dir, 'kaaval', 'kaaval.ico'),
)
