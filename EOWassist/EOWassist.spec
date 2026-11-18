# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

spec_dir = Path(SPEC).resolve().parent
a = Analysis(
    [str(spec_dir / 'Dataext.py')],
    pathex=[str(spec_dir.parent)],
    binaries=[],
    datas=[],
    hiddenimports=['PIL.ImageEnhance', 'PIL.ImageFilter', 'PIL.ImageOps'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
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
    name='EOWassist',
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
