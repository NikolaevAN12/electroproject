# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files

datas = [('D:\\Pythonprogs\\elproject\\_ep_build_tag.txt', '.')]
datas += collect_data_files('docx')


a = Analysis(
    ['D:\\Pythonprogs\\elproject\\app\\main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=['docx', 'docx.oxml', 'openpyxl', 'PIL', 'PIL.Image', 'numpy', 'matplotlib'],
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
    name='Electroproject_v2_9_2_20260606_230050',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
