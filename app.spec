# app.spec
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

from PyInstaller.utils.hooks import copy_metadata, collect_data_files

# --- Metadata and data collection ---
datas = []
datas += copy_metadata('streamlit')
datas += copy_metadata('altair')
datas += copy_metadata('validators')
datas += copy_metadata('pandas')
datas += copy_metadata('numpy')
datas += collect_data_files('streamlit')   # ensure Streamlit templates/configs are bundled

# --- Analysis step ---
a = Analysis(
    ['run_app.py'],   # entry point file
    pathex=['.'],     # search path (repo root)
    binaries=[],
    datas=datas,
    hiddenimports=[
        'streamlit.runtime.scriptrunner.script_runner',
        'blinker'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# --- Python archive ---
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# --- Executable build ---
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='DocketPoint',       # final EXE name (dist/DocketPoint.exe)
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,            # set to True if you want console logs visible
)
