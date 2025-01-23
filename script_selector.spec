# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['script_selector.py'],
    pathex=[],
    binaries=[],
    datas=[('nfs_share_script.py', '.'), ('cifs_share_script.py', '.'), ('readData.py', '.'), ('Documents\\CIFSShares.xlsx', 'Documents'), ('Documents\\NFSShares.xlsx', 'Documents')],
    hiddenimports=[],
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
    name='script_selector',
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
