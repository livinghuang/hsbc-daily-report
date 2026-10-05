# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包設定：單一 exe、有圖示、有版本資訊、不開黑視窗。"""

block_cipher = None

a = Analysis(
    ['hsbc_app.py'],
    pathex=[],
    binaries=[],
    datas=[('app.ico', '.')],
    hiddenimports=['hsbc_core', 'hsbc_config'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # 這些是 yfinance/pandas 拖進來但這支程式用不到的，拿掉可以讓 exe 小很多
    excludes=['matplotlib', 'scipy', 'PIL', 'PyQt5', 'PySide2', 'notebook', 'IPython'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='HSBC Daily Report',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # 視窗程式，排程執行時不會閃黑視窗
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app.ico',
    version='version_info.txt',
)
