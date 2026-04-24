# -*- mode: python ; coding: utf-8 -*-
"""
BookAutoSystem PyInstaller spec file

ビルド方法:
  pyinstaller bookautosystem.spec

出力先: dist/BookAutoSystem/
"""

import os
import sys
from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_submodules

block_cipher = None

# Playwright のデータファイル（ドライバー）を収集
pw_datas, pw_binaries, pw_hiddenimports = collect_all('playwright')

# Windows専用COMモジュール（Linux/macOSでは含めない）
win_imports = (
    ['win32com', 'win32com.client', 'pythoncom', 'pywintypes']
    if sys.platform == 'win32' else []
)

# その他の隠しインポート
hidden_imports = [
    *win_imports,
    'yaml',
    'lxml',
    'lxml.etree',
    'lxml._elementpath',
    'bs4',
    'openpyxl',
    'flask',
    'flask.json',
    'jinja2',
    'markupsafe',
    'werkzeug',
    'sqlite3',
    *pw_hiddenimports,
    *collect_submodules('src'),
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=pw_binaries,
    datas=[
        # Flask テンプレート
        ('src/web/templates', 'src/web/templates'),
        # Playwright データ
        *pw_datas,
    ],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter', '_tkinter', 'matplotlib', 'numpy', 'pandas',
        'scipy', 'PIL', 'cv2', 'torch', 'tensorflow',
    ],
    noarchive=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='BookAutoSystem',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='BookAutoSystem',
)
