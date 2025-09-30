# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

project_dir = Path.cwd()
block_cipher = None

# Collect PyQt5 plugins (platforms, styles, etc.)
pyqt5_plugins = collect_data_files("PyQt5")

# Collect openpyxl templates/styles
openpyxl_data = collect_data_files("openpyxl")

a = Analysis(
    ['CycleAnalyzer.py'],
    pathex=[str(project_dir)],
    binaries=[],
    datas=[
        # App resources
        (str(project_dir / 'diagram.ico'), '.'),
        (str(project_dir / 'siemens.png'), '.'),
        (str(project_dir / 'lines.txt'), '.'),
        (str(project_dir / 'password.txt'), '.'),
        (str(project_dir / 'template.accdb'), '.'),
        (str(project_dir / 'AccessDatabaseEngine_x64.exe'), '.'),
        (str(project_dir / 'AccessDatabaseEngine_x86.exe'), '.'),
    ] + pyqt5_plugins + openpyxl_data,
    hiddenimports=[
        # Stdlib used
        'sys', 'os', 'glob', 're', 'datetime', 'shutil',
        'subprocess', 'platform', 'threading', 'time',
        'pathlib', 'statistics',

        # PyQt5
        'PyQt5',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        'PyQt5.QtPrintSupport',

        # openpyxl
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.workbook',
        'openpyxl.worksheet',

        # ODBC
        'pyodbc',

        # Windows registry
        'winreg',
    ] + collect_submodules('PyQt5'),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    a.datas,
    [],
    name='CycleAnalyzer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,   # GUI mode
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=[str(project_dir / 'diagram.ico')],
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    a.pure,
    a.zipped_data,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='CycleAnalyzer'
)
