# -*- mode: python ; coding: utf-8 -*-
import os
import glob
import sys

compiled_binaries = []
for pyd_path in glob.glob('*.pyd'):
    compiled_binaries.append((pyd_path, '.'))

win7_compat_dll = 'redist/api-ms-win-core-path-l1-1-0.dll'
if os.path.exists(win7_compat_dll):
    compiled_binaries.append((win7_compat_dll, '.'))

entry_script = 'launcher.py' if os.path.exists('launcher.py') else 'main.py'

a = Analysis(
    [entry_script],
    pathex=[],
    binaries=compiled_binaries,
    datas=[
        ('assets', 'assets'),
        ('templates', 'templates')
    ],
    hiddenimports=[
        'openpyxl',
        'bs4',
        'requests',
        'lxml',
        'docx',
        'main'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe_icon = ['assets/mcomaker.ico'] if sys.platform == 'win32' else None

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='MCOMaker',
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
    icon=exe_icon,
)
