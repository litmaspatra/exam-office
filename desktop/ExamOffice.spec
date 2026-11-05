# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None
spec_dir = os.path.dirname(os.path.abspath(SPEC))

# Collect data files for python-docx (templates, default.docx) and lxml
datas = []
datas += collect_data_files('docx')
datas += collect_data_files('lxml')

papers_path = os.path.join(spec_dir, 'papers')
if os.path.exists(papers_path):
    datas.append((papers_path, 'papers'))

icon_path = os.path.join(spec_dir, 'icon.ico')
if os.path.exists(icon_path):
    datas.append((icon_path, '.'))

# Collect all hidden imports
hiddenimports = [
    'docx',
    'docx.oxml',
    'docx.oxml.ns',
    'lxml',
    'lxml.etree',
    'lxml._elementpath',
    'kruti_converter',
    'paper_loader',
    'docx_gen',
    'marker',
    'store',
    'win32gui',
    'win32con',
    'win32process',
    'win32com',
    'win32com.client',
]
hiddenimports += collect_submodules('docx')
hiddenimports += collect_submodules('lxml')
hiddenimports = sorted(list(set(hiddenimports)))

a = Analysis(
    [os.path.join(spec_dir, 'exam_office.py')],
    pathex=[spec_dir],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pytest', 'unittest'],
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
    name='ExamOffice',
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
    icon=icon_path if os.path.exists(icon_path) else None,
)
