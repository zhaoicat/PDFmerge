# -*- mode: python ; coding: utf-8 -*-

import os
import glob

# 获取Tesseract安装路径
tesseract_path = r'C:\Program Files\Tesseract-OCR'
tesseract_binaries = []
tesseract_datas = []

if os.path.exists(tesseract_path):
    # 添加tesseract.exe
    tesseract_binaries.append((os.path.join(tesseract_path, 'tesseract.exe'), 'tesseract'))
    
    # 添加所有DLL文件
    dll_files = glob.glob(os.path.join(tesseract_path, '*.dll'))
    for dll in dll_files:
        tesseract_binaries.append((dll, 'tesseract'))
    
    # 添加tessdata目录
    tessdata_path = os.path.join(tesseract_path, 'tessdata')
    if os.path.exists(tessdata_path):
        tesseract_datas.append((tessdata_path, 'tesseract/tessdata'))

a = Analysis(
    ['pdf_interleave_merger.py'],
    pathex=[],
    binaries=tesseract_binaries,
    datas=[
        ('tesseract/tessdata', 'tesseract/tessdata'),
        ('poppler', 'poppler'),
        ('requirements.txt', '.'),
    ] + tesseract_datas,
    hiddenimports=[
        'pytesseract',
        'pdf2image',
        'PIL',
        'PIL._tkinter_finder',
        'PyPDF2',
        'tqdm',
        'pathlib',
        're',
        'os',
        'sys',
        'time',
    ],
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
    name='PDF合并工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
