# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files

rapidocr_datas = collect_data_files(
    'rapidocr',
    includes=['*.yaml', '**/*.yaml'],
)

a = Analysis(
    ['DnS_Auto_Main.py'],
    pathex=[],
    binaries=[],
    datas=[('ocr_models', 'ocr_models')] + rapidocr_datas,
    hiddenimports=[
        'rapidocr',
        'rapidocr.main',
        'rapidocr.inference_engine.onnxruntime',
        'onnxruntime',
        'pythoncom',
        'win32com.client',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # OCR 추론에 필요하지 않은 ONNX Transformer 도구와 pandas는 패키징에서 제외합니다.
    excludes=['onnxruntime.transformers', 'pandas'],
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
    name='DnS Auto',
    version='version_info.txt',
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
