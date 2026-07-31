# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files

rapidocr_datas = collect_data_files("rapidocr", includes=["*.yaml", "**/*.yaml"])

a = Analysis(
    ["mcp_server.py"],
    pathex=[], binaries=[],
    datas=[("mcp_policy.json", "."), ("ocr_models", "ocr_models")] + rapidocr_datas,
    hiddenimports=["rapidocr", "rapidocr.main", "rapidocr.inference_engine.onnxruntime", "onnxruntime", "pythoncom", "win32com.client"],
    hookspath=[], hooksconfig={}, runtime_hooks=[],
    excludes=["onnxruntime.transformers", "pandas"],
    noarchive=False, optimize=0,
)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name="DnS Auto MCP", debug=False, bootloader_ignore_signals=False, strip=False, upx=False, console=True)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name="DnS Auto")