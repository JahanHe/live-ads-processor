# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules
from pathlib import Path
import shutil


APP_NAME = "直播投放处理器后端"

datas = [("static", "static")]
tesseract = shutil.which("tesseract")
if tesseract:
    datas.append((tesseract, "ocr"))
for tessdata_dir in (Path("/opt/homebrew/share/tessdata"), Path("/usr/local/share/tessdata")):
    if tessdata_dir.exists():
        for lang in ("chi_sim.traineddata", "eng.traineddata", "osd.traineddata"):
            traineddata = tessdata_dir / lang
            if traineddata.exists():
                datas.append((str(traineddata), "ocr/tessdata"))
        break


a = Analysis(
    ["web_app.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=collect_submodules("PIL"),
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
    [],
    exclude_binaries=True,
    name=APP_NAME,
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

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name=APP_NAME,
)
