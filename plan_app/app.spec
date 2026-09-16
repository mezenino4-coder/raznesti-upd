# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec: собирает приложение + модели OCR + onnxruntime в один .exe.
# Запускать на Windows:  pyinstaller app.spec
import os
from PyInstaller.utils.hooks import collect_all

datas, binaries, hiddenimports = [], [], []
for pkg in ("onnxruntime", "rapidocr", "thefuzz", "cv2", "PIL"):
    try:
        d, b, h = collect_all(pkg)
        datas += d
        binaries += b
        hiddenimports += h
    except Exception:
        pass

# модели OCR (папка ../ocr_rus/models)
models_dir = os.path.abspath(os.path.join(SPECPATH, "..", "ocr_rus", "models"))
if os.path.isdir(models_dir):
    datas.append((models_dir, "ocr_rus/models"))

a = Analysis(
    ["gui_tk.py"],
    pathex=[SPECPATH],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=["matplotlib", "PyQt5", "PySide6", "pytest"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="РазнестиУПД",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)
