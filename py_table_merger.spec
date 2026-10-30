# -*- mode: python ; coding: utf-8 -*-
import importlib.util
import sys

# Schutz vor einer kaputten EXE: Wird die Spec mit einem Python ohne die Projekt-Abhängigkeiten
# gebaut (z. B. globales Python statt .venv), fehlen sie in der EXE ("No module named 'pandas'").
REQUIRED_MODULES = ["pandas", "xlsxwriter", "openpyxl"]
_missing = [m for m in REQUIRED_MODULES if importlib.util.find_spec(m) is None]
if _missing:
    raise SystemExit(
        "\n\nBUILD ABGEBROCHEN: Im verwendeten Python fehlen " + ", ".join(_missing) + ".\n"
        "Verwendetes Python: " + sys.executable + "\n\n"
        "Bitte im Projektordner so bauen:\n"
        "    uv run --with pyinstaller pyinstaller py_table_merger.spec --clean\n"
    )


a = Analysis(
    ['src\\table_merger\\py_table_merger.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['xlsxwriter', 'openpyxl'],
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
    name='py_table_merger',
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
    icon=['TableMerger_Logo.ico'],
)
