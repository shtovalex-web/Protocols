# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller: ProtocolOOT (Linux, один бинарник, tkinter)."""

from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

linux_port = Path(SPECPATH).resolve().parent
app = linux_port / "app"
next_dir = app / "ProtocolOHT_next"
bundle = app / "bundle"

datas = []
icon_ico = bundle / "icon.ico"
if not icon_ico.is_file():
    icon_ico = app / "icon.ico"
if icon_ico.is_file():
    datas.append((str(icon_ico), "."))

hiddenimports = [
    "app_paths",
    "program_keys",
    "clipboard_ui",
    "commission_admin",
    "employees_io",
    "excel_data_cache",
    "docx_template_protection",
    "programs_v_prof",
    "v_prof_combinations",
    "faq_viewer",
    "mintrud_export",
    "mintrud_trained_registry",
    "v_program_registry_match",
    "russian_genitive",
    "fpdf",
    "protocol_db",
    "protocol_errors",
    "protocol_paths",
    "protocol_journal",
    "protocol_docx",
    "protocol_output",
    "protocol_recovery",
    "protocol_app_info",
    "protocol_ui",
    "protocol_embedded_assets",
    "docx",
    "docx.oxml",
    "openpyxl",
    "pymorphy2",
    "pymorphy2_dicts_ru",
]
hiddenimports += collect_submodules("openpyxl")
hiddenimports += collect_submodules("pymorphy2")
datas += collect_data_files("pymorphy2_dicts_ru")

a = Analysis(
    [str(app / "main.py")],
    pathex=[str(next_dir), str(app)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pytest", "pre_commit"],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="ProtocolOOT",
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
    icon=str(icon_ico) if icon_ico.is_file() else None,
)
