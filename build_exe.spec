# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

from PyInstaller.utils.hooks import collect_all, collect_submodules

block_cipher = None
root = Path(SPECPATH)

ctk_datas, ctk_binaries, ctk_hidden = collect_all("customtkinter")

datas = [
    (str(root / "app" / "templates" / "informe.docx"), "templates"),
    (str(root / "app" / "templates" / "matriz.xlsx"), "templates"),
    (str(root / "app" / "db" / "schema.sql"), "db"),
] + ctk_datas

hiddenimports = sorted(
    set(
        ctk_hidden
        + collect_submodules("app")
        + collect_submodules("customtkinter")
        + [
            "customtkinter",
            "pandas",
            "openpyxl",
            "docx",
            "dateutil",
            "dateutil.parser",
            "docx.oxml",
            "docx.table",
            "docx.enum.text",
        ]
    )
)

a = Analysis(
    [str(root / "app" / "main.py")],
    pathex=[str(root)],
    binaries=ctk_binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    a.zipfiles,
    a.datas,
    [],
    name="GeneradorInformes",
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
)
