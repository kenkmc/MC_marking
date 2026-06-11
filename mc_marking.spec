# -*- mode: python ; coding: utf-8 -*-

import os

block_cipher = None

icon_path = os.path.abspath("app.ico")
icon = icon_path if os.path.exists(icon_path) else None

hiddenimports = []
datas = []
binaries = []
excludes = [
    "PyQt6",
    "PyQt6.QtCore",
    "PyQt6.QtGui",
    "PyQt6.QtWidgets",
    "paddle",
    "paddleocr",
    "paddlex",
    "tensorboard",
    "tensorflow",
    "tkinter",
    "_tkinter",
    "modelscope",
]


def collect_tree(src_dir, dest_root):
    collected = []
    if not os.path.isdir(src_dir):
        return collected

    for root, _, files in os.walk(src_dir):
        rel_root = os.path.relpath(root, src_dir)
        dest_dir = dest_root if rel_root == "." else os.path.join(dest_root, rel_root)
        for file_name in files:
            collected.append((os.path.join(root, file_name), dest_dir))
    return collected

# Include sample template (optional)
datas += [("template.json", ".")]
datas += collect_tree("template", "template")


a = Analysis(
    ["main.py"],
    pathex=[os.path.abspath(".")],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="CheckMate",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI app
    icon=icon,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="CheckMate",
)
