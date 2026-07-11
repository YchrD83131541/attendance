# -*- mode: python ; coding: utf-8 -*-
# ビルド: pyinstaller attendance_app.spec --clean --noconfirm

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# openpyxl はテンプレートファイルを内包するため collect_data_files が必要
openpyxl_datas = collect_data_files("openpyxl")

a = Analysis(
    ["attendance_gui.py"],
    pathex=[],
    binaries=[],
    datas=openpyxl_datas,
    hiddenimports=[
        "openpyxl",
        "openpyxl.cell._writer",
        "openpyxl.styles",
        "openpyxl.utils",
        "et_xmlfile",
        "csv",
        "datetime",
        "tkinter",
        "tkinter.ttk",
        "tkinter.messagebox",
        "tkinter.filedialog",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "matplotlib", "numpy", "pandas", "scipy",
        "PyQt5", "PyQt6", "PySide2", "PySide6",
        "notebook", "jupyter", "unittest",
    ],
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
    name="勤怠管理",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
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
    name="勤怠管理",
)

