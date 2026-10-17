from pathlib import Path

project_root = Path(SPECPATH or ".").resolve()

datas = []
for folder_name in ("resources", "Guide"):
    folder_path = project_root / folder_name
    if folder_path.is_dir():
        datas.append((str(folder_path), folder_name))

icon_path = project_root / "app.ico"
if not icon_path.is_file():
    raise FileNotFoundError(f"Required application icon not found: {icon_path}")
app_icon = str(icon_path)
datas.append((str(icon_path), "."))

a = Analysis(
    [str(project_root / "WordCloner.py")],
    pathex=[str(project_root)],
    binaries=[],
    datas=datas,
    hiddenimports=[],
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
    name="WordCloner",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    hide_console="hide-late",
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=app_icon,
    contents_directory=".",
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="WordCloner",
)
