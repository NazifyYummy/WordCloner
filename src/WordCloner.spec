# PyInstaller build configuration for a single WordCloner.exe that supports
# both GUI startup (no arguments) and CLI issuance/count/Excel (with -i/-o/-v/-x/--count).
from pathlib import Path

# PyInstaller executes spec files with its own globals; __file__ is not set.
# SPECPATH is supplied by PyInstaller and identifies the spec file's directory.
project_root = Path(SPECPATH or ".").resolve()

# Keep the external resource folders beside WordCloner.exe in the onedir build.
# The Guide directory is optional at build time; if present, both language
# videos are copied with it.
datas = []
for folder_name in ("resources", "Guide"):
    folder_path = project_root / folder_name
    if folder_path.is_dir():
        datas.append((str(folder_path), folder_name))

icon_path = project_root / "app.ico"
app_icon = str(icon_path) if icon_path.is_file() else None
if icon_path.is_file():
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
    name="wordcloner",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    # Build as a console app so CMD waits and CLI output stays attached.
    # Hide the self-owned console on GUI/double-click startup; when launched
    # from an existing CMD window, PyInstaller keeps that console visible.
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
