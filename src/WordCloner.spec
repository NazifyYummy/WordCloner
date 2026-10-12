# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['WordCloner.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('resources', 'resources'),       
    ],
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
    a.binaries,
    a.datas,
    [],
    name='WordCloner',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,                 
    icon='app.ico',                  
)