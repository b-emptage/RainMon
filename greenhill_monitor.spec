# -*- mode: python ; coding: utf-8 -*-
"""Frozen build of the Greenhill weather window.

    pyinstaller greenhill_monitor.spec

Windowed, so no console appears behind it. The site map is bundled; the .ini
is NOT -- it must sit beside the executable where an operator can edit it, and
`config_directory()` looks there.

Build it with the same pinned 3.8 32-bit toolchain as the rain bridge if the
window is to run on the Windows 7 box. On a modern machine any current Python
will do; the code is kept 3.8-compatible so one source tree serves both.
"""

a = Analysis(
    ['greenhill_monitor.py'],
    pathex=[],
    binaries=[],
    datas=[('BT_SiteVectorMap.png', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='GreenhillMonitor',
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
    icon=['RainMon.ico'],
)
