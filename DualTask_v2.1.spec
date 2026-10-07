# -*- mode: python ; coding: utf-8 -*-
# NOTE: openpyxl's lxml/PIL/numpy support is optional and falls back to the
# stdlib automatically when those packages aren't importable (dual_task.py
# only does basic Workbook/Font/cell writes, never charts or embedded images).
# Excluding them — along with unrelated dev-machine packages that previous
# specs accidentally pulled in via collect_submodules('openpyxl') — keeps the
# build to a normal size instead of ballooning to 300+ MB.
a = Analysis(
    ['dual_task.py'],
    pathex=[],
    binaries=[],
    datas=[('images\\neuro_logo.ico', 'images')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'lxml', 'PIL', 'numpy', 'pandas', 'matplotlib', 'scipy',
        'PySide6', 'PySide2', 'PyQt5', 'PyQt6',
        'pygame', 'psutil', 'yaml',
        'IPython', 'jupyter', 'notebook', 'zmq', 'tornado',
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='DualTask_v2.1',
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
    icon=['images\\neuro_logo.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='DualTask_v2.1',
)
