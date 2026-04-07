# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for LIS-ANALYSIS GUI builds on Linux and Windows."""

from pathlib import Path

block_cipher = None
root_dir = Path.cwd()

analysis = Analysis(
    ['main.py'],
    pathex=[str(root_dir)],
    binaries=[],
    datas=[
        ('ACP', 'ACP'),
        ('Listas ATP', 'Listas ATP'),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={
        # Mantem apenas backends realmente usados pela aplicacao.
        'matplotlib': {
            'backends': ['TkAgg', 'Agg'],
        },
    },
    runtime_hooks=[],
    excludes=[
        # Toolkits alternativos de GUI não usados.
        'PyQt5',
        'PyQt6',
        'PySide2',
        'PySide6',
        'wx',
        'gi',
        'PyGObject',
        # Pacotes de notebook/experimentos não usados em runtime.
        'IPython',
        'jupyter',
        'notebook',
        'pytest',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(analysis.pure, analysis.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    analysis.scripts,
    [],
    exclude_binaries=True,
    name='LIS-ANALYSIS',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    analysis.binaries,
    analysis.zipfiles,
    analysis.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='LIS-ANALYSIS',
)
