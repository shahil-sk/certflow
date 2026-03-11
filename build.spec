# PyInstaller spec – optimised for smallest binary
# Usage:  pyinstaller build.spec

import os
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

a = Analysis(
    ['certgen.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('certgen.ico',  '.'),
        ('certgen.png',  '.'),
        ('fonts',        'fonts'),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Exclude heavy, unused packages to shrink binary
    excludes=[
        'matplotlib', 'numpy', 'pandas', 'scipy',
        'PyQt5', 'PyQt6', 'PySide2', 'PySide6',
        'wx', 'gi', 'gtk',
        'IPython', 'notebook', 'jupyter',
        'docutils', 'sphinx',
        'cryptography', 'ssl',
        'email', 'html', 'http', 'urllib',
        'unittest', 'test',
        'tkinter.test',
        'colorsys',   # not used any more
        'ttkthemes',  # dropped; using built-in ttk
    ],
    noarchive=False,
    optimize=2,   # strip docstrings + asserts
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='CertWizard',
    debug=False,
    bootloader_ignore_signals=False,
    strip=True,          # strip debug symbols
    upx=True,            # compress with UPX if available
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,       # no console window
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='certgen.ico',
    onefile=True,        # single-file portable exe
)
