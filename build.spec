# PyInstaller spec – optimised for smallest binary
# Usage:  pyinstaller build.spec

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('certgen.ico',  '.'),
        ('certgen.png',  '.'),
        ('fonts',        'fonts'),
    ],
    hiddenimports=['PIL', 'openpyxl', 'fpdf'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Only exclude things that are genuinely unreachable from this app.
    # Do NOT exclude email/html/http/urllib — fpdf2 uses them internally.
    excludes=[
        'matplotlib', 'numpy', 'pandas', 'scipy',
        'PyQt5', 'PyQt6', 'PySide2', 'PySide6',
        'wx', 'gi', 'gtk',
        'IPython', 'notebook', 'jupyter',
        'docutils', 'sphinx',
        'cryptography',
        'unittest', 'test',
        'tkinter.test',
        'xmlrpc',
        'pydoc',
        'distutils',
        'lib2to3',
        'colorsys',
        'ttkthemes',
    ],
    noarchive=False,
    optimize=2,
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
    strip=True,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='certgen.ico',
    onefile=True,
)
