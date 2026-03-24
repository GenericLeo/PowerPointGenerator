# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file for Windows build

from pathlib import Path

_version_ns = {}
_version_file = Path.cwd() / 'version.py'
exec(_version_file.read_text(encoding='utf-8'), _version_ns)

__version__ = _version_ns.get('__version__', '1.0.0')
__app_name__ = _version_ns.get('__app_name__', 'PowerPoint Generator')

block_cipher = None

a = Analysis(
    ['gui_app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('version.py', '.'),
    ],
    hiddenimports=[
        'PIL._tkinter_finder',
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
        'tkinter.simpledialog',
        'pptx',
        'pptx.util',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'packaging',
        'packaging.version',
        'certifi',
        'update_manager',
        'version',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PowerPointGenerator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Set to False for GUI app (no console window)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add .ico file path here if you have one
)
