# -*- mode: python ; coding: utf-8 -*-

from version import __version__, __bundle_id__, __app_name__

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
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=True,  # Enable for macOS drag & drop
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

app = BUNDLE(
    exe,
    name='PowerPointGenerator.app',
    icon=None,
    bundle_identifier=__bundle_id__,
    info_plist={
        'CFBundleDisplayName': __app_name__,
        'CFBundleShortVersionString': __version__,
        'CFBundleVersion': __version__,
        'NSHighResolutionCapable': 'True',
        'LSBackgroundOnly': 'False',
        'CFBundleDocumentTypes': [
            {
                'CFBundleTypeName': 'Image Files',
                'CFBundleTypeRole': 'Viewer',
                'LSItemContentTypes': [
                    'public.image',
                    'public.tiff',
                    'public.jpeg',
                    'public.png',
                ],
            }
        ],
        'NSRequiresAquaSystemAppearance': 'False',
    },
)
