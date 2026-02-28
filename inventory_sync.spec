# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file for Inventory Sync with bundled Ghostscript

import os
from pathlib import Path

# Find Ghostscript installation
gs_source = None
gs_paths = [
    r"C:\Program Files\gs\gs10.06.0",
    r"C:\Program Files\gs\gs10.01.2",
    r"C:\Program Files\gs\gs10.0.0",
]

for path in gs_paths:
    if os.path.exists(path):
        gs_source = path
        break

if not gs_source:
    print("WARNING: Ghostscript not found! The exe will need Ghostscript installed separately.")
    gs_datas = []
else:
    print(f"Found Ghostscript at: {gs_source}")
    # Include the bin and lib folders from Ghostscript
    gs_datas = [
        (os.path.join(gs_source, 'bin'), 'gs/bin'),
        (os.path.join(gs_source, 'lib'), 'gs/lib'),
    ]

a = Analysis(
    ['inventory_sync.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('CASCADELOGO.png', '.'),  # Include logo if it exists
        ('fedex_shipping.py', '.'),  # Include FedEx shipping module
        ('auto_updater.py', '.'),  # Include auto-updater module
    ] + gs_datas,
    hiddenimports=[
        'PIL._tkinter_finder',
        'sv_ttk',
        'requests',
        'fedex_shipping',
        'auto_updater',
        'pandas',
        'pandas._libs',
        'pandas._libs.tslibs',
        'openpyxl',
        'supabase',
        'postgrest',
        'gotrue',
        'realtime',
        'storage3',
        'supafunc',
        'httpx',
        'httpcore',
        'reportlab',
        'pystray',
        'win32print',
        'win32api',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Exclude packages not needed by this app (saves ~2GB)
        'torch', 'torchvision', 'torchaudio',
        'tensorflow', 'tensorboard', 'keras',
        'scipy',
        'cv2', 'opencv',
        'transformers',
        'matplotlib',
        'IPython', 'jupyter', 'notebook',
        'pytest',
        'sphinx',
        'docutils',
    ],
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
    name='InventorySync',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add icon path here if you have one: icon='app.ico'
)
