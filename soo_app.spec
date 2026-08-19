# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['soo_app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('vi_template_blank.xlsm', '.'),
        ('soo_interactive.html', '.'),
        ('splash_logo.png', '.'),
        ('OPERATING INSTRUCTIONS', 'OPERATING INSTRUCTIONS'),
    ],
    hiddenimports=[
        'pypdf',
        'fpdf',
        'fpdf.enums',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.scrolledtext',
        'collections',
        'collections.abc',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['win32com'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='DFP Turnover Tool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='DFP Turnover Tool',
)
