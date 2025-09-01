# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=['d:\\Jibal\\New'],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('static', 'static'),
        ('shipping.db', '.'),
    ],
    hiddenimports=['sqlite3', 'flask', 'openpyxl', 'xlsxwriter'],
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
    [],
    exclude_binaries=True,
    name='نظام إدارة الشحنات',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,  # تم تغيير هذه القيمة من False إلى True لعرض نافذة موجه الأوامر
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='static/jibal-logo.png',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='app_console',  # تم تغيير اسم المجلد من 'app' إلى 'app_console'
)
