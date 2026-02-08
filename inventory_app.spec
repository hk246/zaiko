# -*- mode: python ; coding: utf-8 -*-

# 在庫管理システム PyInstaller設定ファイル
# このファイルをカスタマイズしてビルドする場合:
#   pyinstaller inventory_app.spec

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('static', 'static'),
    ],
    hiddenimports=[
        'flask',
        'flask_sqlalchemy',
        'flask_wtf',
        'wtforms',
        'wtforms.validators',
        'email.mime.text',
        'email.mime.multipart',
        'sqlalchemy.sql.default_comparator',
        'tkinter',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'webbrowser',
        'threading',
        'shutil',
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
    [],
    exclude_binaries=True,  # ONEDIRモード用
    name='在庫管理システム',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # Falseにするとコンソールウィンドウが表示されない
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # アイコンファイルがあれば 'icon.ico' などを指定
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='在庫管理システム',
)
