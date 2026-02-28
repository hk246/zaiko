# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for 在庫管理システム (onedir distribution)."""

import os

ROOT = os.path.dirname(os.path.abspath(SPEC))

a = Analysis(
    [os.path.join(ROOT, 'run.py')],
    pathex=[ROOT],
    binaries=[],
    datas=[
        # Frontend build output
        (os.path.join(ROOT, 'frontend', 'dist'), os.path.join('frontend', 'dist')),
    ],
    hiddenimports=[
        # FastAPI / Starlette / Uvicorn
        'fastapi',
        'starlette',
        'starlette.middleware',
        'starlette.middleware.cors',
        'starlette.responses',
        'starlette.staticfiles',
        'starlette.routing',
        'uvicorn',
        'uvicorn.config',
        'uvicorn.main',
        'uvicorn.protocols',
        'uvicorn.protocols.http',
        'uvicorn.protocols.http.auto',
        'uvicorn.protocols.http.h11_impl',
        'uvicorn.protocols.http.httptools_impl',
        'uvicorn.protocols.websockets',
        'uvicorn.protocols.websockets.auto',
        'uvicorn.lifespan',
        'uvicorn.lifespan.on',
        'uvicorn.lifespan.off',
        'uvicorn.loops',
        'uvicorn.loops.auto',
        'uvicorn.loops.asyncio',
        # SQLAlchemy
        'sqlalchemy',
        'sqlalchemy.orm',
        'sqlalchemy.pool',
        'sqlalchemy.dialects.sqlite',
        'sqlalchemy.dialects.sqlite.pysqlite',
        # Pydantic
        'pydantic',
        'pydantic_settings',
        'pydantic.deprecated',
        'pydantic.deprecated.decorator',
        # Multipart (file upload)
        'multipart',
        'python_multipart',
        # Excel
        'openpyxl',
        'et_xmlfile',
        # Win32 COM (Outlook email)
        'win32com',
        'win32com.client',
        # Backend modules
        'backend',
        'backend.main',
        'backend.config',
        'backend.database',
        'backend.models',
        'backend.models.material',
        'backend.models.lot',
        'backend.models.label',
        'backend.models.master',
        'backend.models.property_field',
        'backend.models.reservation',
        'backend.models.recipe',
        'backend.models.work_order',
        'backend.models.stock_transaction',
        'backend.routers.materials',
        'backend.routers.lots',
        'backend.routers.reservations',
        'backend.routers.recipes',
        'backend.routers.work_orders',
        'backend.routers.labels',
        'backend.routers.properties',
        'backend.routers.master',
        'backend.routers.dashboard',
        'backend.routers.backup',
        'backend.routers.settings',
        'backend.routers.import_router',
        'backend.routers.email_router',
        'backend.routers.export_router',
        'backend.services.stock_calculator',
        'backend.services.import_service',
        'backend.services.export_service',
        'backend.services.reimport_service',
        'backend.services.backup_service',
        'backend.services.email_service',
        # Standard lib used at runtime
        'email.mime.text',
        'email.mime.multipart',
        'sqlite3',
        'json',
        'csv',
        'io',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'PIL',
        'pytest',
        'setuptools',
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
    name='在庫管理システム',
    debug=True,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='在庫管理システム',
)
