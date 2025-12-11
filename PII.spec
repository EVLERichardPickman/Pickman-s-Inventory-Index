# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['PII.py'],
    pathex=[],
    binaries=[],
    datas=[('favicon.ico', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # GUI/toolkits not used
        'tkinter', 'PySide2', 'PySide6', 'PyQt6', 'wx', 'kivy',
        # Scientific/plotting stacks
        'numpy', 'pandas', 'scipy', 'matplotlib', 'seaborn', 'numexpr', 'sympy',
        # ML / big libs
        'torch', 'tensorflow', 'sklearn', 'xgboost', 'lightgbm',
        # Web frameworks
        'django', 'flask', 'fastapi', 'starlette', 'uvicorn',
        # Database drivers
        'sqlalchemy', 'psycopg2', 'mysqlclient', 'pymongo', 'redis',
        # Packaging/testing/dev
        'pytest', 'pip', 'setuptools', 'wheel', 'ipython', 'notebook', 'jupyter',
        # Unused Qt bindings/plugins
        'PyQt5.QtWebEngine', 'PyQt5.QtWebEngineCore', 'PyQt5.QtWebEngineWidgets', 'PyQt5.QtTest',
        # Standard library modules unlikely needed at runtime
        'unittest', 'test', 'distutils', 'pydoc',
        'xmlrpc', 'wsgiref', 'cgi', 'cgitb',
        'lib2to3', 'turtledemo', 'multiprocessing',
    ],
    noarchive=False,
    optimize=2,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles if hasattr(a, 'zipfiles') else [],
    a.datas,
    [],
    name="Pickman's Inventory Index",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['favicon.ico'],
)
