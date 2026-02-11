# -*- mode: python ; coding: utf-8 -*-
# 精简打包：排除与程序无关的重型库，减小体积、加快启动
# 本应用仅需：tkinter, pandas, openpyxl, weather

block_cipher = None

a = Analysis(
    ['weather_app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['weather', 'openpyxl', 'pandas'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'matplotlib.pyplot', 'matplotlib.backends',
        'scipy', 'numpy.distutils', 'PIL', 'cv2', 'IPython', 'jupyter',
        'notebook', 'pytest', 'sphinx', 'setuptools', 'pip',
        'PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'tkinter.test',
        'pandas.tests',
    ],
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
    name='WeatherQuery',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
