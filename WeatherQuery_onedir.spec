# -*- mode: python ; coding: utf-8 -*-
# 目录版打包：不压成单文件，启动时无需解压，启动更快
# 生成 dist\WeatherQuery\ 文件夹，运行其中的 WeatherQuery.exe

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
    [],
    exclude_binaries=True,
    name='WeatherQuery',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='WeatherQuery',
)
