@echo off
chcp 65001 >nul
echo ========================================
echo   天气数据查询应用 - 打包脚本（目录版）
echo ========================================
echo   生成 dist\WeatherQuery\ 文件夹，运行其中的 WeatherQuery.exe 启动更快。
echo.
REM 使用 WeatherQuery_onedir.spec，已排除无关库以减小体积
echo.

REM 检查 PyInstaller
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo [1/2] 正在安装 PyInstaller...
    pip install pyinstaller -q
    if errorlevel 1 (
        echo 安装失败，请手动执行: pip install pyinstaller
        pause
        exit /b 1
    )
) else (
    echo [1/2] PyInstaller 已安装
)

echo [2/2] 正在打包（目录版）...
echo.
pyinstaller --noconfirm --clean WeatherQuery_onedir.spec

if errorlevel 1 (
    echo.
    echo 打包失败。
    pause
    exit /b 1
)

echo.
echo ========================================
echo   打包完成
echo ========================================
echo   输出目录: dist\WeatherQuery\
echo   运行: dist\WeatherQuery\WeatherQuery.exe
echo   分发时请将整个 WeatherQuery 文件夹一并拷贝。
echo   可同时附带 使用说明.md、开发说明.md 等。
echo ========================================
pause
