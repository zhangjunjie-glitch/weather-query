@echo off
chcp 65001 >nul
echo ========================================
echo   天气数据查询 - 单文件 exe 打包
echo ========================================
echo   生成单个 dist\WeatherQuery.exe，便于拷贝，但启动时需解压会稍慢。
echo.
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo 正在安装 PyInstaller...
    pip install pyinstaller -q
)
echo 正在打包...
pyinstaller --noconfirm --clean WeatherQuery.spec
if errorlevel 1 ( echo 打包失败。 & pause & exit /b 1 )
echo.
echo 完成。可执行文件: dist\WeatherQuery.exe
pause
