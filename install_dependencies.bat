@echo off
chcp 65001 >nul
setlocal

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

echo ========================================
echo 安装 SpectraMax 转换脚本依赖
echo ========================================
echo.
echo 将安装 requirements.txt 中列出的 Python 库：
echo   numpy
echo   pandas
echo   openpyxl
echo.

python -m pip install -r "%SCRIPT_DIR%requirements.txt"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if "%EXIT_CODE%"=="0" (
    echo 依赖安装完成。现在可以双击 easy_process.bat 使用。
) else (
    echo 依赖安装失败。请确认 Python 和 pip 可用，或把上面的错误信息发给我。
)
echo.
pause
exit /b %EXIT_CODE%