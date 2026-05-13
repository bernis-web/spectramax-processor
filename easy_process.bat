@echo off
chcp 65001 >nul
setlocal

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

echo ========================================
echo SpectraMax 一键转换
echo ========================================
echo.
echo 用法：
echo   1. 直接双击本文件：打开图形界面选择 .xls/.txt
echo   2. 把一个或多个 .xls/.txt 文件拖到本文件上：自动转换
echo.

python -X utf8 "%SCRIPT_DIR%easy_process.py" %*
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
    echo.
    echo 处理过程中遇到问题。
    echo 如果提示缺少 numpy / pandas / openpyxl，请先双击 install_dependencies.bat 安装依赖。
)

echo.
pause
exit /b %EXIT_CODE%