@echo off
chcp 65001 >nul
title 调研案例电签表 一键生成工具
echo ============================================================
echo   🚀 调研案例电签表 - 一键智能结算核验工具 (Windows)
echo ============================================================
echo.

cd /d "%~dp0"

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo ❌ 未检测到 Python 环境，请先安装 Python 3.8+！
    pause
    exit /b 1
)

python settlement_survey_case.py "%~dp0"

echo.
echo ============================================================
echo 🎉 处理完成！按任意键退出...
pause >nul
