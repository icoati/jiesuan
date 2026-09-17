@echo off
chcp 65001 >nul
title 北检^&华东科普点评电签表一键生成工具
cd /d "%~dp0"
echo ==================================================
echo   北检^&华东科普点评电签表一键生成工具 (Windows版)
echo ==================================================
echo.
python generate_dianping_sign_table.py
echo.
echo ==================================================
pause
