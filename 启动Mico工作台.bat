@echo off
chcp 65001 >nul
title Mico 选品工作台
cd /d "%~dp0"

echo ================================================
echo   Mico 选品工作台
echo ================================================
echo.

set PY=D:\DSH\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe
if not exist "%PY%" set PY=python

echo 正在启动服务...
start "" http://localhost:8000
"%PY%" app.py

pause
