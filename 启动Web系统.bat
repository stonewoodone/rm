@echo off
chcp 65001
pushd "%~dp0"
title 燃料管理Web系统
echo ==========================================
echo       正在启动燃料管理Web系统...
echo ==========================================
echo.
echo 系统启动成功后，请在浏览器中访问:
echo http://127.0.0.1:5000
echo.
echo (请不要关闭此窗口)
echo.
python app.py
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] 程序启动失败！
    echo 请检查是否安装了 Flask: pip install flask
    pause
)
pause
