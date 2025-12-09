@echo off
pushd "%~dp0"
echo 正在启动燃料管理系统...
python gui_app.py
if %errorlevel% neq 0 (
    echo 程序启动失败，请检查是否安装了Python以及pandas库。
    pause
)
popd
