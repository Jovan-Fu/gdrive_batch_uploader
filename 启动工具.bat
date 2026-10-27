@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在启动 Google Drive 批量分类上传 & 谷歌表格层级归档工具...
python main.py
if errorlevel 1 (
    echo.
    echo 程序运行出错，请检查 Python 环境及依赖是否安装。
    pause
)
