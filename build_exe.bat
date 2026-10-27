@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================================
echo 正在使用 PyInstaller 打包生成独立单文件 GDriveBatchTool.exe ...
echo ========================================================

set BIN_ARGS=
if exist "C:\Program Files\7-Zip\7z.exe" (
    set BIN_ARGS=--add-binary "C:\Program Files\7-Zip\7z.exe;7z_bin" --add-binary "C:\Program Files\7-Zip\7z.dll;7z_bin"
)

pyinstaller --noconfirm --clean --onefile --windowed ^
    --name "GDriveBatchTool" ^
    --collect-submodules core ^
    --collect-submodules ui ^
    %BIN_ARGS% ^
    main.py

if errorlevel 1 (
    echo.
    echo [错误] 打包失败，请检查是否安装了 pyinstaller。
    pause
    exit /b 1
)

echo.
echo ========================================================
echo [完成] EXE 打包成功！文件位置: dist\GDriveBatchTool.exe
echo ========================================================
pause
