@echo off
setlocal
cd /d "%~dp0.."

D:\python3.11\.venv\Scripts\python.exe -m PyInstaller --onefile --windowed --name excel_merge_gui excel_merge_gui.py

:: 清理中间目录
rmdir /s /q build >nul 2>nul
del /q excel_merge_gui.spec >nul 2>nul

echo.
echo 工具打包完成！
echo 可执行文件在 dist\ 目录下。
pause
