@echo off
setlocal
cd /d "%~dp0.."

:: Step 1: 清理 bin 文件夹
if exist bin (
    rmdir /s /q bin
)

:: Step 2: 打包成单文件 exe，需要指定你的python 解释器路径,这个环境中要装上nuitka
:: 为了生成 GUI 子系统的 exe（不弹出命令行窗口），添加 --windows-disable-console
D:\python3.11\.venv\Scripts\python.exe -m nuitka --onefile --enable-plugin=tk-inter --mingw64 --windows-disable-console --jobs=%NUMBER_OF_PROCESSORS% excel_merge_gui.py

:: Step 3: 清理 Nuitka 构建中间目录
rmdir /s /q __pycache__ >nul 2>nul
for /d %%d in (*.build *.dist *.onefile-build) do rmdir /s /q "%%d" >nul 2>nul

:: Step 4: 创建 bin 目录
mkdir bin

:: Step 5: 移动 exe 文件到 bin
move excel_merge_gui.exe bin\

echo.
echo 工具打包完成！
echo 可执行文件已移动到 bin\ 目录！
pause
