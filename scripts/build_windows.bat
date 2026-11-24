@echo off
rem ============================================================
rem  FinSummaryTool Windows 客户端构建脚本
rem  在 Windows 机器上双击或命令行运行本脚本（需已安装 Python 3.11+）
rem  产物：dist\FinSummaryTool\FinSummaryTool.exe（整个文件夹拷给用户）
rem        dist\FinSummaryTool-windows-x64.zip（可直接分发的压缩包）
rem ============================================================
cd /d %~dp0..

rem 1) 独立构建虚拟环境（不污染开发环境）
if not exist .venv-build (
    python -m venv .venv-build
)
call .venv-build\Scripts\activate.bat

rem 2) 安装依赖（运行依赖 + 打包工具）
python -m pip install -r requirements.txt -r requirements-build.txt

rem 3) 打包
python -m PyInstaller --noconfirm --clean FinSummaryTool.spec
if errorlevel 1 (
    echo [失败] 打包出错，请检查上方日志
    exit /b 1
)

rem 4) 自检（验证客户端内资源可解析）
dist\FinSummaryTool\FinSummaryTool.exe --selftest
if errorlevel 1 (
    echo [失败] 自检未通过
    exit /b 1
)

rem 4b) Bundle 校验：确认 VC 运行库已打入（未装运行库的电脑也能打开）
if not exist dist\FinSummaryTool\_internal\vcruntime140.dll (
    echo [警告] 未找到 vcruntime140.dll —— 目标电脑可能需要安装 VC++ 运行库
    echo         建议改用已装 VC++ 运行库的机器构建，或让用户安装运行库
) else (
    echo [OK] VC 运行库已打入，目标电脑无需 Python 与 VC++ 运行库
)

rem 5) 压缩分发包
powershell -NoProfile -Command "if (Test-Path dist\FinSummaryTool-windows-x64.zip) { Remove-Item dist\FinSummaryTool-windows-x64.zip }; Compress-Archive -Path dist\FinSummaryTool -DestinationPath dist\FinSummaryTool-windows-x64.zip"

echo.
echo ============================================================
echo  构建完成：
echo    程序目录: dist\FinSummaryTool\  （整个文件夹发给用户，双击 FinSummaryTool.exe）
echo    分发压缩包: dist\FinSummaryTool-windows-x64.zip
echo ============================================================
