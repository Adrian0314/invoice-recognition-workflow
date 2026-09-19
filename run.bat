@echo off
chcp 936 >nul 2>nul
setlocal
pushd "%~dp0"

rem 解释器：优先用本目录下的 .venv，其次用 PATH 里的 python
set "PYEXE=%~dp0.venv\Scripts\python.exe"
if not exist "%PYEXE%" set "PYEXE=python"
"%PYEXE%" --version >nul 2>nul
if errorlevel 1 (
  echo.
  echo [错误] 未找到可用的 Python。
  echo   请安装 Python 3.9+ 并加入 PATH，或在项目根目录创建 .venv
  pause
  exit /b 1
)

echo ============================================================
echo   发票自动化处理工作流
echo ============================================================
echo.
echo   [1] 处理 invoices_input 目录（生成台账 + 报告）
echo   [2] 干跑预览（只看识别结果，不写任何文件）
echo   [3] 守护模式（发票落盘即自动入账，Ctrl+C 退出）
echo   [4] 重试之前无法识别的图片 / 扫描件
echo   [5] 生成演示发票样本
echo   [0] 退出
echo.
set "choice=1"
set /p "choice=请选择 [1]: "
if "%choice%"=="" set "choice=1"

if "%choice%"=="1" "%PYEXE%" main.py
if "%choice%"=="2" "%PYEXE%" main.py --dry-run
if "%choice%"=="3" "%PYEXE%" watch.py
if "%choice%"=="4" "%PYEXE%" main.py --retry-failed
if "%choice%"=="5" "%PYEXE%" make_samples.py
if "%choice%"=="0" goto :end

echo.
echo 输出目录：%cd%\output
echo.
pause

:end
popd
endlocal
