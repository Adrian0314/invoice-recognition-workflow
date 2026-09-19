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
echo   发票识别工作流
echo ============================================================
echo.
echo   -- v2 单文件版（零依赖，推荐）--
echo   [1] 自检（不读文件、不装任何第三方包，秒级）
echo   [2] 处理 invoices_input 并生成报告
echo   [3] 干跑预览（只看识别结果，不写任何文件）
echo   [4] 强制重算（装好 OCR 或改完版式模板后用）
echo.
echo   -- v1 原方案（需先 pip install -r requirements.txt）--
echo   [5] 处理 invoices_input（写入 output）
echo   [6] 干跑预览
echo   [7] 守护模式（落盘即自动入账，Ctrl+C 退出）
echo   [8] 重试之前无法识别的图片 / 扫描件
echo   [9] 生成演示发票样本
echo.
echo   [0] 退出
echo.
set "choice=2"
set /p "choice=请选择 [2]: "
if "%choice%"=="" set "choice=2"

if "%choice%"=="1" "%PYEXE%" optimized_python\idle_invoice_bot.py --selftest
if "%choice%"=="2" "%PYEXE%" optimized_python\idle_invoice_bot.py -i invoices_input -o optimized_python\output_v2
if "%choice%"=="3" "%PYEXE%" optimized_python\idle_invoice_bot.py -i invoices_input -o optimized_python\output_v2 --dry-run
if "%choice%"=="4" "%PYEXE%" optimized_python\idle_invoice_bot.py -i invoices_input -o optimized_python\output_v2 --force
if "%choice%"=="5" "%PYEXE%" main.py
if "%choice%"=="6" "%PYEXE%" main.py --dry-run
if "%choice%"=="7" "%PYEXE%" watch.py
if "%choice%"=="8" "%PYEXE%" main.py --retry-failed
if "%choice%"=="9" "%PYEXE%" make_samples.py
if "%choice%"=="0" goto :end

echo.
echo 产出目录：
echo   v2  %cd%\optimized_python\output_v2
echo   v1  %cd%\output
echo.
pause

:end
popd
endlocal
