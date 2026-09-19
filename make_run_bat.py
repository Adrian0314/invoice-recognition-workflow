# -*- coding: utf-8 -*-
"""
重新生成 run.bat（GBK 编码 + CRLF 换行）
=========================================
为什么需要这个脚本？
  Windows 的 cmd.exe 解析批处理文件有两个硬性要求：
    1. 换行必须是 CRLF。只有 LF 会让整份脚本被错乱切分，典型症状是
       报出「'C:\\Windows\\System32\\cmd.exe' 不是内部或外部命令」这类莫名的错误。
    2. 中文 Windows 控制台默认代码页是 936(GBK)，文件正文必须是 GBK 字节，
       否则 echo 出来的中文是乱码（UTF-8 字节被按 GBK 解读）。
  而绝大多数编辑器默认按 UTF-8 保存，一不小心就会把 run.bat 弄坏。
  一旦弄坏，运行本脚本即可恢复。

用法：
    python make_run_bat.py
另外脚本会检查「汉字的 GBK 尾字节恰好是 0x5C(反斜杠)」这一批处理经典陷阱。
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
TARGET = ROOT / "run.bat"

# run.bat 源码（用 LF 书写，写出时统一转 CRLF）
BAT_SOURCE = r"""@echo off
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
"""


def check_hazard(text: str) -> list[tuple[str, str]]:
    """检查汉字 GBK 尾字节是否为 0x5C——会让 cmd 把半个汉字当成路径分隔符。"""
    bad = []
    for ch in sorted(set(text)):
        if ord(ch) < 128:
            continue
        try:
            enc = ch.encode("gbk")
        except UnicodeEncodeError:
            bad.append((ch, "GBK 无法表示该字符"))
            continue
        if len(enc) == 2 and enc[1] == 0x5C:
            bad.append((ch, "GBK 尾字节为 0x5C，会破坏批处理语法"))
    return bad


def main() -> int:
    text = BAT_SOURCE.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    bad = check_hazard(text)
    if bad:
        print("[警告] 发现危险字符：")
        for ch, why in bad:
            print(f"   {ch!r} —— {why}")
        print("请改用其他措辞，否则 cmd 解析会出错。")
        return 1

    TARGET.write_bytes(text.encode("gbk"))

    data = TARGET.read_bytes()
    print(f"已生成 {TARGET}")
    print(f"  大小      : {len(data)} 字节")
    print(f"  BOM       : {'有（会导致首行报错，需去掉）' if data[:3] == b'\xef\xbb\xbf' else '无'}")
    print(f"  CRLF 行数 : {data.count(b'\r\n')}    裸 LF: {data.count(b'\n') - data.count(b'\r\n')}")
    try:
        data.decode("gbk")
        print("  GBK 解码  : 通过")
    except UnicodeDecodeError as exc:
        print(f"  GBK 解码  : 失败 -> {exc}")
        return 1
    print(" 危险字符  : 无")
    print("\n完成，可直接双击 run.bat 使用。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
