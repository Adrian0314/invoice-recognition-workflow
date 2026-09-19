# -*- coding: utf-8 -*-
"""发票自动化处理工作流（Invoice Automation Workflow）"""

import sys

__version__ = "1.0.1"
__all__ = ["config", "reader", "parser", "validator", "excel_store", "report", "pipeline",
           "setup_console"]


def setup_console() -> None:
    """
    让控制台输出更健壮。

    中文 Windows 的控制台代码页是 936(GBK)：
      - 直接输出到控制台时，Python 走 Win32 控制台 API，任何字符都能正常显示；
      - 一旦重定向到文件或管道（例如 python main.py > log.txt），Python 改用 GBK
        编码，而 ✓ ✗ 这类符号并不在 GBK 字符集内，会直接抛 UnicodeEncodeError 中断程序。
    这里把 errors 放宽为 replace，遇到无法编码的字符退化成 ?，保证流程不被打断。
    """
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream is not None and hasattr(stream, "reconfigure"):
                stream.reconfigure(errors="replace")
        except Exception:
            pass
