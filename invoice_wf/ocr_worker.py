# -*- coding: utf-8 -*-
"""
OCR 子进程工作器
================
为什么要把 OCR 放进独立进程？
    OCR 依赖 onnxruntime / Paddle / OpenCV 这类**原生库**，它们在部分机器上会直接以
    「访问违例(0xC0000005)」把整个进程打崩。这是 C 层崩溃，Python 的 try/except
    根本拦不住 —— 一旦发生，整批发票的处理结果全部丢失。
    隔离到子进程后，崩溃只影响当前这一个文件：主进程拿到非零退出码，把该文件登记为
    「待人工录入」并继续处理下一张。

用法（由 reader.py 内部调用，一般无需手工执行）：
    python -m invoice_wf.ocr_worker <图片路径> [更多图片路径 ...]

输出：stdout 最后一行 JSON
    成功 {"ok": true,  "texts": ["第1张的文本", "第2张的文本"]}
    失败 {"ok": false, "error": "错误原因"}

注意：JSON 用 ensure_ascii=True 输出，全 ASCII 转义 —— 这样无论子进程的
stdout 是 UTF-8 还是 GBK 代码页，主进程都能正确解析，不会出现编码错乱。
"""
from __future__ import annotations

import json
import sys
import traceback
from pathlib import Path

# 让子进程能 import 到 invoice_wf 包（本文件位于包的上一级目录之下）
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


def main(argv: list[str]) -> int:
    paths = argv[1:]
    if not paths:
        print(json.dumps({"ok": False, "error": "未传入任何图片路径"}, ensure_ascii=True))
        return 2

    try:
        from PIL import Image

        from invoice_wf.reader import _run_ocr_inprocess

        texts = []
        for p in paths:
            img = Image.open(p)
            img.load()
            texts.append(_run_ocr_inprocess(img))
    except Exception as exc:                                # noqa: BLE001
        msg = f"{type(exc).__name__}: {exc}"
        print(json.dumps({"ok": False, "error": msg,
                          "traceback": traceback.format_exc()}, ensure_ascii=True))
        return 1

    print(json.dumps({"ok": True, "texts": texts}, ensure_ascii=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
