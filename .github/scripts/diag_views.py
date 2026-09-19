# -*- coding: utf-8 -*-
"""
诊断：把每个文本视图的抽取结果与解析得分打出来。
=========================================================
用途：当回归任务的台账断言失败时，由 CI 自动运行（if: failure()），
      回答"到底拿到了几个视图、每个视图长什么样、解析得分多少"，
      避免靠猜排障。

本地也能跑：
    python .github/scripts/diag_views.py
"""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

BOT = Path("optimized_python/idle_invoice_bot.py")
INPUT_DIR = Path("invoices_input")

spec = importlib.util.spec_from_file_location("bot", BOT)
bot = importlib.util.module_from_spec(spec)
sys.modules["bot"] = bot
spec.loader.exec_module(bot)

cfg = bot.load_config()
templates = bot.load_templates()

print("=" * 78)
print("环境")
print(f"  PyMuPDF 可用 : {bot._have('pymupdf') or bot._have('fitz')}")
pyv = None
for name in ("pymupdf", "fitz"):
    if bot._have(name):
        try:
            mod = importlib.import_module(name)
            pyv = getattr(mod, "__version__", None) or getattr(mod, "version", None)
            if callable(pyv):
                pyv = pyv()
            break
        except Exception:                                   # noqa: BLE001
            pass
print(f"  PyMuPDF 版本 : {pyv}")
print(f"  openpyxl     : {bot._have('openpyxl')}")
print(f"  Pillow       : {bot._have('PIL')}")
print(f"  OCR          : {bot.ocr_provider()}")
print(f"  输入目录      : {INPUT_DIR.resolve()}")

# 挑几个有代表性的文件：中文样本 + 国际样本
targets = []
for pat in ("01_*.pdf", "05_*.pdf", "T1-*.pdf", "T4-*.pdf", "虚拟IT*.pdf"):
    targets += sorted(INPUT_DIR.glob(pat))
targets = targets[:8]

for p in targets:
    print("=" * 78)
    print(f"文件：{p.name}")
    try:
        views, pages = bot._pdf_text_views(p)
    except Exception as exc:                                # noqa: BLE001
        print(f"  _pdf_text_views 抛异常：{exc}")
        continue
    print(f"  页数={pages}  视图数={len(views)}")
    for src, t in views:
        q = bot.text_quality(t)
        recs = bot.parse_invoice(t, p.name, templates)
        sc = bot._records_score(recs)
        r0 = recs[0] if recs else {}
        print(f"  --- {src}")
        print(f"      len={len(t)}  quality={q['score']}  labels={q['labels']}  "
              f"reasons={q['reasons']}")
        print(f"      解析出 {len(recs)} 条  得分={sc}")
        print(f"      号码={r0.get('invoice_no')!r} 日期={r0.get('issue_date')!r} "
              f"合计={r0.get('total')!r} 销售方={str(r0.get('seller_name'))[:18]!r}")
        head = t[:260].replace("\n", "\\n")
        print(f"      文本头: {head!r}")

    # 实际会选哪个视图
    info = bot.extract_text(p, cfg)
    views_all = [(info["source"], info["text"])] + list(info.get("alts") or [])
    picked = None
    for src_i, txt_i in views_all:
        sc_i = bot._records_score(bot.parse_invoice(txt_i, p.name, templates))
        if picked is None or sc_i > picked[1]:
            picked = (src_i, sc_i)
    print(f"  >>> extract_text 主视图={info['source']} 备选={[s for s, _ in (info.get('alts') or [])]}"
          f"  → 择优结果={picked}")

print("=" * 78)
print("诊断结束")
