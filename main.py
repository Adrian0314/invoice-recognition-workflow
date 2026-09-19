# -*- coding: utf-8 -*-
"""
发票自动化处理工作流 —— 命令行入口
====================================
用法示例：
  # 1) 把发票 PDF / 图片丢进 invoices_input 目录，然后执行：
  python main.py

  # 2) 指定输入输出
  python main.py --input D:\\发票\\2026-09 --output D:\\台账\\发票台账.xlsx

  # 3) 干跑（只解析不写表）
  python main.py --dry-run

  # 4) 安装 OCR 后，重试之前无法识别的图片/扫描件
  python main.py --retry-failed

  # 5) 强制重解同一批文件（忽略文件指纹去重）
  python main.py --force

  # 6) 交互式向导
  python main.py --wizard
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from invoice_wf import reader, setup_console        # noqa: E402

setup_console()
from invoice_wf.config import (DEFAULT_INPUT_DIR, DEFAULT_LEDGER,  # noqa: E402
                               DEFAULT_REPORT)
from invoice_wf.pipeline import Options, run      # noqa: E402


def _print_banner():
    print(r"""
+------------------------------------------------------------------+
|            发票自动化处理工作流  Invoice Automation              |
|   PDF/图片 -> 字段识别 -> 校验去重 -> Excel台账 -> 异常报告       |
+------------------------------------------------------------------+""")


def _wizard() -> Options:
    _print_banner()
    print("\n未提供参数，进入交互向导（直接回车使用默认值）\n")

    def ask(prompt: str, default: str) -> str:
        v = input(f"{prompt} [{default}]: ").strip().strip('"')
        return v or default

    input_dir = ask("发票文件所在目录", DEFAULT_INPUT_DIR)
    ledger = ask("Excel 台账输出路径", DEFAULT_LEDGER)
    report = ask("HTML 报告输出路径", DEFAULT_REPORT)

    if not reader.ocr_available():
        print("\n[提示] 未检测到 OCR 引擎，图片与扫描件 PDF 将被标记为「待人工录入」。")
        print(reader.ocr_hint())

    print()
    return Options(input_dir=input_dir, ledger=ledger, report=report)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        description="发票自动化处理工作流：识别 → 校验 → 去重 → 写入 Excel 台账",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--input", "-i", default=None, help="发票文件目录（默认 invoices_input）")
    ap.add_argument("--output", "-o", default=None, help="Excel 台账路径（默认 output/发票台账.xlsx）")
    ap.add_argument("--report", "-r", default=None, help="HTML 报告路径（默认 output/处理报告.html）")
    ap.add_argument("--no-recursive", action="store_true", help="只扫描顶层目录，不递归子目录")
    ap.add_argument("--retry-failed", action="store_true", help="重试之前「待人工录入」的图片/扫描件")
    ap.add_argument("--force", action="store_true", help="忽略文件指纹去重，强制重新解析")
    ap.add_argument("--dry-run", action="store_true", help="只解析不写台账（结果打印到控制台）")
    ap.add_argument("--quiet", action="store_true", help="精简输出")
    ap.add_argument("--wizard", "-w", action="store_true", help="交互式向导")
    args = ap.parse_args(argv)

    if args.wizard or (args.input is None and Path(DEFAULT_INPUT_DIR).exists()
                       and not any(Path(DEFAULT_INPUT_DIR).iterdir())):
        opts = _wizard()
    else:
        opts = Options(
            input_dir=args.input or DEFAULT_INPUT_DIR,
            ledger=args.output or DEFAULT_LEDGER,
            report=args.report or DEFAULT_REPORT,
        )

    opts.recursive = not args.no_recursive
    opts.retry_failed = args.retry_failed
    opts.force = args.force

    if args.dry_run:
        return _dry_run(opts, quiet=args.quiet)

    summary = run(opts)
    s = summary["stats"]
    if s.get("error") or s.get("ocr_todo"):
        return 2      # 有需要人工处理的项
    return 0


def _dry_run(opts: Options, quiet: bool = False) -> int:
    """干跑：解析并校验，但不写任何文件。"""
    from invoice_wf.parser import parse_invoice
    from invoice_wf.validator import summarize_issues, validate

    _print_banner()
    print("[干跑模式] 只解析校验，不写入台账\n")
    files = reader.discover_files(opts.input_dir, opts.recursive)
    print(f"发现 {len(files)} 个待处理文件\n")
    n_ok = n_bad = 0
    for p in files:
        try:
            info = reader.extract_text(p)
            recs = parse_invoice(info["text"], p.name)
        except Exception as exc:                       # noqa: BLE001
            print(f"✗ {p.name}\n   提取失败：{exc}")
            n_bad += 1
            continue
        for i, r in enumerate(recs, 1):
            issues = validate(r)
            tag = f"（第 {i}/{len(recs)} 张）" if len(recs) > 1 else ""
            print(f"✓ {p.name}{tag}  [{info['source']}]")
            print(f"   发票类型 : {r.get('invoice_type') or '—'}")
            print(f"   发票号码 : {r.get('invoice_no') or '—'}" + (
                f"   发票代码 : {r.get('invoice_code')}" if r.get('invoice_code') else ""))
            print(f"   开票日期 : {r.get('issue_date') or '—'}")
            print(f"   购买方   : {r.get('buyer_name') or '—'}  ({r.get('buyer_tax') or '—'})")
            print(f"   销售方   : {r.get('seller_name') or '—'}  ({r.get('seller_tax') or '—'})")
            print(f"   金额/税额/合计 : {r.get('amount')} / {r.get('tax')} / {r.get('total')}")
            print(f"   税率     : {r.get('tax_rate') or '—'}")
            print(f"   项目     : {r.get('item_name') or '—'}")
            if issues:
                print(f"   校验     : {summarize_issues(issues)}")
                n_bad += 1
            else:
                print("   校验     : 全部通过 ✓")
                n_ok += 1
            print()
    print("=" * 60)
    print(f"干跑结束：解析 {n_ok} 张无问题，{n_bad} 张有问题（未写入任何文件）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
