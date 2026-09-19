# -*- coding: utf-8 -*-
"""
守护模式：监听文件夹，发票文件落盘即自动入账
=============================================
真实「无人值守」用法：
    python watch.py --input invoices_input --interval 5

工作原理：
  每 N 秒扫描一次目录快照（路径 + 修改时间 + 大小），发现「新出现」或
  「刚写完」的文件（大小连续两次不变，避免读到写到一半的文件）就触发一次
  增量处理；由于台账本身带文件指纹去重，重复触发不会产生重复行。

按 Ctrl+C 退出。
"""
from __future__ import annotations

import argparse
import sys
import time
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from invoice_wf import reader, setup_console                    # noqa: E402
from invoice_wf.config import DEFAULT_INPUT_DIR, DEFAULT_LEDGER, DEFAULT_REPORT  # noqa: E402
from invoice_wf.pipeline import Options, run           # noqa: E402

setup_console()


def snapshot(input_dir: Path, recursive: bool) -> dict:
    snap = {}
    for p in reader.discover_files(input_dir, recursive):
        try:
            st = p.stat()
        except OSError:
            continue
        snap[str(p)] = (st.st_mtime, st.st_size)
    return snap


def watch(input_dir: Path, ledger: str, report: str, interval: int,
          recursive: bool, settle_rounds: int = 2) -> int:
    print("=" * 68)
    print("  发票守护模式已启动 —— 把发票文件放进目录即可自动入账")
    print(f"  监听目录：{input_dir.resolve()}")
    print(f"  台账文件：{Path(ledger).resolve()}")
    print(f"  轮询间隔：{interval}s    按 Ctrl+C 退出")
    print("=" * 68)

    def _run_once(tag: str):
        opts = Options(input_dir=str(input_dir), ledger=ledger, report=report,
                       recursive=recursive, verbose=False)
        run(opts)

    # 启动时先把目录里的存量发票处理一遍（已入账的会自动跳过，不会重复）
    prev = snapshot(input_dir, recursive)
    batch_no = 0
    if prev:
        batch_no += 1
        print(f"[{datetime.now():%H:%M:%S}] 发现 {len(prev)} 个存量文件，先做一次全量处理"
              f"（已入账的会自动跳过）…")
        _run_once("启动全量")
        prev = snapshot(input_dir, recursive)

    stable: dict[str, int] = {}      # 文件 + 连续未变化的轮次数
    print(f"[{datetime.now():%H:%M:%S}] 已就绪，持续监听新增发票…\n")

    while True:
        try:
            time.sleep(interval)
            cur = snapshot(input_dir, recursive)

            # 1) 找出本轮「新出现」或「刚发生变化」的文件，重新计时
            for path, sig in cur.items():
                if prev.get(path) != sig:
                    stable[path] = 0

            # 2) 文件被移走则清理计时器
            for path in list(stable):
                if path not in cur:
                    stable.pop(path, None)

            # 3) 连续 settle_rounds 轮签名不变 → 认为写入完成，可以处理
            ready = []
            for path in list(stable):
                stable[path] += 1
                if stable[path] >= settle_rounds:
                    ready.append(path)
                    stable.pop(path, None)

            prev = cur
            if not ready:
                continue

            batch_no += 1
            print(f"\n[{datetime.now():%H:%M:%S}] 检测到 {len(ready)} 个新文件"
                  f"（第 {batch_no} 批）：")
            for r in ready:
                print(f"    + {Path(r).name}")

            _run_once(f"第 {batch_no} 批")
            print(f"[{datetime.now():%H:%M:%S}] 第 {batch_no} 批处理完成，继续监听…\n")
        except KeyboardInterrupt:
            print("\n已退出守护模式。")
            return 0
        except Exception as exc:                        # noqa: BLE001
            print(f"[{datetime.now():%H:%M:%S}] 本轮异常：{exc}，{interval}s 后继续")
            try:
                time.sleep(interval)
                prev = snapshot(input_dir, recursive)
                stable.clear()
            except KeyboardInterrupt:
                return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="发票目录守护模式（落盘即自动入账）")
    ap.add_argument("--input", "-i", default=DEFAULT_INPUT_DIR)
    ap.add_argument("--output", "-o", default=DEFAULT_LEDGER)
    ap.add_argument("--report", "-r", default=DEFAULT_REPORT)
    ap.add_argument("--interval", type=int, default=5, help="轮询间隔秒数，默认 5")
    ap.add_argument("--no-recursive", action="store_true")
    args = ap.parse_args(argv)

    d = Path(args.input)
    d.mkdir(parents=True, exist_ok=True)
    return watch(d, args.output, args.report, args.interval, not args.no_recursive)


if __name__ == "__main__":
    raise SystemExit(main())
