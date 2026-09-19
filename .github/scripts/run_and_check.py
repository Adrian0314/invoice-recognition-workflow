# -*- coding: utf-8 -*-
"""
跑发票机器人，并判断这次运行的结果是否"可接受"。

为什么需要这层包装
------------------
`idle_invoice_bot.py` 的退出码是给**人**看的语义，不是给 CI 用的：
    0 = 全部正常
    2 = 批次里存在「异常」或「待人工录入」的行
    1 = 未预期的错误
    3 = 台账被其他程序占用

本仓库的样本集中**故意**埋了边界用例（金额勾稽不平、税号校验位错误、重复提交、
无文本层的扫描件…），所以回归任务里出现退出码 2 是**预期行为**，
不能因此判定 CI 失败 —— 否则这个 job 永远红着，就失去意义了。

真正的判定交给 `assert_ledger.py`（它检查台账结论是否落在基准区间内）。
本脚本只负责：把"预期内的非零退出"和"真的炸了"区分开，并回显机器人的输出。

用法：
    python run_and_check.py <bot.py> -- <机器人的参数...>
例：
    python run_and_check.py optimized_python/idle_invoice_bot.py -- \
        -i invoices_input -o optimized_python/output_v2
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ACCEPTABLE = {0, 2}
MEANING = {
    0: "全部正常",
    1: "未预期的错误（脚本崩溃）",
    2: "存在需要人工处理的项（样本集下的预期结果）",
    3: "台账被其他程序占用，写不进去",
}


def main(argv: list) -> int:
    if "--" not in argv:
        print(__doc__)
        return 2
    cut = argv.index("--")
    bot_args = argv[1:cut]
    run_args = argv[cut + 1:]
    if not bot_args:
        print(__doc__)
        return 2

    bot = Path(bot_args[0])
    if not bot.exists():
        print(f"[FAIL] 找不到机器人脚本：{bot}")
        return 2

    cmd = [sys.executable, "-u", str(bot)] + run_args
    print(f"[run] {' '.join(cmd)}", flush=True)
    p = subprocess.run(cmd, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")

    # 原样回显，便于在 CI 日志里看到识别过程
    if p.stdout:
        print(p.stdout, end="")
    if p.stderr and p.stderr.strip():
        print("--- stderr ---")
        print(p.stderr, end="")

    rc = p.returncode
    print("=" * 70)
    print(f"[run] 退出码 {rc} —— {MEANING.get(rc, '未知退出码')}")
    if rc in ACCEPTABLE:
        print("[run] 视为通过（真正的内容判定交给 assert_ledger.py）")
        return 0
    print("[run] 判定为失败")
    return 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
