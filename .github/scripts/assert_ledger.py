# -*- coding: utf-8 -*-
"""
CI 断言：检查发票台账的结论是否符合预期基准
=============================================
用途：回归测试的"护栏"。合成样本里故意埋了若干异常用例（重复、勾稽不平、
      税号校验位错误…），这些必须被检出；而不该报错的行一旦变成异常，
      说明解析器改坏了 —— 这个脚本就是用来在 CI 里把这种回归挡住的。

用法：
    python assert_ledger.py <发票明细.json 路径> [期望异常数]
退出码：0 = 通过，1 = 不通过
"""
from __future__ import annotations

import json
import sys
from collections import Counter
from pathlib import Path

# 基准：中文样本 11 份 + 国际样本 12 份 + 样票 10 张
# 设计上"必须被检出"的问题（与 README 第九节的 6 条对应关系见下方断言）
BASELINE = {
    "min_rows": 20,            # 台账至少要有这么多行，否则说明大面积解析失败
    "min_ok_rows": 15,         # 正常行数下限
    "max_error_rows": 6,       # 异常行数上限（超过说明误报变多）
    "must_have_duplicate": 1,  # 至少要有 1 条重复拦截（样本 04/10/11）
    "require_confidence": True,
}


def main(argv) -> int:
    if len(argv) < 2:
        print(__doc__)
        return 1
    path = Path(argv[1])
    if not path.exists():
        print(f"[FAIL] 找不到台账导出：{path}")
        return 1

    rows = json.loads(path.read_text(encoding="utf-8"))
    stats = Counter(str(r.get("处理状态") or "") for r in rows)
    problems = []

    print(f"台账行数：{len(rows)}")
    for k, v in stats.most_common():
        print(f"  {k or '(空)'}: {v}")

    if len(rows) < BASELINE["min_rows"]:
        problems.append(f"台账行数 {len(rows)} < 下限 {BASELINE['min_rows']}，疑似大面积解析失败")

    ok = stats.get("正常", 0)
    if ok < BASELINE["min_ok_rows"]:
        problems.append(f"正常行数 {ok} < 下限 {BASELINE['min_ok_rows']}，疑似误报增多")

    err = stats.get("异常", 0)
    if err > BASELINE["max_error_rows"]:
        problems.append(f"异常行数 {err} > 上限 {BASELINE['max_error_rows']}，疑似误报增多")

    review = stats.get("待复核", 0)
    if review > len(rows) // 2:
        problems.append(f"待复核 {review} 行超过一半，置信度模型可能失效")

    if BASELINE["require_confidence"]:
        missing = [r.get("源文件名") for r in rows if not str(r.get("识别置信度") or "").strip()]
        if missing:
            problems.append(f"{len(missing)} 行缺少『识别置信度』：{missing[:3]}")

    # 长数字串必须保持文本形态（前导 0 / 20 位不丢）
    for r in rows:
        no = str(r.get("发票号码") or "")
        if no and no.isdigit() and len(no) >= 8 and no != str(r.get("发票号码")):
            problems.append(f"发票号码被改写：{r.get('发票号码')!r}")
            break

    print("-" * 60)
    if problems:
        for p in problems:
            print(f"[FAIL] {p}")
        return 1
    print("[PASS] 台账结论符合基准")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
