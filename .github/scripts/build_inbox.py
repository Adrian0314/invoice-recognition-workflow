# -*- coding: utf-8 -*-
"""
从仓库内**已提交的样本**构建回归测试用的收件箱。

为什么需要它
-----------
`invoices_input/` 是**运行时收件箱**（工作目录）：内容随时会被人工/程序改变，
所以它本身不进版本库。但 CI 的回归测试需要一个**确定性**的输入，
否则每次跑的东西都不一样，断言就失去意义 ——
（此前 CI 现场用 PyMuPDF 重新生成样本，结果随 runner 上的 PyMuPDF 版本漂移，
  把回归任务弄红过两次。）

所以收件箱改由已提交的样本拼出来：

    samples/*.pdf                             11 份中文增值税样本
  + samples_intl/*.pdf                        12 份国际商务样本
  + samples_mock/虚拟IT公司样票_10张.pdf        1 份 10 页合成的多票样张
  = 24 个文件  →  基准 31 行台账

用法：
    python .github/scripts/build_inbox.py [目标目录，默认 invoices_input]
    python .github/scripts/build_inbox.py --show     # 只列清单不复制
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SRC_DIRS = [ROOT / "samples", ROOT / "samples_intl"]
SRC_FILES = [ROOT / "samples_mock" / "虚拟IT公司样票_10张.pdf"]
EXTS = {".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}


def main(argv: list) -> int:
    show = "--show" in argv
    rest = [a for a in argv[1:] if not a.startswith("--")]
    inbox = Path(rest[0]) if rest else (ROOT / "invoices_input")
    if not inbox.is_absolute():
        inbox = ROOT / inbox

    files = []
    for d in SRC_DIRS:
        if d.is_dir():
            files += [p for p in sorted(d.iterdir())
                      if p.is_file() and p.suffix.lower() in EXTS]
    for f in SRC_FILES:
        if f.is_file():
            files.append(f)

    print(f"来源样本共 {len(files)} 个文件：")
    for f in files:
        print(f"  {f.relative_to(ROOT)}")

    if show:
        print("（--show：未做任何复制）")
        return 0

    missing = [f for f in SRC_FILES if not f.is_file()]
    if missing:
        print("[FAIL] 缺少必需样本：" + ", ".join(str(m.relative_to(ROOT)) for m in missing))
        return 1
    if not files:
        print("[FAIL] 没找到任何样本，检查 samples/ 与 samples_intl/ 是否完好")
        return 1

    inbox.mkdir(parents=True, exist_ok=True)
    # 清空旧内容，确保每次输入的集合完全一致
    for p in list(inbox.iterdir()):
        if p.is_file() and p.suffix.lower() in EXTS:
            p.unlink()
        elif p.is_dir():
            shutil.rmtree(p, ignore_errors=True)

    for f in files:
        shutil.copyfile(f, inbox / f.name)

    got = [p for p in inbox.iterdir() if p.is_file() and p.suffix.lower() in EXTS]
    print(f"已写入 {inbox.relative_to(ROOT)}：{len(got)} 个文件")
    return 0 if len(got) == len(files) else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
