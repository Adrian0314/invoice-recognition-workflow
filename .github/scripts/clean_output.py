# -*- coding: utf-8 -*-
"""
CI 辅助：清空输出目录
=====================
为什么单独写一个脚本而不是用 rm -rf：
  1. 跨平台（CI 上有 ubuntu 也有 windows runner，rm -rf 在 windows 上不可用）；
  2. 本机有安全策略会拦截递归删除，逐文件删更稳。

用法：
    python clean_output.py <目录>
"""
from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path


def main(argv) -> int:
    if len(argv) < 2:
        print(__doc__)
        return 1
    target = Path(argv[1])
    if not target.exists():
        print(f"[skip] 目录不存在：{target}")
        return 0

    removed = 0
    for p in sorted(target.rglob("*"), key=lambda x: -len(str(x))):
        try:
            if p.is_file() or p.is_symlink():
                os.remove(p)
                removed += 1
            elif p.is_dir():
                try:
                    p.rmdir()
                except OSError:
                    pass
        except OSError as exc:
            print(f"[warn] 删除失败 {p}: {exc}")
    print(f"[clean] {target} 已清空，删除 {removed} 个文件")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
