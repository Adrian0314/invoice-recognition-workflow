# -*- coding: utf-8 -*-
"""
校验 docs/ 在线版站点的完整性（零依赖）。

为什么需要它
------------
网站**直接引用仓库里的文件**，而不是拷贝一份：
    docs/app.js  →  ../optimized_python/idle_invoice_bot.py
                 →  ../optimized_python/invoice_templates.json
                 →  ../samples*/…（示例数据清单）
这种"零副本"设计保证网站与主实现永不脱节，代价是名字一改就会**静默弄坏网站**。
所以每次改动都跑一遍本脚本，把这种破坏变成 CI 上的红灯。

检查项：
  1. 站点三个文件都在（index.html / app.js / style.css）
  2. app.js 声明引用的引擎与模板，按相对路径真实存在
  3. app.js 的示例清单（DEMO_MANIFEST）里每个文件都存在
  4. index.html 正确引用 style.css / app.js，并含 Pyodide CDN 脚本

用法：python .github/scripts/assert_site.py
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DOCS = ROOT / "docs"

ok: list[str] = []
bad: list[str] = []


def check(cond: bool, good: str, bad_msg: str) -> None:
    (ok if cond else bad).append(good if cond else bad_msg)


# ---------- 1) 站点文件 ----------
files_ok = True
for name in ("index.html", "app.js", "style.css"):
    if not (DOCS / name).is_file():
        bad.append(f"缺少站点文件：docs/{name}")
        files_ok = False
check(files_ok, "站点三件套齐全（index.html / app.js / style.css）", "")

if not files_ok:
    print("站点文件检查失败：")
    for p in bad:
        print(f"  ✗ {p}")
    raise SystemExit(1)

app = (DOCS / "app.js").read_text(encoding="utf-8")
html = (DOCS / "index.html").read_text(encoding="utf-8")


def ref_exists(var: str, label: str) -> None:
    m = re.search(rf"{var}\s*=\s*new URL\(\s*'([^']+)'", app)
    if not m:
        bad.append(f"app.js 里找不到 {var} 声明")
        return
    ref = m.group(1)
    # 注意：不要用 lstrip("./") —— 它按字符集剥离，会把 "../" 的前导点与斜杠一起吃掉。
    # Path 本身就能正确处理 ".."，交给它即可。
    target = (DOCS / ref).resolve()
    check(target.is_file(),
          f"{label}引用有效 → {target.relative_to(ROOT)}",
          f"{label}引用失效：app.js 里写的是 {ref!r}，解析到 {target}，但该文件不存在")


ref_exists("ENGINE_URL", "引擎源文件")
ref_exists("TEMPLATE_URL", "版式模板")

# ---------- 3) 示例清单 ----------
block = re.search(r"const DEMO_MANIFEST\s*=\s*\[(.*?)\];", app, re.S)
if not block:
    bad.append("app.js 里找不到 DEMO_MANIFEST 定义")
else:
    paths = re.findall(r"'([^']+)'", block.group(1))
    check(len(paths) >= 20,
          f"示例清单包含 {len(paths)} 项", f"示例清单只解析出 {len(paths)} 项，偏少")
    missing = [p for p in paths if not (ROOT / p).is_file()]
    check(not missing,
          f"示例清单 {len(paths)} 个文件全部存在",
          "示例清单里有文件不存在：" + "、".join(missing[:6])
          + ("…" if len(missing) > 6 else ""))

# ---------- 4) index.html ----------
check('href="style.css"' in html, "index.html 引用 style.css", "index.html 没有引用 style.css")
check('src="app.js"' in html, "index.html 引用 app.js", "index.html 没有引用 app.js")
check("cdn.jsdelivr.net/pyodide/" in html, "index.html 含 Pyodide CDN 脚本",
      "index.html 里找不到 Pyodide CDN 脚本（网站将无法启动）")

# ---------- 输出 ----------
print("docs/ 站点完整性检查：")
for line in ok:
    print(f"  ✓ {line}")
if bad:
    print("")
    print("以下问题会导致网站不可用：")
    for p in bad:
        print(f"  ✗ {p}")
    raise SystemExit(1)
print("")
print("[PASS] 站点引用完整：引擎、版式模板、示例清单都能解析到真实文件")
