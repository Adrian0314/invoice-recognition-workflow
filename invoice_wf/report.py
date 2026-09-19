# -*- coding: utf-8 -*-
"""
报告层：生成自包含的 HTML 处理报告（可直接在浏览器打开 / 打印归档）
"""
from __future__ import annotations

import html
from datetime import datetime
from pathlib import Path

from .config import FIELDS, SETTINGS, STATUS_DUPLICATE, STATUS_ERROR, STATUS_OCR_TODO
from .reader import ocr_hint


def _e(v) -> str:
    return html.escape(str(v if v is not None else ""))


def _money(v) -> str:
    if v in ("", None):
        return "—"
    try:
        return f"{float(v):,.2f}"
    except (TypeError, ValueError):
        return _e(v)


def _badge(status: str) -> str:
    cls = "b-ok"
    if status == STATUS_ERROR:
        cls = "b-err"
    elif status in ("存在提示", "警告"):
        cls = "b-warn"
    elif status == STATUS_OCR_TODO:
        cls = "b-todo"
    elif "重复" in status:
        cls = "b-dup"
    return f'<span class="badge {cls}">{_e(status)}</span>'


def build_report(ledger, summary: dict, out_path: str | Path) -> Path:
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    s = ledger.stats

    # ---------------- KPI 卡片 ----------------
    cards = [
        ("台账发票行数", f"{len(ledger.rows)}", "一张发票一行"),
        ("本次新增发票", f"{s.get('new_invoices', 0)}", f"新增文件 {s.get('new_files', 0)} 个"),
        ("校验异常", f"{s.get('error', 0)}", "需人工核对", "err"),
        ("存在提示", f"{s.get('warn', 0)}", "建议复核", "warn"),
        ("重复拦截", f"{s.get('duplicate', 0)}", "已阻止重复入账", "dup"),
        ("待人工录入", f"{s.get('ocr_todo', 0)}", "图片/扫描件", "todo"),
        ("合计金额(不含税)", f"¥{s.get('sum_amount', 0):,.2f}", "", "money"),
        ("合计税额", f"¥{s.get('sum_tax', 0):,.2f}", "", "money"),
        ("合计价税合计", f"¥{s.get('sum_total', 0):,.2f}", "", "money"),
    ]
    card_html = "".join(
        f'<div class="card {c[3] if len(c) > 3 else ""}">'
        f'<div class="c-label">{_e(c[0])}</div>'
        f'<div class="c-value">{_e(c[1])}</div>'
        f'<div class="c-sub">{_e(c[2])}</div></div>'
        for c in cards)

    # ---------------- 异常清单 ----------------
    order = {"严重": 0, "警告": 1, "重复": 2, "提示": 3}
    issues = sorted(ledger.issue_rows, key=lambda r: order.get(str(r.get("问题级别")), 9))
    rows_issue = "".join(
        '<tr class="lv-{k}"><td>{i}</td><td class="mono">{f}</td><td class="mono">{n}</td>'
        '<td><span class="badge b-{k2}">{lv}</span></td><td>{fd}</td><td>{msg}</td>'
        '<td class="dim">{sg}</td></tr>'.format(
            i=i, k={"严重": "err", "警告": "warn", "重复": "dup", "提示": "hint"}.get(
                str(r.get("问题级别")), "hint"),
            k2={"严重": "err", "警告": "warn", "重复": "dup", "提示": "hint"}.get(
                str(r.get("问题级别")), "hint"),
            f=_e(r.get("源文件名", "")), n=_e(r.get("发票号码", "")),
            lv=_e(r.get("问题级别", "")), fd=_e(r.get("涉及字段", "")),
            msg=_e(r.get("问题描述", "")), sg=_e(r.get("建议处理", "")))
        for i, r in enumerate(issues, 1))
    if not rows_issue:
        rows_issue = '<tr><td colspan="7" class="empty">本次没有发现任何问题，全部发票校验通过 ✓</td></tr>'

    # ---------------- 明细表 ----------------
    show_cols = [f for f in FIELDS if not f.get("hidden") and f["key"] != "issue"]
    thead = "".join(f"<th>{_e(f['header'])}</th>" for f in show_cols)
    body = []
    for r in ledger.rows:
        tds = []
        for f in show_cols:
            v = r.get(f["key"])
            if f["kind"] == "money":
                cls = "num"
                txt = _money(v)
            elif f["key"] == "status":
                tds.append(f'<td class="ctr">{_badge(str(v or ""))}</td>')
                continue
            else:
                cls = "mono" if f["key"] in ("invoice_no", "invoice_code", "buyer_tax",
                                             "seller_tax", "file_hash") else ""
                txt = _e(v if v not in ("", None) else "—")
            tds.append(f'<td class="{cls}">{txt}</td>')
        row_cls = ""
        if str(r.get("status")) == STATUS_ERROR:
            row_cls = ' class="row-err"'
        elif str(r.get("status")) == STATUS_OCR_TODO:
            row_cls = ' class="row-todo"'
        elif "重复" in str(r.get("status")):
            row_cls = ' class="row-dup"'
        body.append(f"<tr{row_cls}>{''.join(tds)}</tr>")
    if not body:
        body = [f'<tr><td colspan="{len(show_cols)}" class="empty">台账中暂无记录</td></tr>']

    # ---------------- OCR 提示 ----------------
    ocr_banner = ""
    if not summary.get("ocr_available"):
        ocr_banner = (
            '<div class="banner">'
            '<b>⚠ 未检测到 OCR 引擎</b>：电子版 PDF 发票已正常识别；'
            '图片（JPG/PNG）与扫描件 PDF 已被登记为「待人工录入」并计入异常清单。'
            '安装 OCR 后执行 <code>python main.py --retry-failed</code> 即可自动补全。'
            f'<pre>{_e(ocr_hint())}</pre></div>')

    html_doc = f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>发票处理报告 · {_e(summary.get('finished_at', datetime.now()).strftime('%Y-%m-%d %H:%M'))}</title>
<style>
  :root {{
    --bg:#f5f7fb; --panel:#ffffff; --line:#e3e8f0; --text:#1e2a3a; --dim:#7b8794;
    --blue:#1f4e79; --blue-soft:#eaf1f8; --red:#c0392b; --red-soft:#fdecea;
    --amber:#b7791f; --amber-soft:#fff7e6; --green:#1e7145; --green-soft:#e8f5ee;
  }}
  * {{ box-sizing:border-box; }}
  body {{ margin:0; background:var(--bg); color:var(--text);
    font-family:"Microsoft YaHei","PingFang SC","Segoe UI",system-ui,sans-serif;
    font-size:14px; line-height:1.6; }}
  .wrap {{ max-width:1500px; margin:0 auto; padding:28px 24px 60px; }}
  header {{ background:linear-gradient(135deg,#1f4e79,#2e6da4); color:#fff;
    border-radius:14px; padding:26px 30px; box-shadow:0 6px 22px rgba(31,78,121,.18); }}
  header h1 {{ margin:0 0 6px; font-size:24px; letter-spacing:.5px; }}
  header .meta {{ opacity:.9; font-size:13px; }}
  header .meta span {{ margin-right:18px; white-space:nowrap; }}
  .cards {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(160px,1fr));
    gap:14px; margin:22px 0 8px; }}
  .card {{ background:var(--panel); border:1px solid var(--line); border-radius:12px;
    padding:16px 18px; box-shadow:0 2px 8px rgba(16,24,40,.04); }}
  .card .c-label {{ font-size:12px; color:var(--dim); }}
  .card .c-value {{ font-size:24px; font-weight:700; margin:4px 0 2px; color:var(--blue); }}
  .card .c-sub {{ font-size:11px; color:var(--dim); }}
  .card.err .c-value {{ color:var(--red); }}
  .card.warn .c-value {{ color:var(--amber); }}
  .card.dup .c-value, .card.todo .c-value {{ color:#8a6d3b; }}
  .card.money .c-value {{ font-size:20px; }}
  .banner {{ background:var(--amber-soft); border:1px solid #f0d9a8; border-left:4px solid var(--amber);
    border-radius:10px; padding:14px 18px; margin:18px 0; font-size:13px; }}
  .banner pre {{ margin:8px 0 0; white-space:pre-wrap; font-size:12px; color:#5c4a1f; }}
  h2 {{ font-size:17px; margin:30px 0 12px; padding-left:11px; border-left:4px solid var(--blue); }}
  .panel {{ background:var(--panel); border:1px solid var(--line); border-radius:12px;
    overflow:auto; max-height:640px; box-shadow:0 2px 8px rgba(16,24,40,.04); }}
  table {{ border-collapse:separate; border-spacing:0; width:100%; font-size:12.5px; }}
  th {{ position:sticky; top:0; background:#eef3f9; color:var(--blue); font-weight:600;
    text-align:left; padding:10px 10px; border-bottom:1px solid var(--line); white-space:nowrap; z-index:2; }}
  td {{ padding:8px 10px; border-bottom:1px solid #f0f3f8; vertical-align:top; }}
  tbody tr:hover {{ background:#f8fbff; }}
  tr.row-err {{ background:var(--red-soft); }}
  tr.row-err:hover {{ background:#fbdad6; }}
  tr.row-dup {{ background:var(--amber-soft); }}
  tr.row-todo {{ background:#f2eefb; }}
  td.num {{ text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }}
  td.ctr {{ text-align:center; white-space:nowrap; }}
  td.mono {{ font-family:Consolas,"Courier New",monospace; font-size:12px; white-space:nowrap; }}
  td.dim {{ color:var(--dim); font-size:12px; }}
  td.empty {{ text-align:center; color:var(--green); padding:26px; font-weight:600; }}
  .badge {{ display:inline-block; padding:2px 9px; border-radius:20px; font-size:11.5px;
    font-weight:600; white-space:nowrap; }}
  .b-ok {{ background:var(--green-soft); color:var(--green); }}
  .b-err {{ background:#f8d7d3; color:var(--red); }}
  .b-warn {{ background:#fdeec7; color:var(--amber); }}
  .b-dup {{ background:#fdeec7; color:#8a6d3b; }}
  .b-todo {{ background:#e6e0f5; color:#5b4b8a; }}
  .b-hint {{ background:#e9eef5; color:#5a6b7d; }}
  footer {{ margin-top:34px; color:var(--dim); font-size:12px; line-height:1.9; }}
  code {{ background:#eef2f7; padding:1px 6px; border-radius:5px;
    font-family:Consolas,monospace; font-size:12px; }}
</style></head><body><div class="wrap">
<header>
  <h1>发票自动化处理报告</h1>
  <div class="meta">
    <span>生成时间：{_e(summary.get('finished_at', datetime.now()).strftime('%Y-%m-%d %H:%M:%S'))}</span>
    <span>批次耗时：{summary.get('elapsed', 0):.2f}s</span>
    <span>输入目录：{_e(summary.get('input_dir', ''))}</span>
    <span>扫描文件：{summary.get('files', 0)} 个</span>
  </div>
</header>
{ocr_banner}
<div class="cards">{card_html}</div>

<h2>异常与提示清单（{len(issues)} 条）</h2>
<div class="panel"><table>
  <thead><tr><th>#</th><th>源文件名</th><th>发票号码</th><th>级别</th>
  <th>涉及字段</th><th>问题描述</th><th>建议处理</th></tr></thead>
  <tbody>{rows_issue}</tbody>
</table></div>

<h2>发票明细台账（{len(ledger.rows)} 行 · 一张发票一行）</h2>
<div class="panel"><table>
  <thead><tr>{thead}</tr></thead>
  <tbody>{''.join(body)}</tbody>
</table></div>

<footer>
  台账文件：<code>{_e(summary.get('ledger', ''))}</code><br>
  结构化导出：<code>{_e(summary.get('extras', {}).get('csv', ''))}</code> ·
  <code>{_e(summary.get('extras', {}).get('json', ''))}</code><br>
  说明：金额合计仅统计「真正入账」的行；标记为「重复未入账」「待人工录入」的行不计入合计。
</footer>
</div></body></html>"""

    out.write_text(html_doc, encoding="utf-8")
    return out
