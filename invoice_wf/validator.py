# -*- coding: utf-8 -*-
"""
校验层
------
对解析结果做「财务级」校验，产出分级问题清单：
  严重  → 该行数据不可直接使用，需人工核对
  警告  → 数据存疑，建议核对
  提示  → 仅供参考的合规/格式提醒

校验项：
  1. 必填字段缺失（号码 / 开票日期 / 价税合计 / 购销双方名称）
  2. 发票号码位数与版式是否匹配
  3. 开票日期可解析性、是否未来日期、是否过于久远
  4. 三值勾稽：金额(不含税) + 税额 = 价税合计（容差 0.02）
  5. 金额为负 → 红字发票提醒
  6. 纳税人识别号格式 + 统一社会信用代码校验位（GB 32100-2015）
  7. 购销双方同名（自查自开，多为信息填写错误）
  8. 税率是否在合法集合内
  9. 老版发票缺少发票代码
"""
from __future__ import annotations

import re
from datetime import date

from .config import SETTINGS, VALID_TAX_RATES

LEVEL_ERROR, LEVEL_WARN, LEVEL_HINT = "严重", "警告", "提示"

# GB 32100-2015 统一社会信用代码：base31 字符表 + 前 17 位加权因子
USCC_CHARS = "0123456789ABCDEFGHJKLMNPQRTUWXY"
USCC_WEIGHTS = [1, 3, 9, 27, 19, 26, 16, 17, 20, 29, 25, 13, 8, 24, 10, 30, 28]


def uscc_check(code: str):
    """返回 True/False；非 18 位则返回 None（不适用）。"""
    code = (code or "").strip().upper()
    if len(code) != 18:
        return None
    total = 0
    for ch, w in zip(code[:17], USCC_WEIGHTS):
        idx = USCC_CHARS.find(ch)
        if idx < 0:
            return False
        total += idx * w
    c = 31 - (total % 31)
    if c == 31:
        c = 0
    return USCC_CHARS[c] == code[17]


def validate(rec: dict, today: date | None = None) -> list[dict]:
    today = today or date.today()
    issues: list[dict] = []

    def add(level: str, field: str, message: str, suggest: str = ""):
        issues.append({"level": level, "field": field,
                       "message": message, "suggest": suggest})

    # --- 1. 必填 ---
    required = [
        ("invoice_no", "发票号码", "请确认发票 PDF 是否为完整原件，必要时人工补录"),
        ("issue_date", "开票日期", "检查票据是否清晰、是否被裁切"),
        ("total", "价税合计", "价税合计未能识别，请核对发票‘价税合计(小写)’字样"),
        ("seller_name", "销售方名称", "销售方名称未识别，请人工补录"),
        ("buyer_name", "购买方名称", "购买方名称未识别（个人抬头可能为空，请确认）"),
    ]
    for key, label, suggest in required:
        if rec.get(key) in (None, "", []):
            add(LEVEL_ERROR, label, f"未识别到「{label}」", suggest)

    # --- 2. 发票号码位数 ---
    no = str(rec.get("invoice_no") or "")
    itype = rec.get("invoice_type") or ""
    if no:
        if not no.isdigit():
            add(LEVEL_ERROR, "发票号码", f"发票号码含非数字字符：{no}", "人工核对")
        elif "数电" in itype or len(no) == 20:
            if len(no) != 20:
                add(LEVEL_WARN, "发票号码",
                    f"数电发票号码应为 20 位，当前 {len(no)} 位", "人工核对号码是否被截断")
        elif len(no) not in (8, 9, 10, 12):
            add(LEVEL_WARN, "发票号码", f"发票号码为 {len(no)} 位，与常见版式不符", "人工核对")

    # --- 3. 开票日期 ---
    ds = str(rec.get("issue_date") or "")
    if ds:
        if rec.get("_bad_date"):
            add(LEVEL_ERROR, "开票日期", f"开票日期不是合法日期：{ds}", "人工核对")
        else:
            try:
                d = date.fromisoformat(ds)
                if d > today:
                    add(LEVEL_ERROR, "开票日期", f"开票日期 {ds} 晚于当前日期", "核对是否识别错误或票据日期异常")
                elif d < date(2000, 1, 1):
                    add(LEVEL_WARN, "开票日期", f"开票日期 {ds} 过早，疑似识别错误", "人工核对")
            except ValueError:
                add(LEVEL_ERROR, "开票日期", f"开票日期无法解析：{ds}", "人工核对")

    # --- 4. 三值勾稽 ---
    a, t, tot = rec.get("amount"), rec.get("tax"), rec.get("total")
    if isinstance(a, (int, float)) and isinstance(t, (int, float)) and isinstance(tot, (int, float)):
        diff = round(a + t - tot, 2)
        if abs(diff) > SETTINGS["balance_tolerance"]:
            add(LEVEL_ERROR, "金额勾稽",
                f"金额(不含税){a:,.2f} + 税额{t:,.2f} = {a + t:,.2f}，与价税合计{tot:,.2f} 相差 {diff:,.2f}",
                "核对是否漏识别某一行明细或合计数被遮挡")
    elif isinstance(tot, (int, float)) and tot is not None:
        if a is None or t is None:
            add(LEVEL_WARN, "金额勾稽", "金额(不含税)或税额缺失，无法完成勾稽校验", "人工补充后核对")

    # --- 5. 红字/负数 ---
    if isinstance(tot, (int, float)) and tot is not None:
        if tot < 0:
            add(LEVEL_HINT, "价税合计", f"价税合计为负数（{tot:,.2f}），属于红字发票", "按红字发票入账，注意冲减方向")
        elif tot == 0:
            add(LEVEL_WARN, "价税合计", "价税合计为 0.00", "核对是否为零金额发票或识别异常")

    # --- 6. 税号格式与校验位 ---
    for key, label in (("buyer_tax", "购买方税号"), ("seller_tax", "销售方税号")):
        code = str(rec.get(key) or "").strip().upper()
        if not code:
            if key == "seller_tax":
                add(LEVEL_WARN, label, "销售方纳税人识别号缺失", "人工补录")
            continue
        if not re.fullmatch(r"[0-9A-Z]{15,20}", code):
            add(LEVEL_ERROR, label, f"税号格式非法：{code}", "人工核对")
            continue
        if len(code) not in (15, 17, 18, 20):
            add(LEVEL_WARN, label, f"税号长度 {len(code)} 位较为罕见（常见 15/17/18/20 位）", "人工核对")
        res = uscc_check(code)
        if res is False:
            add(LEVEL_WARN, label, f"统一社会信用代码校验位不通过：{code}",
                "疑似识别错误（如 0/O、1/I 混淆），建议人工核对")

    # --- 7. 购销同名 ---
    bn, sn = (rec.get("buyer_name") or "").strip(), (rec.get("seller_name") or "").strip()
    if bn and sn and bn == sn:
        add(LEVEL_WARN, "购销双方", "购买方与销售方名称完全相同", "核对是否填错，或属于自开自抵情形")

    # --- 8. 税率 ---
    rate = str(rec.get("tax_rate") or "")
    if rate:
        for r in rate.split("/"):
            if r and r not in VALID_TAX_RATES:
                add(LEVEL_HINT, "税率", f"税率 {r} 不在常用税率集合内", "确认是否为特殊征收率/减免政策")
                break

    # --- 9. 老版发票缺代码 ---
    if "数电" not in itype and not rec.get("invoice_code") and no:
        add(LEVEL_WARN, "发票代码", "未识别到发票代码（非数电发票通常应有 12 位发票代码）", "人工核对")

    return issues


def rollup_status(issues: list[dict]) -> str:
    from .config import STATUS_ERROR, STATUS_OK, STATUS_WARN

    levels = {i["level"] for i in issues}
    if LEVEL_ERROR in levels:
        return STATUS_ERROR
    if LEVEL_WARN in levels or LEVEL_HINT in levels:
        return STATUS_WARN
    return STATUS_OK


def summarize_issues(issues: list[dict], limit: int = 6) -> str:
    """把问题清单压成台账中的一行说明文字。"""
    if not issues:
        return ""
    order = {LEVEL_ERROR: 0, LEVEL_WARN: 1, LEVEL_HINT: 2}
    srt = sorted(issues, key=lambda i: order.get(i["level"], 9))
    parts = [f"[{i['level']}] {i['field']}: {i['message']}" for i in srt[:limit]]
    if len(srt) > limit:
        parts.append(f"……等共 {len(srt)} 条，详见「异常记录」页")
    return "； ".join(parts)
