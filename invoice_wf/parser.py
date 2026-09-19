# -*- coding: utf-8 -*-
"""
字段解析层
----------
把发票原始文本解析成结构化字段。

兼容版式：
  1. 数电发票（全电发票）  —— 电子发票（普通发票）/（增值税专用发票）
  2. 增值税电子普通发票 / 电子专用发票
  3. 纸质增值税专用发票 / 普通发票（PDF 扫描或电子版）
  4. 一个 PDF 内含多张发票（连打/合并导出）——自动拆分，每张发票一行

容错设计：
  - 先做 NFKC 归一化（全角→半角、￥→¥、（）→() ），再同时生成
    「按行」与「整行去空格」两份视图，兼顾表格版式与竖直排版两种抽取顺序
  - 每个字段都记录 evidence（命中的原文片段），便于人工复核
"""
from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime

MONEY_STR = r"-?\d[\d,]*\.\d{1,2}"

# 匹配 % 前的整个数字串（必须从「非数字」处起算，才能拿到完整的粘连 token）
_RATE_RE = re.compile(r"(?<![\d.])(\d[\d.]*)%")
# 金额前缀：整数 + 恰好两位小数
_MONEY_PREFIX = re.compile(r"^\d[\d,]*\.\d{2}")


def _strip_money_prefix(tok: str) -> str:
    """从粘连 token 左侧逐段剥离「两位小数的金额」，剩下的就是税率。"""
    rest = tok
    while True:
        m = _MONEY_PREFIX.match(rest)
        if not m:
            break
        nxt = rest[m.end():]
        if not nxt:          # token 全是金额，说明没有税率
            break
        rest = nxt
    return rest


def extract_rates(one: str) -> list[str]:
    """
    从压平文本中提取税率。

    难点：PDF 抽取时各列会粘在一起，例如
        数量1 + 金额1000.00 + 金额1000.00 + 税率6%
        ->  "11000.001000.006%"
    此时 % 前面的 token 是 "11000.001000.006"。
    由于金额固定两位小数，从左往右逐段剥掉「整数.两位小数」：
        "11000.001000.006" -> "1000.006" -> "6"
    剩下 "6" 就是税率。税率若是 1.5% 这类一位小数，剥离后会原样保留。
    """
    out: list[str] = []
    for m in _RATE_RE.finditer(one):
        val = _strip_money_prefix(m.group(1)).strip(".")
        if not val or len(val) > 5:
            continue
        try:
            f = float(val)
        except ValueError:
            continue
        if not (0 <= f <= 100):
            continue
        s = f"{f:g}%"
        if s not in out:
            out.append(s)
    return out

# 名称字段的终止符（防止把相邻单元格的内容一起抓进来）
NAME_STOP = (
    r"(?=销售方|购买方|销货方|购货方|统一社会信用代码|纳税人识别号|税号|"
    r"项目名称|货物或应税|规格型号|单位|数量|单价|金额|税率|征收率|价税合计|"
    r"合计|备注|开票人|收款人|复核人|地址|开户行|电话|账号|$)"
)

def detect_type(one: str) -> str:
    """
    识别发票版式。注意：入参必须是 NFKC 归一化后的文本（全角括号已变为半角）。
    判断顺序很重要——专用发票要排在普通发票之前。
    """
    if "数电" in one or "全电" in one:
        return "数电发票"
    if "电子发票" in one and "专用发票" in one:
        return "数电发票(增值税专用发票)"
    if "电子发票" in one and "普通发票" in one:
        return "数电发票(普通发票)"
    if "增值税电子专用发票" in one:
        return "电子专用发票"
    if "增值税电子普通发票" in one:
        return "电子普通发票"
    if "机动车销售统一发票" in one:
        return "机动车销售统一发票"
    if "二手车销售统一发票" in one:
        return "二手车销售统一发票"
    if "增值税专用发票" in one:
        return "增值税专用发票"
    if "增值税普通发票" in one:
        return "增值税普通发票"
    if "发票" in one:
        return "电子发票(其他)"
    return "未知类型"


# ==========================================================================
# 文本归一化
# ==========================================================================
def norm_text(raw: str) -> str:
    t = raw.replace("\r\n", "\n").replace("\r", "\n")
    for ch in ("\u00a0", "\u3000", "\u200b", "\ufeff"):
        t = t.replace(ch, " ")
    return unicodedata.normalize("NFKC", t)


def to_oneline(t: str) -> str:
    """去掉所有空白，压成一行——用于跨单元格/竖直排版的关键词定位。"""
    return re.sub(r"\s+", "", t)


def to_lines(t: str) -> list[str]:
    out = []
    for ln in t.split("\n"):
        ln = re.sub(r"[ \t]+", "", ln).strip()
        if ln:
            out.append(ln)
    return out


def to_money(s: str | None) -> float | None:
    if not s:
        return None
    s = s.replace(",", "").replace("¥", "").replace("￥", "").strip()
    try:
        return round(float(s), 2)
    except ValueError:
        return None


def _clean_name(v: str | None) -> str:
    if not v:
        return ""
    v = re.sub(r"^[:：\s]+", "", v)
    v = re.split(
        r"统一社会信用代码|纳税人识别号|销售方|购买方|销货方|购货方|项目名称|"
        r"货物或应税|规格型号|地址|开户行|电话|账号|备注|开票人",
        v,
    )[0]
    v = re.sub(r"[:：\s]+$", "", v)
    v = re.sub(r"^[（(].*?[)）]", "", v)          # 去掉「（小写）」之类前缀
    return v.strip(" :：,，、")


# ==========================================================================
# 主入口
# ==========================================================================
def parse_invoice(text: str, file_name: str = "") -> list[dict]:
    """
    返回发票记录列表（一个文件可能含多张发票，故为 list）。
    """
    normalized = norm_text(text)
    segments = _split_multi_invoice(normalized)
    doc_type = detect_type(to_oneline(normalized))

    records, seen = [], set()
    for idx, seg in enumerate(segments):
        rec = _parse_single(seg, file_name)
        # 段落自身识别不出类型时（例如标题被切掉），回退到整份文件的类型
        if rec.get("invoice_type") in ("未知类型", "电子发票(其他)") and doc_type != "未知类型":
            rec["invoice_type"] = doc_type
        no = rec.get("invoice_no") or ""
        if no and no in seen:
            continue  # 同一文件内解析出重复号码（页码/条码区误命中）→ 丢弃副本
        if no:
            seen.add(no)
        rec["_segment"] = idx + 1
        rec["_segment_total"] = len(segments)
        records.append(rec)

    # 一个文件多张发票时，文件名后加标记，便于追溯
    if len(records) > 1:
        for i, r in enumerate(records, 1):
            r["file_name"] = f"{file_name} [第{i}张/共{len(records)}张]"
    return records


_TITLE_RE = re.compile(r"^(电子发票|全电发票|增值税|机动车销售|二手车销售)")


def _looks_like_title(ln: str) -> bool:
    """判断某行是否为发票标题行（如『电子发票（普通发票）』）。"""
    return bool(_TITLE_RE.match(ln)) and "发票" in ln and "号码" not in ln and "代码" not in ln


def _split_multi_invoice(t: str) -> list[str]:
    """
    把「一个文件含多张发票」的文本拆成多段，保证一张发票一行。

    切分点（仅当当前段落已具备核心字段时才切）：
      1. 新的「发票号码」行
      2. 新的发票标题行（标题在号码之前，必须在标题处就切开，
         否则第 2 张发票会丢掉标题导致版式识别为「其他」）
    """
    per_line = re.sub(r"\s+", "", t)
    hits = [m.start() for m in re.finditer(r"发票号码", per_line)]
    titles = [m.start() for m in re.finditer(r"(电子发票|全电发票|增值税)", per_line)]
    if len(hits) <= 1 and len(titles) <= 1:
        return [t]

    lines = to_lines(t)
    segs: list[list[str]] = []
    cur: list[str] = []
    for ln in lines:
        if cur and _has_core_fields(cur) and ("发票号码" in ln or _looks_like_title(ln)):
            segs.append(cur)
            cur = []
        cur.append(ln)
    if cur:
        segs.append(cur)
    return ["\n".join(s) for s in segs] if len(segs) > 1 else [t]


def _has_core_fields(lines: list[str]) -> bool:
    joined = "".join(lines)
    return ("价税合计" in joined or "税额" in joined) and "开票日期" in joined


def _parse_single(t: str, file_name: str) -> dict:
    lines = to_lines(t)
    flatin = re.sub(r"[ \t]+", "", t)
    one = to_oneline(t)
    ev: dict[str, str] = {}

    rec: dict = {k: "" for k in (
        "invoice_type", "invoice_code", "invoice_no", "issue_date",
        "buyer_name", "buyer_tax", "seller_name", "seller_tax",
        "amount", "tax", "total", "tax_rate", "item_name", "check_code",
        "drawer",
    )}
    rec["file_name"] = file_name

    # ---------------- 发票类型 ----------------
    rec["invoice_type"] = detect_type(one)
    ev["invoice_type"] = rec["invoice_type"]

    # ---------------- 发票号码 / 发票代码 ----------------
    m = re.search(r"发票号码[:：]?\s*(\d{8,20})", one)
    if m:
        rec["invoice_no"] = m.group(1)
        ev["invoice_no"] = m.group(0)
    m = re.search(r"发票代码[:：]?\s*(\d{10,12})", one)
    if m:
        rec["invoice_code"] = m.group(1)
        ev["invoice_code"] = m.group(0)

    # ---------------- 开票日期 ----------------
    m = re.search(r"开票日期[:：]?\s*(\d{4})[年\-/.]?(\d{1,2})[月\-/.]?(\d{1,2})日?", one)
    if m:
        y, mo, d = (int(g) for g in m.groups())
        try:
            rec["issue_date"] = date(y, mo, d).isoformat()
            ev["issue_date"] = m.group(0)
        except ValueError:
            rec["issue_date"] = f"{y:04d}-{mo:02d}-{d:02d}"
            rec["_bad_date"] = True
            ev["issue_date"] = m.group(0)

    # ---------------- 校验码（仅老版电子发票有） ----------------
    m = re.search(r"校验码[:：]?\s*(\d{15,25})", one)
    if m:
        rec["check_code"] = m.group(1)
        ev["check_code"] = m.group(0)

    # ---------------- 购买方 / 销售方 ----------------
    buyer_name, buyer_tax, ev_bn, ev_bt = _party(one, flatin, lines, "购买方")
    seller_name, seller_tax, ev_sn, ev_st = _party(one, flatin, lines, "销售方")
    rec["buyer_name"], rec["buyer_tax"] = buyer_name, buyer_tax
    rec["seller_name"], rec["seller_tax"] = seller_name, seller_tax
    if ev_bn:
        ev["buyer_name"] = ev_bn
    if ev_bt:
        ev["buyer_tax"] = ev_bt
    if ev_sn:
        ev["seller_name"] = ev_sn
    if ev_st:
        ev["seller_tax"] = ev_st

    # ---------------- 金额 / 税额 / 价税合计 ----------------
    amount = tax = total = None

    m = re.search(r"价税合计[^¥￥]{0,60}?[¥￥]?(" + MONEY_STR + r")", one)
    if not m:
        m = re.search(r"[(（]小写[)）][¥￥]?(" + MONEY_STR + r")", one)
    if not m:
        m = re.search(r"价税合计[^0-9]{0,60}(" + MONEY_STR + r")", one)
    if m:
        total = to_money(m.group(1))
        ev["total"] = m.group(0)

    m = re.search(r"合计[¥￥](" + MONEY_STR + r")[¥￥](" + MONEY_STR + r")", one)
    if m:
        amount, tax = to_money(m.group(1)), to_money(m.group(2))
        ev["amount"], ev["tax"] = m.group(0), m.group(0)
    else:
        m = re.search(r"合计金额[¥￥]?(" + MONEY_STR + r")", one)
        if m:
            amount = to_money(m.group(1))
            ev["amount"] = m.group(0)
        m = re.search(r"(?:合计税额|税额合计)[¥￥]?(" + MONEY_STR + r")", one)
        if m:
            tax = to_money(m.group(1))
            ev["tax"] = m.group(0)
        # 兜底：文本中出现相邻的两个「¥金额 ¥金额」——通常是合计行
        if amount is None and tax is None:
            pairs = re.findall(r"[¥￥](" + MONEY_STR + r")[¥￥](" + MONEY_STR + r")", one)
            if pairs:
                amount, tax = to_money(pairs[-1][0]), to_money(pairs[-1][1])
                ev["amount"], ev["tax"] = "兜底匹配", "兜底匹配"

    # 用价税合计反推缺失项
    if total is not None and tax is not None and amount is None:
        amount = round(total - tax, 2)
        ev["amount"] = "由价税合计-税额推导"
    if total is not None and amount is not None and tax is None:
        tax = round(total - amount, 2)
        ev["tax"] = "由价税合计-金额推导"
    rec["amount"], rec["tax"], rec["total"] = amount, tax, total

    # ---------------- 税率 ----------------
    rates = extract_rates(one)
    if not rates:
        for kw in ("免税", "不征税"):
            if kw in one:
                rates.append(kw)
    rec["tax_rate"] = "/".join(rates)
    if rates:
        ev["tax_rate"] = "、".join(rates)

    # ---------------- 项目名称 ----------------
    items = []
    for ln in lines:
        for it in re.findall(r"\*[^*]{1,20}\*[^*0-9]{0,30}", ln):
            it = it.strip()
            if it.endswith("*"):
                it = it[:-1]
            if it and it not in items:
                items.append(it)
    if items:
        rec["item_name"] = "; ".join(items)[:120]
        ev["item_name"] = rec["item_name"]
    else:
        m = re.search(r"(?:货物或应税劳务、?服务名称|项目名称)[:：]?([^\d¥￥]{2,30})" + NAME_STOP, one)
        if m:
            rec["item_name"] = m.group(1)
            ev["item_name"] = m.group(0)

    # ---------------- 开票人 ----------------
    m = re.search(r"开票人[:：]?([\u4e00-\u9fa5A-Za-z]{2,8})", one)
    if not m:
        m = re.search(r"开票人[:：]?([\u4e00-\u9fa5A-Za-z]{2,8})(?=收款人|复核人|$)", one)
    if m:
        rec["drawer"] = m.group(1)
        ev["drawer"] = m.group(0)

    rec["_evidence"] = ev
    rec["_raw_text"] = t
    return rec


# ==========================================================================
# 购销双方解析（双通道：整行视图 + 按行视图）
# ==========================================================================
def _party(one: str, flat: str, lines: list[str], who: str):
    """返回 (名称, 税号, 证据名称, 证据税号)"""
    is_buyer = who == "购买方"
    other = "销售方" if is_buyer else "购买方"
    aliases = (["购买方", "购货方", "付款方"] if is_buyer
               else ["销售方", "销货方", "收款方"])

    name = tax = ""
    ev_name = ev_tax = ""

    # ---- 通道 A：整行视图（对竖直排版最稳） ----
    for al in aliases:
        if name:
            break
        m = re.search(al + r"(?:信息)?名称[:：]?(.{2,50}?)" + NAME_STOP, one)
        if m:
            name = _clean_name(m.group(1))
            ev_name = m.group(0)
    # 只有「购买方信息 / 销售方信息」块、名称带冒号的情况
    if not name:
        for al in aliases:
            m = re.search(al + r"(?:信息)?[:：]?名称[:：]?(.{2,50}?)" + NAME_STOP, one)
            if m:
                name = _clean_name(m.group(1))
                ev_name = m.group(0)
                break

    # ---- 通道 B：按行视图 ----
    if not name or not tax:
        b_i = s_i = None
        for i, ln in enumerate(lines):
            if b_i is None and any(a in ln for a in ("购买方", "购货方", "付款方")):
                b_i = i
            if s_i is None and any(a in ln for a in ("销售方", "销货方", "收款方")):
                s_i = i
        start = b_i if is_buyer else s_i
        if start is not None:
            if is_buyer and s_i is not None and s_i > start:
                end = s_i
            else:
                end = min(len(lines), start + 8)
            section = lines[start:end]
            if not name:
                nm, ev = _scan_section(section, r"名称[:：]?(.+)$")
                if nm:
                    name = _clean_name(nm)
                    ev_name = ev
            if not tax:
                tx, ev = _scan_section(section, r"(?:统一社会信用代码|纳税人识别号|税号)[/]?[:：]?([0-9A-Z]{15,20})")
                if tx:
                    tax = tx
                    ev_tax = ev

    # ---- 税号通道 A：整行视图（限定在本方区块内，防止窜到对方） ----
    if not tax:
        tax_re = re.compile(
            re.escape(who) + r"(?:信息)?.{0,120}?(?:统一社会信用代码|纳税人识别号|税号)[/]?[:：]?([0-9A-Z]{15,20})"
        )
        m = tax_re.search(one)
        if m:
            cand = m.group(1)
            # 简单防窜：若候选税号在「对方」关键词之后，则丢弃
            pos = one.find(cand)
            opp = one.find(other)
            if not (opp != -1 and pos > opp and one.find(who) < opp):
                tax = cand
                ev_tax = m.group(0)

    return name, tax, ev_name, ev_tax


def _scan_section(section: list[str], pattern: str):
    for ln in section:
        m = re.search(pattern, ln)
        if m:
            return m.group(1), ln
        if "名称" in ln and ":" in ln.replace("：", ":"):
            pass
    return "", ""
