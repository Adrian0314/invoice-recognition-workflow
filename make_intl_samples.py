# -*- coding: utf-8 -*-
"""
国际商务发票样本生成器
======================
来源与授权说明（重要）
----------------------
目标站点 https://www.realguts.com.tw/invoice 经核实**不提供任何可下载的模板文件**：
  · 首页仅列出 6 个模板条目（缩略图 + 详情页链接），无 .psd/.docx/.pdf/.zip 文件链接
  · 详情页正文为通用介绍文案，两次抓取除标题外逐字一致，"下载"链接为占位符 "#"
  · 该站点对本机 HTTPS 直连被 TLS 重置，二进制文件亦无法获取
因此本脚本**不包含任何下载内容**，而是依据该站点自己列出的 6 个模板类型、
以及其在页面中明确描述的字段结构（公司 Logo 区、客户信息、发票编号、开票日期、
付款方式、商品或服务清单、单价、数量、税额、总金额），重建等效版式后生成虚拟样本。

生成物声明
----------
所有公司名、人名、税号、银行账号均为**虚构**，发票编号带 SAMPLE 标识，
仅供软件测试 / 教学演示使用，不具备任何票据效力。

用法
----
    python make_intl_samples.py
产物
----
    samples_intl/            生成的发票 PDF + 样本清单.csv
    invoices_input/          同步一份，供工作流直接处理
"""
from __future__ import annotations

import csv
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "samples_intl"
INBOX = ROOT / "invoices_input"

try:
    import pymupdf
except ImportError:
    import fitz as pymupdf

sys.path.insert(0, str(ROOT))
from invoice_wf import setup_console  # noqa: E402
from invoice_wf.code_rules import e_invoice_no, invoice_code  # noqa: E402
from invoice_wf.csv_util import text_cell, write_csv  # noqa: E402

setup_console()

# 三家销售方都在中国台湾，地区码取全国行政区划代码的台湾省 7100
TW_REGION = "台湾"

PAGE_W, PAGE_H = 595, 842
M = 45                      # 页边距


# --------------------------------------------------------------------------
# 字体：复用主样本生成器验证过的方式（内置中文字体，拉丁字符同样可用）
# --------------------------------------------------------------------------
def _font_spec():
    """
    优先使用系统里的真实字体。

    为什么不直接用 PyMuPDF 内置的 china-s：内置 CJK 字体在**宽度度量**上把
    拉丁字符也当作全宽（约 1em）处理，于是 "USD 6,405.00" 会被算成 12em 宽，
    insert_textbox 认为放不下就把整段文字**静默丢弃**（表现为金额、标题凭空消失），
    get_text_length 也会因此偏大导致右对齐错位。真实字体按比例度量，这些问题一次消失。
    """
    for fp in ("C:/Windows/Fonts/Deng.ttf", "C:/Windows/Fonts/simhei.ttf",
               "C:/Windows/Fonts/msyh.ttc"):
        if not Path(fp).exists():
            continue
        try:
            doc = pymupdf.open()
            pg = doc.new_page()
            rc = pg.insert_textbox(pymupdf.Rect(20, 20, 320, 60),
                                   "中文 Abc 123 发票 INVOICE",
                                   fontsize=10, fontname="cn", fontfile=fp)
            doc.close()
            if rc >= 0:
                return "cn", fp
        except Exception:
            continue
    return "china-s", None      # 兜底（会有上面的宽度度量问题）


FN, FF = _font_spec()

_ALIGN = {"l": pymupdf.TEXT_ALIGN_LEFT,
          "c": pymupdf.TEXT_ALIGN_CENTER,
          "r": pymupdf.TEXT_ALIGN_RIGHT}

_OVERFLOW = []


def put(page, x0, x1, baseline, s, size=9.0, color=(0, 0, 0), align="l"):
    """
    在 [x0, x1] 区间内、基线 baseline 处放一行文本。

    这里刻意用 insert_textbox 而不是 insert_text + 手算宽度：
    pymupdf.get_text_length(fontname="china-s") 对中英混排的宽度估算明显偏大，
    手算右对齐会让整块文字错位、互相压字。交给排版引擎自己算才准。
    """
    rect = pymupdf.Rect(x0, baseline - size * 1.08, x1, baseline + size * 0.72)
    rc = page.insert_textbox(rect, str(s), fontsize=size, fontname=FN, fontfile=FF,
                             color=color, align=_ALIGN[align])
    if rc < 0:
        _OVERFLOW.append((str(s)[:40], round(x1 - x0), round(size, 1)))


def hair(page, x1, y1, x2, y2, color=(0.80, 0.83, 0.87), w=0.5):
    page.draw_line(pymupdf.Point(x1, y1), pymupdf.Point(x2, y2), color=color, width=w)


def fill_rect(page, x, y, w, h, color, radius=None):
    page.draw_rect(pymupdf.Rect(x, y, x + w, y + h), color=None, fill=color, radius=radius)


def money(v: float, cur: str) -> str:
    # 金额数字放在货币代码之前：解析器要求「合计金额/合计税额」这类中文标签后面紧跟数字，
    # 若写成 "USD 6,100.00" 则标签与数字之间夹了货币代码，取不到数值。
    return f"{v:,.2f} {cur}"


L, R = M, PAGE_W - M          # 内容区左右边界 45 / 550
MID = 330                     # 左右分栏位置


# --------------------------------------------------------------------------
# 各区块绘制
# --------------------------------------------------------------------------
def draw_header(page, d, y):
    accent = d["accent"]
    if d["logo"]:
        fill_rect(page, L, y - 4, 34, 34, accent, radius=0.12)
        put(page, L, L + 34, y + 17, d["seller_short"][:1], size=16,
            color=(1, 1, 1), align="c")
        tx = L + 44
        put(page, tx, MID + 60, y + 11, d["seller_name"], size=11.5)
        put(page, tx, MID + 60, y + 24, d["seller_en"], size=8,
            color=(0.42, 0.46, 0.52))
        put(page, tx, MID + 60, y + 35, d["seller_addr"], size=7.4,
            color=(0.55, 0.58, 0.63))
    else:
        tx = L
        put(page, tx, MID + 40, y + 11, d["seller_name"], size=12)
        put(page, tx, MID + 40, y + 24, d["seller_en"], size=8,
            color=(0.42, 0.46, 0.52))

    put(page, MID + 70, R, y + 12, d["title"], size=19, color=accent, align="r")
    put(page, MID + 70, R, y + 26, d["title_en"], size=8.2,
        color=(0.45, 0.49, 0.55), align="r")

    y += 44
    hair(page, L, y, R, y, color=accent, w=1.4)
    return y + 15


def draw_meta_and_parties(page, d, y):
    # ---- 左：购销双方 ----
    # 顺序必须是「购买方名称 → 购买方税号 → 销售方名称 → 销售方税号」。
    # 解析器用「购买方/销售方」之后遇到的第一个「纳税人识别号」做配对，
    # 顺序颠倒会张冠李戴（把对方的税号配到自己身上）。
    put(page, L, MID, y, "开票双方 / PARTIES", size=8, color=(0.45, 0.49, 0.55))
    yy = y + 13
    # 标签措辞有两个约束（都来自解析器的取值边界）：
    #  ① 税号标签必须是纯中文「纳税人识别号」——前面若带英文 "Tax ID /"，
    #     名称字段的终止符匹配不到它，会把 "TaxID/" 一起并进名称里；
    #  ② 紧跟在税号后面的那一行，标签必须以中文开头——否则税号的
    #     [0-9A-Z]{15,20} 会把下一个标签的首字母吃掉（如 "Address" 的 A）。
    for label, value, size in (
        ("Bill To / 购买方名称", d["buyer_name"], 9.2),
        ("纳税人识别号", d["buyer_taxid"], 7.4),
        ("地址 / Address", d["buyer_addr"], 7.2),
        ("联系人 / Attn", d["buyer_contact"], 7.2),
        ("Seller / 销售方名称", d["seller_name"], 9.2),
        ("纳税人识别号", d["seller_tax"], 7.4),
        ("备注 / Remarks", "虚拟样票 · 仅供软件测试", 7.2),
    ):
        put(page, L, L + 116, yy, label, size=7.2, color=(0.52, 0.56, 0.62))
        put(page, L + 120, MID - 8, yy, value, size=size)
        yy += 12.5
    left_end = yy

    # ---- 右：发票要素 ----
    # 中文标签必须紧贴取值（中间不能夹英文或冒号），否则解析器的
    # 「标签 + 可选冒号 + 取值」模式匹配不到。所以英文标签一律写在中文之前。
    rows = [
        ("Invoice No. / 发票号码", d["invoice_no"]),
        ("Invoice Code / 发票代码", d["invoice_code"]),
        ("Invoice Date / 开票日期", d["date"]),
        ("Due Date / 付款期限", d["due"]),
        ("Payment / 付款方式", d["payment"]),
        ("Currency / 币别", d["currency"]),
        ("Reference / 样票编号", d["inv_no"]),
    ]
    yy = y
    for k, v in rows:
        put(page, MID, MID + 116, yy, k, size=7.2, color=(0.52, 0.56, 0.62))
        put(page, MID + 120, R, yy, v, size=8.2, align="r")
        yy += 12.2
        hair(page, MID, yy - 4, R, yy - 4, color=(0.90, 0.92, 0.95))

    y = max(left_end, yy) + 14
    if d.get("band"):
        fill_rect(page, L, y - 13, R - L, 19, d["band_color"], radius=0.08)
        put(page, L, R, y, d["band"], size=9.2, color=(1, 1, 1), align="c")
        y += 26
    return y + 6


def draw_items(page, d, y):
    heads = [("#", L + 4, L + 16, "l", 7.8),
             ("项目 / Description", L + 22, MID - 10, "l", 7.8),
             ("数量 / Qty", 330, 390, "r", 7.8),
             ("单价 / Price", 395, 470, "r", 7.8),
             ("金额 / Amount", 475, R, "r", 7.8)]
    fill_rect(page, L, y - 13, R - L, 18, (0.93, 0.95, 0.98))
    for name, x0, x1, al, sz in heads:
        put(page, x0, x1, y, name, size=sz, color=(0.28, 0.33, 0.40), align=al)
    y += 6
    hair(page, L, y, R, y)
    y += 3

    for i, (desc, qty, price) in enumerate(d["items"], 1):
        y += 14
        put(page, L + 4, L + 16, y, str(i), size=8.4, color=(0.45, 0.49, 0.55))
        put(page, L + 22, MID - 10, y, desc, size=8.6)
        put(page, 330, 390, y, f"{qty:g}", size=8.6, align="r")
        put(page, 395, 470, y, f"{price:,.2f}", size=8.6, align="r")
        put(page, 475, R, y, f"{qty * price:,.2f}", size=8.6, align="r")
        hair(page, L, y + 6.5, R, y + 6.5, color=(0.91, 0.93, 0.96))

    return y + 22


def draw_totals(page, d, y):
    # 标签用中文「合计金额 / 合计税额 / 价税合计」，且金额不带货币前缀
    # （写成 "USD 6,100.00" 会让标签与数字之间夹入货币代码，解析取不到值）；
    # 币别已在右上「币别 / Currency」处标明，不丢信息。
    lab0, lab1 = 250, 470
    rows = [
        ("Subtotal / 合计金额", d["subtotal"], False),
        (f"Tax {d['tax_rate']:g}% / 合计税额", d["tax"], False),
        ("Total Due / 价税合计", d["total"], True),
    ]

    for name, val, strong in rows:
        if strong:
            # 分隔线要落在「上一行文字下缘」与「本行文字上缘」之间，
            # 直接写 y-5 会正好压在本行金额的字身上（看起来像删除线）
            y += 4
            hair(page, lab0, y - 10, R, y - 10, color=d["accent"], w=1.1)
        put(page, lab0, lab1, y, name, size=9.2 if strong else 8.4,
            color=d["accent"] if strong else (0.32, 0.36, 0.42))
        put(page, lab1 + 5, R, y, money(val, d["currency"]),
            size=9.8 if strong else 8.4,
            color=d["accent"] if strong else (0.15, 0.18, 0.22), align="r")
        y += 15
    return y + 10


def draw_footer(page, d, y):
    y += 4
    hair(page, L, y, R, y)
    y += 13

    if d.get("terms"):
        put(page, L, R, y, "条款与条件 / Terms & Conditions", size=8.2,
            color=(0.35, 0.39, 0.45))
        y += 12
        for line in d["terms"]:
            put(page, L, R, y, line, size=7.2, color=(0.48, 0.52, 0.58))
            y += 10
        y += 8

    if d.get("speaker"):
        put(page, L, R, y, "演讲信息 / Speaker Details", size=8.2,
            color=(0.35, 0.39, 0.45))
        y += 12
        for k, v in d["speaker"]:
            put(page, L, R, y, f"{k}：{v}", size=7.4, color=(0.48, 0.52, 0.58))
            y += 10
        y += 8

    # 底部信息块：左侧银行/联络信息，右侧签章区（两区横向错开，互不压字）
    base = max(y + 16, PAGE_H - 118)
    put(page, L, 380, base, f"收款账户 / Bank：{d['bank']}", size=7.4,
        color=(0.48, 0.52, 0.58))
    put(page, L, 380, base + 11, f"开户行 / Bank Name：{d['bank_name']}", size=7.4,
        color=(0.48, 0.52, 0.58))
    put(page, L, 380, base + 22, f"电话 / Tel：{d['seller_tel']}", size=7.4,
        color=(0.48, 0.52, 0.58))
    put(page, L, 380, base + 33, f"邮箱 / Email：{d['seller_mail']}", size=7.4,
        color=(0.48, 0.52, 0.58))

    if d.get("sign"):
        put(page, 400, R, base, "授权签章 / Authorized Signature", size=7.4,
            color=(0.55, 0.58, 0.63), align="r")
        hair(page, 400, base + 24, R, base + 24, color=(0.62, 0.65, 0.70))

    put(page, L, R, PAGE_H - 26,
        "本文件为软件测试用虚拟样本（SAMPLE），全部信息虚构，不具票据效力",
        size=7, color=(0.62, 0.65, 0.70), align="c")


def render(d, path: Path):
    _OVERFLOW.clear()
    doc = pymupdf.open()
    page = doc.new_page(width=PAGE_W, height=PAGE_H)
    y = draw_header(page, d, M - 8)
    y = draw_meta_and_parties(page, d, y)
    y = draw_items(page, d, y)
    y = draw_totals(page, d, y)
    draw_footer(page, d, y)

    # 必须做字体子集化：完整嵌入一个中文字体会让每份 PDF 涨到 5~15MB、保存耗时十几秒
    # （这正是之前生成"卡住"的原因）。子集化后只保留用到的字形，体积降到 25KB 上下。
    try:
        doc.subset_fonts()
    except Exception:
        pass

    doc.save(str(path), garbage=3, deflate=True)
    doc.close()
    return list(_OVERFLOW)


# --------------------------------------------------------------------------
# 6 种模板类型（依据站点列出的模板名称）+ 12 份样本数据
# --------------------------------------------------------------------------
ACCENT_BLUE = (0.12, 0.31, 0.47)
ACCENT_TEAL = (0.06, 0.43, 0.38)
ACCENT_AMBER = (0.62, 0.42, 0.05)
ACCENT_PLUM = (0.42, 0.20, 0.42)
ACCENT_SLATE = (0.28, 0.32, 0.38)
ACCENT_RED = (0.66, 0.20, 0.16)

TEMPLATES = {
    "T1": dict(name="专业发票", title="INVOICE", title_en="商业发票 / Commercial Invoice",
               accent=ACCENT_BLUE, logo=True, sign=True),
    "T2": dict(name="代理发票", title="AGENCY INVOICE",
               title_en="代理服务发票 / Agency Service Invoice",
               accent=ACCENT_TEAL, logo=True, sign=True),
    "T3": dict(name="估算发票", title="ESTIMATE", title_en="估价单 / Estimate",
               accent=ACCENT_AMBER, logo=True, sign=False,
               band="估价单 · 非付款凭证 / ESTIMATE — NOT A REQUEST FOR PAYMENT",
               band_color=ACCENT_AMBER),
    "T4": dict(name="条款发票", title="INVOICE", title_en="含条款与条件 / With Terms",
               accent=ACCENT_PLUM, logo=True, sign=True),
    "T5": dict(name="简式发票", title="INVOICE", title_en="简式发票 / Simplified Invoice",
               accent=ACCENT_SLATE, logo=False, sign=False),
    "T6": dict(name="演讲者发票", title="SPEAKER INVOICE",
               title_en="演讲服务发票 / Speaker Invoice",
               accent=ACCENT_RED, logo=True, sign=True),
}

# --------------------------------------------------------------------------
# 统一社会信用代码：按 GB 32100-2015 生成合法校验位
# 说明：为了让样本能通过校验规则（税号须为 15~20 位字母数字且校验位正确），
# 这里使用大陆统一社会信用代码格式。台湾地区公司实际用的是 8 位「统一编号」，
# 若将来要还原成真实格式，需要放宽 validator 里的税号格式规则。
# --------------------------------------------------------------------------
_USCC_CHARS = "0123456789ABCDEFGHJKLMNPQRTUWXY"
_USCC_W = [1, 3, 9, 27, 19, 26, 16, 17, 20, 29, 25, 13, 8, 24, 10, 30, 28]


def uscc(body17: str) -> str:
    total = sum(_USCC_CHARS.index(c) * w for c, w in zip(body17, _USCC_W))
    c = 31 - total % 31
    return body17 + _USCC_CHARS[0 if c == 31 else c]


_SELLERS = {
    "A": dict(seller_name="环宇国际贸易股份有限公司",
              seller_en="Universal Trade International Co., Ltd.",
              seller_short="环", seller_addr="台北市信义区松高路 88 号 12 楼",
              seller_tax=uscc("91310115MA1H8WXYQ"),
              bank="6220-8800-1122-3344", bank_name="华南商业银行 信义分行",
              seller_tel="+886-2-8788-1234", seller_mail="billing@universal-trade.example"),
    "B": dict(seller_name="鼎新资讯科技股份有限公司",
              seller_en="Dingxin Information Technology Co., Ltd.",
              seller_short="鼎", seller_addr="新竹市东区光复路二段 101 号 5 楼",
              seller_tax=uscc("91320594MA1H8WXQ"),
              bank="8080-5566-7788-9900", bank_name="玉山商业银行 新竹分行",
              seller_tel="+886-3-571-2200", seller_mail="invoice@dingxin-tech.example"),
    "C": dict(seller_name="明诚专业顾问有限公司",
              seller_en="Mingcheng Professional Consulting Ltd.",
              seller_short="明", seller_addr="台中市西区台湾大道二段 501 号 9 楼",
              seller_tax=uscc("91350200MA1H8WXP"),
              bank="7000-1234-5678-9012", bank_name="第一商业银行 台中分行",
              seller_tel="+886-4-2320-8899", seller_mail="ar@mingcheng-consulting.example"),
}

_BUYERS = [
    ("光点设计工作室", "Brightpoint Design Studio", "台北市大安区复兴南路一段 200 号 7 楼",
     "陈映洁", uscc("91310108MA01ABCDE")),
    ("宏远建设股份有限公司", "Hongyuan Construction Co., Ltd.", "新北市板桥区文化路二段 300 号 18 楼",
     "林建宏", uscc("91320100MA01ABCDG")),
    ("星海教育基金会", "Xinghai Education Foundation", "高雄市苓雅区四维三路 6 号 4 楼",
     "黄子瑜", uscc("91440300MA01ABCDF")),
]

# (模板, 编号, 日期, 到期, 卖方, 买方索引, 币别, 税率, 付款方式, 明细, 额外费用, 条款/演讲, 场景)
RECORDS = [
    ("T1", "INV-2026-0001-SAMPLE", "2026年06月03日", "2026年07月03日", "A", 0, "USD", 5.0,
     "T/T 电汇",
     [("企业官网改版设计与前端开发 Contract Development", 1, 4800.00),
      ("UI 组件库交付与文档 Document Package", 2, 650.00)],
     None, None, "软件开发服务"),

    ("T1", "INV-2026-0002-SAMPLE", "2026年06月21日", "2026年07月21日", "B", 1, "CNY", 6.0,
     "银行转账",
     [("服务器设备采购 Server Hardware", 6, 3200.00),
      ("三年上门维保服务 On-site Maintenance", 1, 10500.00)],
     None, None, "设备采购"),

    ("T2", "AGC-2026-0011-SAMPLE", "2026年07月02日", "2026年08月01日", "C", 2, "USD", 5.0,
     "T/T 电汇",
     [("海外客户代理开发 Agency Development", 1, 3000.00),
      ("成交佣金 Commission (3%)", 1, 3650.00)],
     ("垫付款 / Advance", 420.00), None, "代理服务费"),

    ("T2", "AGC-2026-0012-SAMPLE", "2026年07月18日", "2026年08月17日", "C", 0, "TWD", 5.0,
     "支票",
     [("展会代理摊位费 Exhibitor Booth Agency", 2, 48000.00),
      ("现场执行人力 On-site Staffing", 12, 2600.00)],
     None, None, "佣金结算"),

    ("T3", "EST-2026-0021-SAMPLE", "2026年07月25日", "2026年08月24日", "B", 1, "CNY", 6.0,
     "待确认",
     [("ERP 系统导入（一期）ERP Phase 1", 1, 86000.00),
      ("数据迁移与培训 Migration & Training", 1, 24000.00)],
     None, None, "项目预估价"),

    ("T3", "EST-2026-0022-SAMPLE", "2026年08月05日", "2026年09月04日", "C", 2, "TWD", 5.0,
     "待确认",
     [("年度品牌顾问费（预估）Brand Consulting", 12, 15000.00)],
     None, None, "季度预算估算"),

    ("T4", "INV-2026-0031-SAMPLE", "2026年08月12日", "2026年09月11日", "B", 0, "CNY", 6.0,
     "银行转账",
     [("机房年度维保 Annual Data-center Support", 1, 42000.00),
      ("7×24 监控值守 Monitoring Service", 12, 1800.00)],
     None,
     ["1. 本发票所列服务自开票日起提供，服务期内如需变更应以书面确认。",
      "2. 付款期限为开票后 30 日内，逾期未付按日万分之五计收违约金。",
      "3. 本发票为软件测试用虚拟样本，不构成任何真实债权债务关系。"],
     "年度维保"),

    ("T4", "INV-2026-0032-SAMPLE", "2026年08月28日", "2026年09月27日", "A", 1, "USD", 5.0,
     "T/T 电汇",
     [("跨境海运整柜 Freight (FCL)", 3, 2750.00),
      ("报关与仓储 Customs & Storage", 1, 1180.00)],
     None,
     ["1. 运费以实际装船日运价为准，汇率波动超过 3% 时双方另行议定。",
      "2. 货物交付后 7 日内为异议期，逾期视为验收合格。",
      "3. 本发票为软件测试用虚拟样本，不构成任何真实债权债务关系。"],
     "物流服务"),

    ("T5", "INV-2026-0041-SAMPLE", "2026年09月02日", "2026年09月02日", "C", 2, "TWD", 5.0,
     "现金",
     [("办公文具与耗材 Office Supplies", 1, 3860.00)],
     None, None, "办公用品"),

    ("T5", "INV-2026-0042-SAMPLE", "2026年09月06日", "2026年09月21日", "C", 0, "CNY", 6.0,
     "微信支付",
     [("技术文档中英翻译 Translation", 46000, 0.28)],
     None, None, "翻译服务"),

    ("T6", "SPK-2026-0051-SAMPLE", "2026年09月08日", "2026年09月30日", "B", 1, "CNY", 6.0,
     "银行转账",
     [("主题演讲《工业数据治理》Keynote", 1, 28000.00),
      ("差旅补贴 Travel Allowance", 1, 3200.00)],
     None, None, "主题演讲"),

    ("T6", "SPK-2026-0052-SAMPLE", "2026年09月10日", "2026年10月10日", "A", 2, "TWD", 5.0,
     "T/T 电汇",
     [("两天企业内训课程 In-house Training (2d)", 2, 22000.00),
      ("教材印制与授权 Material License", 1, 6000.00)],
     None, None, "培训课程"),
]

SPEAKER_INFO = {
    "主题演讲": [("演讲人 / Speaker", "周彦廷 · 首席架构师"),
             ("演讲主题 / Topic", "工业数据治理与实时湖仓实践"),
             ("时间地点 / When & Where", "2026-09-26 · 台北国际会议中心 3F"),
             ("时长 / Duration", "90 分钟（含问答）")],
    "培训课程": [("讲师 / Lecturer", "周彦廷 · 首席架构师"),
             ("课程名称 / Course", "数据平台架构设计实战（2 日）"),
             ("时间地点 / When & Where", "2026-10-15 ~ 10-16 · 客户内训教室"),
             ("人数 / Attendees", "24 人")],
}


def build_records():
    out = []
    for idx, (tpl, no, date, due, sk, bi, cur, rate, pay, items, extra, terms, scene) in \
            enumerate(RECORDS, 1):
        d = dict(TEMPLATES[tpl])
        d.update(_SELLERS[sk])
        bn, ben, baddr, bcontact, btax = _BUYERS[bi]
        d.update(buyer_name=bn, buyer_en=ben, buyer_addr=baddr,
                 buyer_contact=bcontact, buyer_taxid=btax)
        # inv_no 保留为「样票编号 / Reference」；发票代码与发票号码按官方规则生成
        d.update(inv_no=no, date=date, due=due, payment=pay, currency=cur,
                 tax_rate=rate, terms=terms)

        # 12 位发票代码 = 0 + 台湾(7100) + 年度 + 批次 + 票种(11 电子普票)
        # 20 位发票号码 = 年份 + 省级(71) + 渠道(1) + 15 位顺序码
        # 批次位按票递增，使 12 份样本的代码两两不同，便于逐份核对
        year = int(date[:4])
        d.update(invoice_no=e_invoice_no(TW_REGION, year, idx),
                 invoice_code=invoice_code(TW_REGION, year, idx))

        # 附加项（如代理垫付款）并入明细行，让「合计金额 + 合计税额 = 价税合计」口径一致
        rows = list(items)
        if extra:
            rows.append((extra[0], 1, round(extra[1], 2)))
        d["items"] = rows

        subtotal = round(sum(q * p for _, q, p in rows), 2)
        tax = round(subtotal * rate / 100.0, 2)
        d.update(subtotal=subtotal, tax=tax, total=round(subtotal + tax, 2))
        d["speaker"] = SPEAKER_INFO.get(scene)
        d["template_key"] = tpl
        d["template_name"] = TEMPLATES[tpl]["name"]
        d["scene"] = scene
        d["seq"] = idx
        out.append(d)
    return out


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    INBOX.mkdir(parents=True, exist_ok=True)
    records = build_records()

    manifest = []
    overflows = []
    for d in records:
        fname = f"{d['template_key']}-{d['seq']:02d}_{d['template_name']}_{d['scene']}.pdf"
        path = OUT / fname
        of = render(d, path)
        if of:
            overflows.append((fname, of))
        shutil.copyfile(path, INBOX / fname)
        manifest.append({
            "文件名": fname, "模板类型": d["template_name"], "版式": d["title_en"],
            "发票号码": d["invoice_no"], "发票代码": d["invoice_code"],
            "样票编号": d["inv_no"], "开票日期": d["date"], "到期日": d["due"],
            "卖方": d["seller_name"], "卖方税号": d["seller_tax"],
            "买方": d["buyer_name"], "买方税号": d["buyer_taxid"],
            "币别": d["currency"], "合计金额": d["subtotal"],
            "税率": f"{d['tax_rate']:g}%", "合计税额": d["tax"], "价税合计": d["total"],
            "项目数": len(d["items"]), "付款方式": d["payment"],
        })
        print(f"  {fname}   {d['inv_no']}   {d['currency']} {d['total']:>12,.2f}")

    csv_path = OUT / "样本清单.csv"
    # 发票号码/发票代码/税号是纯数字长串，按文本导出，
    # 否则 Excel 打开会丢前导 0、20 位号码变科学计数法
    fields = list(manifest[0].keys())
    write_csv(csv_path, fields,
              [[text_cell(row[k]) for k in fields] for row in manifest])

    print()
    if overflows:
        print("[警告] 以下文本超出可用宽度，可能被截断：")
        for fn, items in overflows:
            for text, w, sz in items:
                print(f"   {fn}  「{text}」 可用宽 {w}pt / 字号 {sz}")
        print()
    else:
        print("版式自检：所有文本均在可用宽度内，无截断 ✓")
        print()

    print(f"模板类型 : {len(TEMPLATES)} 种（{'、'.join(v['name'] for v in TEMPLATES.values())}）")
    print(f"生成份数 : {len(records)} 份（每种模板 2 份）")
    print(f"输出目录 : {OUT}")
    print(f"同步副本 : {INBOX}")
    print(f"样本清单 : {csv_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
