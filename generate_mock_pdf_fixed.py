# -*- coding: utf-8 -*-
"""
IT 服务费用明细样票生成器 —— 修正版
====================================
基于用户提供的 `generate_mock_pdf.py`，**保持原有版式、样式参数与输出文件名不变**，
只做两类修正：

【一】修渲染缺陷（原脚本自带）
  1. reportlab 的 Table 单元格若直接放裸字符串，会用默认字体 Helvetica 绘制，
     **不会继承自定义 CJK 字体** → 中西文混排的「项」渲染成黑色方块（notdef）。
     故所有单元格统一改用 `Paragraph(..., style)`。
  2. 裸字符串单元格**不解析 reportlab 标记语言** → `"<b>合计</b>"` 会原样打印成字面
     `<b>合计</b>`。改用 Paragraph，并注册字体族让 `<b>` 真正生效。

【二】补中国发票要素字段，使其可被识别
  · 发票号码（20 位）、发票代码（12 位，逐页唯一）—— 位数与结构与官方规则一致，
    规则集中在 invoice_wf/code_rules.py，此处直接复用，避免两处各写一套
  · 开票日期 / 购买方名称 / 销售方名称 / 纳税人识别号（购销双方各一）
  · 合计金额 / 合计税额 / 价税合计（引入 6% 税率，保证三者勾稽相等）
  说明：标签与取值必须紧邻（中间不夹英文或冒号），且取值后面不能紧跟字母数字，
  否则解析器的「标签+取值」边界会失配。税号采用大陆统一社会信用代码格式（18 位、
  校验位正确），以满足校验规则。

产物：虚拟IT公司样票_10张.pdf（A4 × 10 页，文件名与格式同原版）
所有数据均为虚构，仅用于软件测试与教学演示，不具备发票效力。
"""
import os
import random
import sys
from datetime import date, timedelta

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable
)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from invoice_wf.code_rules import e_invoice_no, invoice_code  # noqa: E402

OUTPUT = "虚拟IT公司样票_10张.pdf"

# 尝试注册中文字体
font_candidates = [
    ("SimSun", "C:/Windows/Fonts/simsun.ttc"),
    ("MicrosoftYaHei", "C:/Windows/Fonts/msyh.ttc"),
    ("NotoSansCJK", "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"),
    ("NotoSansCJK", "/usr/share/fonts/truetype/noto/NotoSansCJKsc-Regular.otf"),
]

FONT_NAME = "Helvetica"
for name, path in font_candidates:
    try:
        pdfmetrics.registerFont(TTFont(name, path))
        FONT_NAME = name
        break
    except Exception:
        pass

# 修正一之二：<b> 需要字体族中的 bold 面，否则标记不生效
try:
    pdfmetrics.registerFont(TTFont("SimSun-Bold", "C:/Windows/Fonts/simhei.ttf"))
    pdfmetrics.registerFontFamily(FONT_NAME, normal=FONT_NAME, bold="SimSun-Bold",
                                  italic=FONT_NAME, boldItalic="SimSun-Bold")
except Exception:
    pdfmetrics.registerFontFamily(FONT_NAME, normal=FONT_NAME, bold=FONT_NAME,
                                  italic=FONT_NAME, boldItalic=FONT_NAME)

PAGE_W, PAGE_H = A4


# --------------------------------------------------------------------------
# 统一社会信用代码：GB 32100-2015 校验位
# --------------------------------------------------------------------------
_USCC_CHARS = "0123456789ABCDEFGHJKLMNPQRTUWXY"
_USCC_W = [1, 3, 9, 27, 19, 26, 16, 17, 20, 29, 25, 13, 8, 24, 10, 30, 28]


def uscc(body17):
    total = sum(_USCC_CHARS.index(c) * w for c, w in zip(body17, _USCC_W))
    c = 31 - total % 31
    return body17 + _USCC_CHARS[0 if c == 31 else c]


companies = [
    "深圳星云代码技术工作室",
    "杭州云栈数据服务中心",
    "成都极客方舟软件工作室",
    "广州智链信息技术中心",
    "苏州蓝鲸云计算工作室",
    "武汉数桥科技服务中心",
    "厦门像素引擎软件工作室",
    "南京矩阵数据技术中心",
    "重庆光年网络服务工作室",
    "西安微核信息技术中心",
]

services = [
    ("软件系统维护服务", 1, 680),
    ("网站界面设计服务", 1, 1280),
    ("数据整理与清洗服务", 2, 460),
    ("云服务器配置服务", 1, 980),
    ("小程序功能开发服务", 1, 2380),
    ("数据库优化咨询服务", 3, 520),
    ("IT技术培训服务", 4, 360),
    ("网络安全检测服务", 1, 1680),
    ("数据可视化制作服务", 2, 760),
    ("程序测试服务", 5, 280),
]

customers = [
    "深圳市远景商贸有限公司",
    "广州新图教育咨询有限公司",
    "东莞市简易制造有限公司",
    "佛山市清源文化传播有限公司",
    "珠海市拓维电子商务有限公司",
    "惠州市南风设计有限公司",
    "中山市汇星企业服务有限公司",
    "深圳市青禾供应链有限公司",
]

# 服务方所在城市 → 全国行政区划「地市级」4 位码
# 用于发票代码第 2-5 位与发票号码第 3-4 位（省级取前两位）
# 直辖市用「省级码 + 00」：重庆 5000；计划单列市用自身地市码：深圳 4403、厦门 3502
COMPANY_REGION = {
    "深圳星云代码技术工作室": "4403",
    "杭州云栈数据服务中心": "3301",
    "成都极客方舟软件工作室": "5101",
    "广州智链信息技术中心": "4401",
    "苏州蓝鲸云计算工作室": "3205",
    "武汉数桥科技服务中心": "4201",
    "厦门像素引擎软件工作室": "3502",
    "南京矩阵数据技术中心": "3201",
    "重庆光年网络服务工作室": "5000",
    "西安微核信息技术中心": "6101",
}

TAX_RATE = 0.06


def money(value):
    return f"¥{value:,.2f}"


def random_date(i):
    start = date(2026, 1, 1)
    return start + timedelta(days=random.randint(0, 260))


def draw_watermark(canvas, doc):
    canvas.saveState()
    canvas.setFont(FONT_NAME, 25)
    canvas.setFillColorRGB(0.88, 0.88, 0.88)
    canvas.translate(PAGE_W / 2, PAGE_H / 2)
    canvas.rotate(35)
    canvas.drawCentredString(
        0, 0,
        "虚拟样票 · 仅供测试/教学演示 · 不具备发票效力"
    )
    canvas.restoreState()

    canvas.saveState()
    canvas.setFont(FONT_NAME, 8)
    canvas.setFillColor(colors.HexColor("#666666"))
    canvas.drawCentredString(
        PAGE_W / 2, 12 * mm,
        "虚拟样票｜不得用于报销、结算、入账或其他商业用途"
    )
    canvas.restoreState()


styles = getSampleStyleSheet()

title_style = ParagraphStyle(
    "TitleCN",
    parent=styles["Title"],
    fontName=FONT_NAME,
    fontSize=20,
    leading=26,
    alignment=TA_CENTER,
    textColor=colors.HexColor("#163A5F"),
    spaceAfter=5 * mm,
)

subtitle_style = ParagraphStyle(
    "Subtitle",
    parent=styles["Normal"],
    fontName=FONT_NAME,
    fontSize=10,
    leading=14,
    alignment=TA_CENTER,
    textColor=colors.HexColor("#B00020"),
)

normal_style = ParagraphStyle(
    "NormalCN",
    parent=styles["Normal"],
    fontName=FONT_NAME,
    fontSize=10,
    leading=16,
)

small_style = ParagraphStyle(
    "SmallCN",
    parent=styles["Normal"],
    fontName=FONT_NAME,
    fontSize=8.5,
    leading=13,
)

right_style = ParagraphStyle(
    "RightCN",
    parent=normal_style,
    alignment=TA_RIGHT,
)

center_style = ParagraphStyle(
    "CenterCN",
    parent=normal_style,
    alignment=TA_CENTER,
)

story = []

random.seed()

for index in range(10):
    company = companies[index]
    customer = random.choice(customers)
    service_name, quantity, base_price = random.choice(services)

    # 增加随机性
    quantity = random.randint(1, max(1, quantity + 2))
    unit_price = base_price + random.randint(-80, 180)
    unit_price = max(100, unit_price)
    amount = quantity * unit_price
    service_fee = random.randint(20, 120)
    subtotal = amount + service_fee                      # 合计金额（不含税）
    tax = round(subtotal * TAX_RATE, 2)                  # 合计税额
    total = subtotal + tax                               # 价税合计

    doc_date = random_date(index)
    serial = f"DEMO-{doc_date.strftime('%Y%m%d')}-{index + 1:03d}"
    # 逐页唯一的 20 位发票号码 与 12 位发票代码：
    # 地区码取「服务方所在城市」的全国行政区划地市码，年度取开票日期年份，
    # 批次/顺序码按页递增，保证两份都不会重复
    region4 = COMPANY_REGION[company]
    year = doc_date.year
    inv_no = e_invoice_no(region4, year, index + 1)
    inv_code = invoice_code(region4, year, index + 1)
    # 补：购销双方税号（18 位、校验位正确）
    seller_tax = uscc("9131010" + f"{index:02d}" + "MA01ABCD")
    buyer_tax = uscc("9144030" + f"{customers.index(customer):02d}" + "MA01ABCD")

    story.append(Paragraph("IT 服务费用明细样票", title_style))
    story.append(Paragraph(
        "虚拟样票｜仅供软件测试与教学演示｜不具备发票效力｜不得报销",
        subtitle_style
    ))
    story.append(Spacer(1, 7 * mm))

    # 修正：标签与取值写在同一段内且紧邻（不加冒号以外的分隔），保证解析器能取值
    info_data = [
        [
            Paragraph("发票号码：" + inv_no, normal_style),
            Paragraph("开票日期：" + doc_date.strftime("%Y年%m月%d日"), normal_style),
        ],
        [
            Paragraph("购买方名称：" + customer, normal_style),
            Paragraph("纳税人识别号：" + buyer_tax, normal_style),
        ],
        [
            Paragraph("销售方名称：" + company, normal_style),
            Paragraph("纳税人识别号：" + seller_tax, normal_style),
        ],
        [
            Paragraph("发票代码：" + inv_code, normal_style),
            Paragraph("币别：CNY", normal_style),
        ],
        [
            Paragraph("样票编号：" + serial, normal_style),
            Paragraph("用途：IT服务项目费用明细演示　结算状态：测试数据", normal_style),
        ],
    ]

    info_table = Table(info_data, colWidths=[88 * mm, 88 * mm])
    info_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#AAB7C4")),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F7FAFC")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))

    story.append(info_table)
    story.append(Spacer(1, 7 * mm))

    # 修正：所有单元格改为 Paragraph —— 否则中文走 Helvetica 变方块、<b> 标记原样输出
    detail_data = [
        [
            Paragraph("<b>序号</b>", center_style),
            Paragraph("<b>服务内容</b>", normal_style),
            Paragraph("<b>数量</b>", center_style),
            Paragraph("<b>单位</b>", center_style),
            Paragraph("<b>单价</b>", right_style),
            Paragraph("<b>金额</b>", right_style),
        ],
        [
            Paragraph("1", center_style),
            Paragraph(service_name, normal_style),
            Paragraph(str(quantity), center_style),
            Paragraph("项", center_style),
            Paragraph(money(unit_price), right_style),
            Paragraph(money(amount), right_style),
        ],
        [
            Paragraph("", normal_style),
            Paragraph("项目管理及沟通服务费", small_style),
            Paragraph("1", center_style),
            Paragraph("项", center_style),
            Paragraph(money(service_fee), right_style),
            Paragraph(money(service_fee), right_style),
        ],
        [
            Paragraph("", normal_style),
            Paragraph("", normal_style),
            Paragraph("", normal_style),
            Paragraph("", normal_style),
            Paragraph("<b>价税合计</b>", right_style),
            # ¥ 刻意放在 <b> 之外：加粗面用的是 SimHei，它没有 U+00A5 字形，
            # 放进加粗段里会被渲染成空方块（notdef）。
            Paragraph(f"¥<b>{total:,.2f}</b>", right_style),
        ],
    ]

    detail_table = Table(
        detail_data,
        colWidths=[14 * mm, 65 * mm, 20 * mm, 20 * mm, 32 * mm, 35 * mm],
        rowHeights=[12 * mm, 17 * mm, 17 * mm, 14 * mm],
    )

    detail_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.HexColor("#7D8B99")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DCEAF7")),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#F1F5F9")),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("ALIGN", (2, 0), (2, -1), "CENTER"),
        ("ALIGN", (3, 0), (3, -1), "CENTER"),
        ("ALIGN", (4, 0), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))

    story.append(detail_table)
    story.append(Spacer(1, 8 * mm))

    note_data = [
        [Paragraph("<b>金额合计：</b>", normal_style),
         Paragraph(
             f"合计金额 ¥{subtotal:,.2f}　合计税额 ¥{tax:,.2f}"
             f"　价税合计 ¥{total:,.2f}（税率 {TAX_RATE:.0%}）",
             normal_style
         )],
        [Paragraph("<b>备注：</b>", normal_style),
         Paragraph(
             "本文件为虚拟样票，不属于国家税务机关监制的发票，不含真实税控信息、校验码或二维码。",
             small_style
         )],
        [Paragraph("<b>签署栏：</b>", normal_style),
         Paragraph("服务方：________________　委托方：________________", normal_style)],
    ]

    note_table = Table(note_data, colWidths=[28 * mm, 148 * mm])
    note_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#AAB7C4")),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F7FAFC")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))

    story.append(note_table)
    story.append(Spacer(1, 20 * mm))

    story.append(Paragraph(
        "本页为演示用费用明细单，不得作为发票、收据、报销凭证或付款凭证使用。",
        subtitle_style
    ))

    if index < 9:
        story.append(PageBreak())

doc = SimpleDocTemplate(
    OUTPUT,
    pagesize=A4,
    rightMargin=17 * mm,
    leftMargin=17 * mm,
    topMargin=18 * mm,
    bottomMargin=20 * mm,
    title="虚拟IT公司样票（10张）",
    author="Demo Generator",
)

doc.build(
    story,
    onFirstPage=draw_watermark,
    onLaterPages=draw_watermark
)

print(f"已生成：{OUTPUT}")
