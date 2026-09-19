import random
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

PAGE_W, PAGE_H = A4

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
    total = amount + service_fee

    doc_date = random_date(index)
    serial = f"DEMO-{doc_date.strftime('%Y%m%d')}-{index + 1:03d}"

    story.append(Paragraph("IT 服务费用明细样票", title_style))
    story.append(Paragraph(
        "虚拟样票｜仅供软件测试与教学演示｜不具备发票效力｜不得报销",
        subtitle_style
    ))
    story.append(Spacer(1, 7 * mm))

    info_data = [
        [
            Paragraph("<b>样票编号：</b>" + serial, normal_style),
            Paragraph(
                "<b>生成日期：</b>" + doc_date.strftime("%Y年%m月%d日"),
                normal_style
            ),
        ],
        [
            Paragraph("<b>服务方：</b>" + company, normal_style),
            Paragraph("<b>委托方：</b>" + customer, normal_style),
        ],
        [
            Paragraph("<b>用途：</b>IT服务项目费用明细演示", normal_style),
            Paragraph("<b>结算状态：</b>测试数据", normal_style),
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

    detail_data = [
        [
            Paragraph("<b>序号</b>", normal_style),
            Paragraph("<b>服务内容</b>", normal_style),
            Paragraph("<b>数量</b>", normal_style),
            Paragraph("<b>单位</b>", normal_style),
            Paragraph("<b>单价</b>", normal_style),
            Paragraph("<b>金额</b>", normal_style),
        ],
        [
            "1",
            Paragraph(service_name, normal_style),
            str(quantity),
            "项",
            money(unit_price),
            money(amount),
        ],
        [
            "",
            Paragraph("项目管理及沟通服务费", small_style),
            "1",
            "项",
            money(service_fee),
            money(service_fee),
        ],
        ["", "", "", "", "<b>合计</b>", f"<b>{money(total)}</b>"],
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
        ("ALIGN", (2, 0), (-1, -1), "CENTER"),
        ("ALIGN", (4, 1), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))

    story.append(detail_table)
    story.append(Spacer(1, 8 * mm))

    note_data = [
        [Paragraph("<b>金额说明：</b>", normal_style),
         Paragraph(f"本页金额仅为随机生成的测试数据，合计：{money(total)}", normal_style)],
        [Paragraph("<b>备注：</b>", normal_style),
         Paragraph(
             "本文件为虚拟样票，不属于国家税务机关监制的发票，不含真实税控信息、发票代码、校验码或二维码。",
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
