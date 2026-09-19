# -*- coding: utf-8 -*-
"""
生成演示用发票样本（覆盖全部边界场景）
======================================
运行：python make_samples.py
产物目录：samples/  （同时会复制到 invoices_input/ 供直接跑通流程）

覆盖场景：
  01 数电普票·正常              -> 正常入账
  02 数电专票·正常              -> 正常入账
  03 电子普票(老版：12 位代码 + 8 位号码)  -> 正常入账
  04 与 01 发票号码相同         -> 发票级重复，拦截
  05 金额勾稽不平               -> 严重异常
  06 税号校验位错误 + 购销同名   -> 警告
  07 一页含两张发票             -> 一个文件生成两行
  08 扫描件 PDF（无文本层）     -> 待人工录入（OCR 提示）
  09 图片发票 PNG（无文本层）   -> 待人工录入（OCR 提示）
  10 与 01 完全相同的字节副本   -> 文件级重复，拦截
  11 07 的换名副本（同事重新转发）-> 文件级重复，拦截
"""
from __future__ import annotations

import io
import os
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "samples"
INBOX = ROOT / "invoices_input"

try:
    import pymupdf
except ImportError:      # 兼容旧包名
    import fitz as pymupdf

sys.path.insert(0, str(ROOT))
from invoice_wf import setup_console  # noqa: E402
from invoice_wf.code_rules import e_invoice_no, invoice_code  # noqa: E402

setup_console()

from PIL import Image  # noqa: E402

PAGE_W, PAGE_H = 595, 842          # A4


# --------------------------------------------------------------------------
# 字体：优先使用 PyMuPDF 内置中文字体，失败则回退到系统黑体
# --------------------------------------------------------------------------
def _font_spec():
    try:
        doc = pymupdf.open()
        pg = doc.new_page()
        pg.insert_text((50, 50), "中文测试", fontname="china-s", fontsize=10)
        doc.close()
        return "china-s", None
    except Exception:
        pass
    for fp in ("C:/Windows/Fonts/simhei.ttf", "C:/Windows/Fonts/Deng.ttf"):
        if Path(fp).exists():
            return "cjk", fp
    raise RuntimeError("找不到可用的中文字体")


FONTNAME, FONTFILE = _font_spec()


def txt(page, x, y, s, size=9, color=(0, 0, 0)):
    page.insert_text((x, y), s, fontsize=size, fontname=FONTNAME,
                     fontfile=FONTFILE, color=color)


def line(page, x1, y1, x2, y2, width=0.6, color=(0.35, 0.4, 0.5)):
    page.draw_line(pymupdf.Point(x1, y1), pymupdf.Point(x2, y2),
                   color=color, width=width)


# --------------------------------------------------------------------------
# 发票绘制
# --------------------------------------------------------------------------
def draw_invoice(page, d: dict, top: float = 0.0, scale: float = 1.0) -> float:
    """在页面上绘制一张发票，返回占用高度。d 为字段字典。"""
    is_elec = d.get("style") == "dz"        # 数电票
    is_old = d.get("style") == "old"        # 老版电子普票

    x0, x1 = 45, PAGE_W - 45
    y = top + 40

    title = {
        "dz": "电子发票（增值税专用发票）" if d.get("special") else "电子发票（普通发票）",
        "old": "增值税电子普通发票",
        "paper": "增值税专用发票" if d.get("special") else "增值税普通发票",
    }.get(d.get("style", "dz"), "电子发票（普通发票）")

    # 标题
    tw = pymupdf.get_text_length(title, fontname="china-s", fontsize=15) * scale
    txt(page, (PAGE_W - tw * 0.98) / 2, y, title, size=15 * scale)
    y += 8
    line(page, x0, y, x1, y, width=1.1, color=(0.12, 0.31, 0.47))
    y += 16

    # 号码 / 日期
    txt(page, x0, y, f"发票号码：{d['invoice_no']}", size=9.5 * scale,
        color=(0.75, 0.10, 0.10))
    txt(page, x1 - 165, y, f"开票日期：{d['issue_date']}", size=9.5 * scale)
    y += 14
    if d.get("invoice_code"):
        txt(page, x0, y, f"发票代码：{d['invoice_code']}", size=9.5 * scale)
        y += 14
    if d.get("check_code"):
        txt(page, x0, y, f"校验码：{d['check_code']}", size=9 * scale)
        y += 14
    line(page, x0, y, x1, y, width=0.5)
    y += 14

    # 购买方
    txt(page, x0, y, "购买方信息", size=9, color=(0.2, 0.3, 0.45))
    y += 13
    txt(page, x0 + 8, y, f"名称：{d['buyer_name']}", size=9 * scale)
    y += 13
    txt(page, x0 + 8, y, f"统一社会信用代码/纳税人识别号：{d['buyer_tax']}", size=9 * scale)
    y += 6
    line(page, x0, y, x1, y, width=0.5)
    y += 14

    # 销售方
    txt(page, x0, y, "销售方信息", size=9, color=(0.2, 0.3, 0.45))
    y += 13
    txt(page, x0 + 8, y, f"名称：{d['seller_name']}", size=9 * scale)
    y += 13
    txt(page, x0 + 8, y, f"统一社会信用代码/纳税人识别号：{d['seller_tax']}", size=9 * scale)
    y += 6
    line(page, x0, y, x1, y, width=0.5)
    y += 14

    # 明细表
    cols = [("货物或应税劳务、服务名称", x0 + 4), ("数量", 300), ("单价", 348),
            ("金额", 400), ("税率", 455), ("税额", 510)]
    for head, cx in cols:
        txt(page, cx, y, head, size=8.2 * scale, color=(0.2, 0.3, 0.45))
    y += 4
    line(page, x0, y, x1, y, width=0.5)
    y += 13

    txt(page, x0 + 4, y, f"*{d['item_cat']}*{d['item_name']}", size=8.6 * scale)
    txt(page, 300, y, "1", size=8.6 * scale)
    txt(page, 348, y, f"{d['amount']:.2f}", size=8.6 * scale)
    txt(page, 400, y, f"{d['amount']:.2f}", size=8.6 * scale)
    txt(page, 455, y, d["rate_label"], size=8.6 * scale)
    txt(page, 510, y, f"{d['tax']:.2f}", size=8.6 * scale)
    y += 8
    line(page, x0, y, x1, y, width=0.5)
    y += 13

    # 合计 / 价税合计
    txt(page, x0 + 4, y, "合    计", size=9 * scale)
    txt(page, 400, y, f"¥{d['amount']:,.2f}", size=9 * scale, color=(0.75, 0.10, 0.10))
    txt(page, 510, y, f"¥{d['tax']:,.2f}", size=9 * scale, color=(0.75, 0.10, 0.10))
    y += 8
    line(page, x0, y, x1, y, width=0.5)
    y += 15
    txt(page, x0 + 4, y, f"价税合计（大写）{d['total_cn']}   （小写）¥{d['total']:,.2f}",
        size=9.6 * scale, color=(0.75, 0.10, 0.10))
    y += 8
    line(page, x0, y, x1, y, width=0.5)
    y += 16
    txt(page, x0 + 4, y, f"开票人：{d.get('drawer', '张伟')}", size=9 * scale)
    y += 6
    line(page, x0, y, x1, y, width=0.5)
    return (y - top) + 30


def new_page(doc):
    return doc.new_page(width=PAGE_W, height=PAGE_H)


# --------------------------------------------------------------------------
# 样本数据
# --------------------------------------------------------------------------
# --------------------------------------------------------------------------
# 统一社会信用代码：按 GB 32100-2015 生成合法校验位
# （真实发票的税号是合法的，样本也必须合法，否则会误报「校验位不通过」）
# --------------------------------------------------------------------------
_USCC_CHARS = "0123456789ABCDEFGHJKLMNPQRTUWXY"
_USCC_W = [1, 3, 9, 27, 19, 26, 16, 17, 20, 29, 25, 13, 8, 24, 10, 30, 28]


def uscc(body17: str) -> str:
    total = sum(_USCC_CHARS.index(c) * w for c, w in zip(body17, _USCC_W))
    c = 31 - total % 31
    return body17 + _USCC_CHARS[0 if c == 31 else c]


def flip_last(code: str) -> str:
    """故意改坏最后一位校验位，用于演示「税号校验位不通过」告警。"""
    for ch in _USCC_CHARS:
        if ch != code[-1]:
            return code[:-1] + ch
    return code


SELLER = ("上海云启信息技术有限公司", uscc("91310115MA1H8WXYQ"))
BUYER = ("北京星辰科技有限公司", uscc("91110108MA01ABCDE"))
BAD_SELLER_TAX = flip_last(SELLER[1])

# --------------------------------------------------------------------------
# 发票代码 / 发票号码：按官方规则生成（规则见 invoice_wf/code_rules.py）
#   · 12 位代码 = 0 + 上海(3100) + 年度 + 批次 + 票种（11 普票 / 13 专票）
#   · 20 位号码 = 年份 + 省级(31) + 渠道 + 15 位顺序码
# 说明：真实场景中同一批次的发票共用同一个代码、靠号码区分；这里为了让每个
# 样本的代码都唯一、便于逐份核对，批次位按票递增。
# --------------------------------------------------------------------------
REGION = "上海"          # 销售方所在地，决定代码第 2-5 位与号码第 3-4 位
YEAR = 2026              # 与票面开票日期同年


def _no(seq: int) -> str:
    return e_invoice_no(REGION, YEAR, seq)


def _code(seq: int, kind: str = "普票") -> str:
    return invoice_code(REGION, YEAR, seq, kind)


BASE = dict(
    invoice_no=_no(1), invoice_code=_code(1), issue_date="2026年08月15日",
    buyer_name=BUYER[0], buyer_tax=BUYER[1], seller_name=SELLER[0], seller_tax=SELLER[1],
    amount=1000.00, tax=60.00, total=1060.00, total_cn="壹仟零陆拾圆整",
    rate_label="6%", item_cat="信息技术服务", item_name="信息系统技术服务费",
    drawer="张伟", style="dz", special=False, check_code="",
)


def sample(n: int, **kw) -> dict:
    d = dict(BASE)
    d.update(kw)
    return d


def build() -> list[tuple[str, str, list[dict]]]:
    """返回 [(文件名, 形态, [发票数据...])]，形态: pdf / scanpdf / png"""
    return [
        ("01_数电普票_技术服务费.pdf", "pdf", [
            sample(1)]),
        ("02_数电专票_咨询服务费.pdf", "pdf", [
            sample(2, invoice_no=_no(2), invoice_code=_code(2, "专票"), special=True,
                   item_cat="鉴证咨询服务", item_name="技术咨询服务费",
                   rate_label="6%", drawer="李娜")]),
        # 老版电子普票：12 位代码 + 8 位号码（号码按年度、分批次编制）
        ("03_电子普票_老版_办公用品.pdf", "pdf", [
            sample(3, style="old", invoice_no="04532198", invoice_code=_code(3),
                   check_code="08376 29511 88423 10665",
                   item_cat="办公用品", item_name="打印纸/硒鼓",
                   amount=800.00, tax=104.00, total=904.00, total_cn="玖佰零肆圆整",
                   rate_label="13%", drawer="王强")]),
        # 04 必须与 01 完全同号同码，才能命中「发票级重复」
        ("04_数电普票_与01同号_重复.pdf", "pdf", [
            sample(4, invoice_no=BASE["invoice_no"], invoice_code=BASE["invoice_code"],
                   issue_date="2026年08月15日")]),
        ("05_数电普票_勾稽不平.pdf", "pdf", [
            sample(5, invoice_no=_no(4), invoice_code=_code(4),
                   item_cat="信息技术服务", item_name="云服务器租赁费",
                   amount=2000.00, tax=120.00, total=2300.00,   # 2120 ≠ 2300
                   total_cn="贰仟叁佰圆整")]),
        ("06_数电专票_税号异常.pdf", "pdf", [
            sample(6, invoice_no=_no(5), invoice_code=_code(5, "专票"), special=True,
                   buyer_name=SELLER[0], buyer_tax=SELLER[1],      # 购销同名
                   seller_name=SELLER[0], seller_tax=BAD_SELLER_TAX,  # 校验位错误
                   amount=5000.00, tax=300.00, total=5300.00, total_cn="伍仟叁佰圆整")]),
        ("07_一页两张发票.pdf", "pdf", [
            sample(7, invoice_no=_no(6), invoice_code=_code(6), amount=300.00, tax=18.00,
                   total=318.00, total_cn="叁佰壹拾捌圆整",
                   item_cat="餐饮服务", item_name="会议餐费", drawer="赵敏"),
            sample(8, invoice_no=_no(7), invoice_code=_code(7), amount=1200.00, tax=72.00,
                   total=1272.00, total_cn="壹仟贰佰柒拾贰圆整",
                   item_cat="住宿服务", item_name="住宿费", drawer="赵敏")]),
        ("08_扫描件_无文本层.pdf", "scanpdf", [
            sample(9, invoice_no=_no(8), invoice_code=_code(8), amount=660.00, tax=39.60,
                   total=699.60, total_cn="陆佰玖拾玖圆陆角整",
                   item_cat="信息技术服务", item_name="软件开发服务")]),
        ("09_图片发票.png", "png", [
            sample(10, invoice_no=_no(9), invoice_code=_code(9), amount=450.00, tax=27.00,
                   total=477.00, total_cn="肆佰柒拾柒圆整",
                   item_cat="办公用品", item_name="办公耗材")]),
    ]


# --------------------------------------------------------------------------
# 落盘
# --------------------------------------------------------------------------
def render_pdf(items: list[dict], out: Path):
    doc = pymupdf.open()
    page = new_page(doc)
    y = 0.0
    for d in items:
        h = draw_invoice(page, d, top=y)
        if y + h > PAGE_H - 60:
            page = new_page(doc)
            y = 0.0
            h = draw_invoice(page, d, top=y)
        y += h + 10
    doc.save(str(out))
    doc.close()


def render_scan_pdf(items: list[dict], out: Path):
    """把发票渲染成位图再嵌入 → 生成没有文本层的「扫描件」PDF。"""
    src = pymupdf.open()
    page = new_page(src)
    for d in items:
        draw_invoice(page, d)
    pix = page.get_pixmap(matrix=pymupdf.Matrix(2.0, 2.0))
    src.close()

    img = Image.open(io.BytesIO(pix.tobytes("png")))
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=75)          # 用 JPEG 控制体积，更贴近真实扫描件

    doc = pymupdf.open()
    pg = doc.new_page(width=PAGE_W, height=PAGE_H)
    pg.insert_image(pymupdf.Rect(0, 0, PAGE_W, PAGE_H), stream=buf.getvalue())
    doc.save(str(out))
    doc.close()


def render_png(items: list[dict], out: Path):
    src = pymupdf.open()
    page = new_page(src)
    for d in items:
        draw_invoice(page, d)
    pix = page.get_pixmap(matrix=pymupdf.Matrix(2.2, 2.2))
    src.close()
    Image.open(io.BytesIO(pix.tobytes("png"))).save(str(out), "PNG")


def _to_png_bytes(img) -> bytes:
    buf = io.BytesIO()
    img.save(buf, "PNG")
    return buf.getvalue()


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    INBOX.mkdir(parents=True, exist_ok=True)

    made = []
    for name, kind, items in build():
        path = OUT / name
        if kind == "pdf":
            render_pdf(items, path)
        elif kind == "scanpdf":
            render_scan_pdf(items, path)
        elif kind == "png":
            render_png(items, path)
        made.append(path)
        print(f"  生成 {name}  ({path.stat().st_size / 1024:.1f} KB)")

    # 10) 与 01 字节完全相同的副本 → 触发「文件级重复」
    src = OUT / "01_数电普票_技术服务费.pdf"
    dup = OUT / "10_同一文件的副本.pdf"
    shutil.copyfile(src, dup)
    made.append(dup)
    print(f"  生成 {dup.name}  (01 的字节副本，用于演示文件级去重)")

    # 11) 07 的换名副本 —— 演示「同事把同一份发票又转发了一次」→ 文件级重复拦截。
    #     修改时间刻意设成与 07 完全相同：处理顺序的排序键是 (修改时间, 文件名)，
    #     两者同时刻时按名称排定，保证「07 原件先入账、11 转发件被拦截」每次一致。
    src7 = OUT / "07_一页两张发票.pdf"
    fwd = OUT / "11_同事又转发了一次.pdf"
    shutil.copyfile(src7, fwd)
    os.utime(fwd, (src7.stat().st_atime, src7.stat().st_mtime))
    made.append(fwd)
    print(f"  生成 {fwd.name}  (07 的换名副本，用于演示换名重复提交)")

    # 复制到输入目录，便于直接运行
    n = 0
    for p in made:
        target = INBOX / p.name
        shutil.copyfile(p, target)
        n += 1
    print(f"\n共生成 {n} 个样本，已复制到：{INBOX}")
    print("下一步：python main.py            （处理并生成台账）")
    print("        python main.py --dry-run  （只看识别结果，不写文件）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
