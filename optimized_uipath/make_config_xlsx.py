# -*- coding: utf-8 -*-
"""
生成 UiPath 的 Config.xlsx（REFramework 风格的配置表）
=====================================================
为什么不用 openpyxl：本机无法联网安装第三方包。所以这里用**纯标准库**
（zipfile + 手写最小 OOXML）生成 xlsx —— 生成出来的文件 Excel / WPS /
UiPath 的 Read Range 都能正常打开。

用法（在 IDLE 里直接 F5 也行）：
    python make_config_xlsx.py
产物：
    Config.xlsx        —— 两个 sheet：Settings / Constants
    Config.csv         —— 同样的内容，纯文本备份（xlsx 打不开时用）
"""
from __future__ import annotations

import csv
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

HERE = Path(__file__).resolve().parent

# --------------------------------------------------------------------------
# 配置内容（改这里即可）
# --------------------------------------------------------------------------
SETTINGS = [
    ("Name", "Value", "说明"),
    # 路径 —— 用绝对路径最稳；也可写成相对 Studio 工程根目录的相对路径
    ("InputFolder", r"D:\WorkBuddy-WorkSpace\invoice-automation\invoices_input",
     "发票收件箱：机器人扫描这个目录"),
    ("OutputFolder", r"D:\WorkBuddy-WorkSpace\invoice-automation\optimized_uipath\output",
     "输出目录：台账 / 报告 / 待复核清单都写在这里"),
    ("ProcessedFolder", r"D:\WorkBuddy-WorkSpace\invoice-automation\optimized_uipath\output\_processed",
     "处理完成的原件归档目录（MoveFile=True 时使用）"),
    ("FailedFolder", r"D:\WorkBuddy-WorkSpace\invoice-automation\optimized_uipath\output\_failed",
     "处理失败的原件隔离目录，供人工排查"),
    ("LedgerFile", "发票台账.xlsx", "台账文件名（放在 OutputFolder 下）"),
    ("ReportFile", "处理报告.html", "HTML 报告文件名"),
    ("ReviewFile", "待复核.csv", "低置信度/待人工录入清单，人工只看这一个文件"),
    ("AuditFile", "审计轨迹.jsonl", "每张票一行 JSON：提取器、质量分、逐字段置信度与证据"),

    # 行为开关
    ("MaxRetryNumber", "2", "单张发票的系统异常重试次数（REFramework 的 MaxRetryNumber）"),
    ("MoveFile", "False", "True = 处理完把原件移到 ProcessedFolder；False = 原地保留"),
    ("SendEmailOnFailure", "False", "True = 批次结束有问题时发邮件（需在 Orchestrator 配 Asset）"),
    ("AlertEmail", "", "告警收件人（SendEmailOnFailure=True 时必填）"),

    # 识别阈值
    ("MinTextChars", "20", "PDF 文本层少于这么多字符视为扫描件，转 OCR"),
    ("TextQualityMin", "0.55", "文本质量评分低于此值 → 升级提取方式（换 OCR / 调参数重试）"),
    ("OcrEnabled", "True", "是否允许调用 OCR（Document Understanding / Read PDF With OCR）"),
    ("OcrEngine", "UiPathDocumentOCR", "Document Understanding 的 OCR 引擎名"),

    # 置信度与校验阈值
    ("MinConfidence", "0.75", "记录级加权置信度阈值，低于此值进「待复核」"),
    ("MinCriticalConfidence", "0.5", "号码/日期/价税合计任一项低于此值 → 待复核"),
    ("BalanceTolerance", "0.02", "|金额+税额-价税合计| 容差"),
    ("RateRelTolerance", "0.02", "|税额-金额×税率| 相对容差"),
    ("RateAbsFloor", "0.05", "上面那条的绝对下限容差"),
    ("VendorSimilarity", "0.9", "购销方名称相似度告警阈值"),

    # 去重与人工回环
    ("OnDuplicate", "skip", "skip=不入账 | mark=入账但标注重复"),
    ("ReviewQueue", "InvoiceReview", "Orchestrator 队列名：低置信度发票进这里做人工复核"),
    ("UseOrchestratorQueue", "False",
     "True = 用 Orchestrator 队列做事务源（推荐生产）；False = 直接遍历文件夹（首次跑通更方便）"),
    ("QueueName", "InvoiceQueue", "UseOrchestratorQueue=True 时使用的事务队列名"),
]

CONSTANTS = [
    ("Name", "Value", "说明"),
    ("OK", "正常", "校验全部通过"),
    ("WARN", "存在提示", "有警告或提示，建议复核"),
    ("ERROR", "异常", "存在严重问题，需人工处理"),
    ("REVIEW", "待复核", "置信度不足，需人工确认"),
    ("TODO", "待人工录入", "图片/扫描件无法自动识别"),
    ("DUP", "重复未入账", "重复发票已阻止重复入账"),
    ("SheetDetail", "发票明细", "台账 sheet 名"),
    ("SheetIssue", "异常记录", "异常清单 sheet 名"),
    ("SheetSummary", "汇总统计", "统计 sheet 名"),
]


# --------------------------------------------------------------------------
# 最小 OOXML 写入器（纯标准库）
# --------------------------------------------------------------------------
def _col_letter(idx: int) -> str:
    """1 → A, 27 → AA"""
    s = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def _sheet_xml(rows) -> str:
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<worksheet xmlns="http://schemas.openxmlformats.org/'
             'spreadsheetml/2006/main"><sheetData>']
    for r, row in enumerate(rows, 1):
        parts.append(f'<row r="{r}">')
        for c, val in enumerate(row, 1):
            ref = f"{_col_letter(c)}{r}"
            txt = escape("" if val is None else str(val))
            # 全部按 inlineStr 写：省掉 sharedStrings 这一层，够用且不易出错
            parts.append(f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">'
                         f'{txt}</t></is></c>')
        parts.append("</row>")
    parts.append("</sheetData></worksheet>")
    return "".join(parts)


def write_xlsx(path: Path, sheets) -> None:
    """sheets: [(sheet_name, rows), ...]，rows 是 list[list[str]]。"""
    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
          '<Default Extension="xml" ContentType="application/xml"/>',
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>']
    for i in range(1, len(sheets) + 1):
        ct.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
                  f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    ct.append("</Types>")

    wb = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
          "<sheets>"]
    for i, (name, _rows) in enumerate(sheets, 1):
        wb.append(f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>')
    wb.append("</sheets></workbook>")

    rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for i in range(1, len(sheets) + 1):
        rels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>'
                    % (i, i))
    rels.append("</Relationships>")

    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
                 'officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                 "</Relationships>")

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", "".join(ct))
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", "".join(wb))
        z.writestr("xl/_rels/workbook.xml.rels", "".join(rels))
        for i, (_name, rows) in enumerate(sheets, 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", _sheet_xml(rows))


def main() -> int:
    xlsx = HERE / "Config.xlsx"
    write_xlsx(xlsx, [("Settings", SETTINGS), ("Constants", CONSTANTS)])
    print(f"已生成 {xlsx}")

    for name, rows in (("Config.csv", SETTINGS), ("Constants.csv", CONSTANTS)):
        p = HERE / name
        with open(p, "w", newline="", encoding="utf-8-sig") as fh:
            csv.writer(fh).writerows(rows)
        print(f"已生成 {p}")

    # 自检：能被 zipfile 重新读出来，且 XML 可解析
    import xml.etree.ElementTree as ET

    with zipfile.ZipFile(xlsx) as z:
        names = z.namelist()
        ok = 0
        for n in names:
            if n.endswith(".xml") or n.endswith(".rels"):
                ET.fromstring(z.read(n))
                ok += 1
    print(f"自检通过：{len(names)} 个部件，{ok} 个 XML 解析正常")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
