# -*- coding: utf-8 -*-
"""
统一配置中心
------------
- FIELDS      : Excel 台账的字段定义（key / 中文表头 / 列宽 / 数据类型）
- 阈值与开关  : 金额勾稽容差、OCR 触发阈值、重复处理策略等
调整字段顺序或表头，只需改这里，解析、校验、写表、报告会自动跟随。
"""

# --------------------------------------------------------------------------
# 1. Excel 台账字段定义（顺序 = 台账列顺序）
#    kind: text 文本 | money 金额(数值格式 #,##0.00) | int 整数 | date 文本日期
# --------------------------------------------------------------------------
FIELDS = [
    {"key": "seq",          "header": "序号",              "width": 6,  "kind": "int"},
    {"key": "file_name",    "header": "源文件名",          "width": 32, "kind": "text"},
    {"key": "invoice_type", "header": "发票类型",          "width": 22, "kind": "text"},
    {"key": "invoice_code", "header": "发票代码",          "width": 14, "kind": "text"},
    {"key": "invoice_no",   "header": "发票号码",          "width": 22, "kind": "text"},
    {"key": "issue_date",   "header": "开票日期",          "width": 13, "kind": "text"},
    {"key": "buyer_name",   "header": "购买方名称",        "width": 34, "kind": "text"},
    {"key": "buyer_tax",    "header": "购买方税号",        "width": 22, "kind": "text"},
    {"key": "seller_name",  "header": "销售方名称",        "width": 34, "kind": "text"},
    {"key": "seller_tax",   "header": "销售方税号",        "width": 22, "kind": "text"},
    {"key": "amount",       "header": "金额(不含税)",      "width": 14, "kind": "money"},
    {"key": "tax",          "header": "税额",              "width": 12, "kind": "money"},
    {"key": "total",        "header": "价税合计",          "width": 14, "kind": "money"},
    {"key": "tax_rate",     "header": "税率/征收率",       "width": 12, "kind": "text"},
    {"key": "item_name",    "header": "主要项目/货物名称", "width": 30, "kind": "text"},
    {"key": "check_code",   "header": "校验码",            "width": 24, "kind": "text"},
    {"key": "drawer",       "header": "开票人",            "width": 10, "kind": "text"},
    {"key": "status",       "header": "处理状态",          "width": 12, "kind": "text"},
    {"key": "issue",        "header": "异常/提示说明",     "width": 50, "kind": "text"},
    {"key": "processed_at", "header": "处理时间",          "width": 20, "kind": "text"},
    {"key": "file_hash",    "header": "文件指纹",          "width": 20, "kind": "text",
     "hidden": True},
]

HEADERS = [f["header"] for f in FIELDS]
KEYS = [f["key"] for f in FIELDS]

# --------------------------------------------------------------------------
# 2. 处理参数
# --------------------------------------------------------------------------
SETTINGS = {
    # 金额勾稽容差：|金额 + 税额 - 价税合计| 超过该值判为异常
    "balance_tolerance": 0.02,
    # PDF 文本层字符数低于该值，视为扫描件/图片，走 OCR
    "pdf_min_text_chars": 20,
    # OCR 子进程超时（秒）。首次运行需联网下载模型，可适当调大
    "ocr_timeout": 600,
    # 重复发票处理策略：skip=不重复入账（推荐）| mark=入账但标记为重复
    "on_duplicate": "skip",
    # 同一文件内容指纹重复时是否强制重新解析
    "reparse_same_file": False,
    # 是否把「提示」级别问题也写进异常记录页
    "report_hint_level": True,
    # 台账 sheet 名
    "sheet_detail": "发票明细",
    "sheet_issue": "异常记录",
    "sheet_summary": "汇总统计",
}

# --------------------------------------------------------------------------
# 3. 支持的文件类型
# --------------------------------------------------------------------------
PDF_EXTS = {".pdf"}
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}
SUPPORTED_EXTS = PDF_EXTS | IMAGE_EXTS

# 处理状态枚举
STATUS_OK = "正常"
STATUS_WARN = "存在提示"
STATUS_ERROR = "异常"
STATUS_OCR_TODO = "待人工录入"
STATUS_DUPLICATE = "重复未入账"

# 合法税率集合（用于提示非法税率）
VALID_TAX_RATES = {"0%", "1%", "1.5%", "3%", "5%", "6%", "9%", "10%", "11%",
                   "13%", "16%", "17%", "免税", "不征税", "***"}

# 默认输入/输出目录（相对于项目根）
DEFAULT_INPUT_DIR = "invoices_input"
DEFAULT_LEDGER = "output/发票台账.xlsx"
DEFAULT_REPORT = "output/处理报告.html"
