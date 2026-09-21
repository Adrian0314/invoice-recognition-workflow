# -*- coding: utf-8 -*-
"""
发票识别工作流机器人 · 纯 Python 单文件版（v2）
================================================
设计目标：**在 Python IDLE 里按 F5 就能跑**，不需要 pip install 也能工作。

    ┌─────────────────────────────────────────────────────────────────┐
    │  在 IDLE 中：File → Open → 本文件 → Run Module (F5)             │
    │  或命令行：python idle_invoice_bot.py --input 发票目录          │
    └─────────────────────────────────────────────────────────────────┘

相对 v1（invoice_wf 包）的改进，逐项对应 GitHub 上主流方案的成熟做法：

  A. 模板化解析（借鉴 invoice2data 的 YAML 模板体系）
     版式规则搬到 invoice_templates.json：keywords 选模板、每字段可写多条
     正则、可加静态字段。新增一种供应商版式 = 加一条 JSON，不动代码。

  B. 级联文本提取 + 文本质量闸门（借鉴 invoice2data 的 cascading backend）
     PyMuPDF → 纯标准库 PDF 解析器 → 预处理 + 多参数 OCR 重试。
     每档产物先过"质量评分"，不合格才升档，而不是只看字符数。

  C. 逐字段置信度 + 待复核队列（借鉴 AI-Invoice-Processor / Smart Invoice
     Parser 的 confidence gating，对应 UiPath 的 Validation Station）
     字段级 conf 0~1 → 记录级加权置信度 → 低于阈值进「待复核」，不再只有
     "正常/异常" 两态，人工只需看真正可疑的少数几行。

  D. 确定性解析 + LLM 兜底（默认关闭，离线也能跑）
     仅当置信度低且用户显式开启时才调 LLM；LLM 结果只做"合并"，且必须通过
     同一套校验，不会覆盖高置信度的确定性结果。

  E. 校验规则注册表（可开关、可扩展）
     每条规则带 id/级别/说明，可在 config.json 里单独禁用；新增规则只需加
     一个函数，不动机器。

  F. 新增 4 类财务级交叉校验（提升准确性的关键）
     税率×金额 与 税额 勾稽、数电号码结构（年份前缀+省级码）、发票代码结构、
     开票日期与号码年份一致性。这几条能抓住"单个字段看起来合法、组合起来
     不自洽"的识别错误。

  G. 审计轨迹 JSONL（借鉴 GitHub Actions 方案的 artifact/审计思路）
     每张票一行 JSON：提取器、质量分、每字段置信度与命中证据、耗时。
     出问题时可回溯"当时到底看到了什么"，而不是只有最终结果。

  H. 零依赖降级
     没装 openpyxl → 只出 CSV/JSON（功能不减，仅少一个 xlsx）；
     没装 PyMuPDF → 用内置标准库 PDF 解析器；
     没装 OCR → 图片登记为「待人工录入」。

  I. 幂等与安全
     文件指纹 + 发票唯一键双去重；重跑不膨胀；台账被 Excel 占用时给中文提示
     而不是抛堆栈。

用法速查：
    python idle_invoice_bot.py                     # 处理默认目录
    python idle_invoice_bot.py -i D:\\发票 -o D:\\台账
    python idle_invoice_bot.py --dry-run           # 只看识别结果，不写文件
    python idle_invoice_bot.py --selftest          # 内置自检（不读任何文件）
    python idle_invoice_bot.py --explain           # 打印模板与配置的当前值
"""

from __future__ import annotations

import argparse
import hashlib
import html as _html
import json
import os
import re
import sys
import time
import unicodedata
import zlib
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

__version__ = "2.1.0"

APP_DIR = Path(__file__).resolve().parent


# ==========================================================================
# 0. 配置区 —— 在 IDLE 里直接改这里最省事，也可写 config.json / 环境变量覆盖
# ==========================================================================
DEFAULT_CONFIG = {
    # ---- 目录 ----
    "input_dir": str(APP_DIR / "invoices_input"),
    "output_dir": str(APP_DIR.parent / "output"),  # 统一写入项目根的 output/，不再区分 v1/v2
    "recursive": True,
    "move_processed": False,                    # True = 处理完把原件移到 processed_dir
    "processed_dir": "_processed",
    "failed_dir": "_failed",

    # ---- 产出文件名 ----
    "ledger_name": "发票台账.xlsx",
    "report_name": "处理报告.html",
    "csv_name": "发票明细.csv",
    "json_name": "发票明细.json",
    "issue_csv_name": "异常记录.csv",
    "review_csv_name": "待复核.csv",
    "audit_name": "审计轨迹.jsonl",
    # 内部状态文件：零依赖路径下用它跨批次恢复台账（勿手工编辑）
    "state_name": "台账索引.json",

    # ---- 提取档位 ----
    "min_text_chars": 20,        # 文本层少于这么多字符 → 认为没有文本层
    "text_quality_min": 0.55,    # 文本质量评分低于此值 → 升级提取方式
    "ocr_enabled": "auto",       # auto | off
    "ocr_timeout": 600,          # OCR 子进程超时（秒），实际生效值
    "ocr_zoom": 2.5,             # 扫描件栅格化倍率

    # ---- 标准库 PDF 解析器的解码上限（性能护栏，见 pdf_text_stdlib 说明）----
    "cmap_probe_max_kb": 512,    # 第一遍找 ToUnicode CMap 时，超过此大小的流不解码
    "pdf_stream_max_mb": 16,     # 内容流解码上限；超过则跳过（避免被图像流拖垮）

    # ---- 置信度 ----
    "min_confidence": 0.75,          # 记录级加权置信度阈值
    "min_critical_confidence": 0.50, # 任一关键字段（号码/日期/价税合计）低于此值 → 待复核

    # ---- 校验 ----
    "balance_tolerance": 0.02,       # |金额+税额-价税合计| 容差
    "rate_rel_tolerance": 0.02,      # |税额-金额×税率| 相对容差
    "rate_abs_floor": 0.05,          # 上面那条的绝对下限容差
    "vendor_similarity": 0.90,       # 供应商名称相似度告警阈值
    "future_date_grace_days": 0,     # 允许"开票日期"最多超前今天几天（0=不允许）

    # ---- 去重 ----
    "on_duplicate": "skip",          # skip=不入账 | mark=入账但标注重复

    # ---- LLM 兜底（默认关闭；开启才会联网）----
    "llm_fallback": {
        "enabled": False,
        "endpoint": "https://api.openai.com/v1/chat/completions",
        "api_key_env": "INVOICE_LLM_API_KEY",
        "model": "gpt-4o-mini",
        "timeout": 30,
        "max_text_chars": 6000,
    },

    # ---- 规则开关（key = 规则 id，值 False 表示禁用该规则）----
    "rules_disabled": [],
}

# 合法的税率写法（用于"税率是否常见"这一条提示类规则）
VALID_TAX_RATES = {"0%", "1%", "1.5%", "3%", "5%", "6%", "9%", "10%", "11%",
                   "13%", "16%", "17%", "免税", "不征税", "***"}

# 台账列定义（顺序 = 落表顺序）
FIELDS = [
    ("seq",          "序号",              "int"),
    ("file_name",    "源文件名",          "text"),
    ("invoice_type", "发票类型",          "text"),
    ("invoice_code", "发票代码",          "text"),
    ("invoice_no",   "发票号码",          "text"),
    ("issue_date",   "开票日期",          "text"),
    ("buyer_name",   "购买方名称",        "text"),
    ("buyer_tax",    "购买方税号",        "text"),
    ("seller_name",  "销售方名称",        "text"),
    ("seller_tax",   "销售方税号",        "text"),
    ("amount",       "金额(不含税)",      "money"),
    ("tax",          "税额",              "money"),
    ("total",        "价税合计",          "money"),
    ("tax_rate",     "税率/征收率",       "text"),
    ("item_name",    "主要项目/货物名称", "text"),
    ("check_code",   "校验码",            "text"),
    ("drawer",       "开票人",            "text"),
    ("confidence",   "识别置信度",        "text"),
    ("extractor",    "提取方式",          "text"),
    ("status",       "处理状态",          "text"),
    ("issue",        "异常/提示说明",     "text"),
    ("processed_at", "处理时间",          "text"),
    ("file_hash",    "文件指纹",          "text"),
]
HEADERS = [h for _, h, _ in FIELDS]
KEYS = [k for k, _, _ in FIELDS]
MONEY_KEYS = [k for k, _, t in FIELDS if t == "money"]

# 处理状态
ST_OK = "正常"
ST_WARN = "存在提示"
ST_ERROR = "异常"
ST_REVIEW = "待复核"
ST_TODO = "待人工录入"
ST_DUP = "重复未入账"

PDF_EXTS = {".pdf"}
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}
SUPPORTED_EXTS = PDF_EXTS | IMAGE_EXTS


def load_config() -> dict:
    """默认值 ← config.json ← 环境变量 INVOICE_XXX（后者优先级最高）。"""
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))  # deep copy

    side = APP_DIR / "config.json"
    if side.exists():
        try:
            user = json.loads(side.read_text(encoding="utf-8"))
            for k, v in (user or {}).items():
                if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                    cfg[k].update(v)
                else:
                    cfg[k] = v
        except Exception as exc:                     # noqa: BLE001
            print(f"[配置] config.json 读取失败，已忽略：{exc}")

    for key in list(cfg):
        if isinstance(cfg[key], dict):
            continue
        env = os.environ.get("INVOICE_" + key.upper())
        if env is None:
            continue
        cur = cfg[key]
        try:
            if isinstance(cur, bool):
                cfg[key] = env.strip().lower() in ("1", "true", "yes", "on")
            elif isinstance(cur, int):
                cfg[key] = int(env)
            elif isinstance(cur, float):
                cfg[key] = float(env)
            else:
                cfg[key] = env
        except ValueError:
            pass

    # 相对路径统一按脚本所在目录解析，方便"换个地方双击也能跑"
    for key in ("input_dir", "output_dir"):
        p = Path(cfg[key])
        if not p.is_absolute():
            cfg[key] = str(APP_DIR / p)
    return cfg


# ==========================================================================
# 1. 版式模板（可被 invoice_templates.json 整体覆盖）
#    —— 借鉴 invoice2data：keywords 选模板 / 每字段多条正则 / 静态字段
# ==========================================================================
DEFAULT_TEMPLATES = [
    {
        "name": "通用-中国大陆增值税/数电发票",
        "keywords": ["发票"],
        "exclude_keywords": [],
        "type": "",                       # 留空 = 交给 detect_type() 判断
        "fields": {},                     # 留空 = 全部走内置启发式（见第 4 节）
        "static": {},
    },
]

# 示例：一个"给某供应商加专用正则"的模板长这样（写进 invoice_templates.json 即可生效）
TEMPLATE_EXAMPLE = {
    "name": "示例-某供应商专用版式",
    "keywords": ["某某科技有限公司"],
    "exclude_keywords": [],
    "type": "电子普通发票",
    "fields": {
        "invoice_no": [r"发票号码[:：]?\s*([0-9]{8,20})"],
        "total": [r"价税合计.{0,20}?[¥￥]?\s*(-?[0-9][0-9,]*\.[0-9]{2})"],
    },
    "static": {"invoice_type": "电子普通发票"},
}


def load_templates() -> list:
    side = APP_DIR / "invoice_templates.json"
    if not side.exists():
        return list(DEFAULT_TEMPLATES)
    try:
        data = json.loads(side.read_text(encoding="utf-8"))
        if isinstance(data, list) and data:
            print(f"[模板] 已加载 {len(data)} 个外部版式模板：{side.name}")
            return data
        print("[模板] invoice_templates.json 内容不是非空数组，改用内置模板")
    except Exception as exc:                          # noqa: BLE001
        print(f"[模板] invoice_templates.json 解析失败，改用内置模板：{exc}")
    return list(DEFAULT_TEMPLATES)


# ==========================================================================
# 2. 文本工具
# ==========================================================================
def setup_console() -> None:
    """中文 Windows 控制台是 GBK，重定向到文件时 ✓ ✗ 会抛 UnicodeEncodeError。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream is not None and hasattr(stream, "reconfigure"):
                stream.reconfigure(errors="replace")
        except Exception:
            pass


def norm_text(raw: str) -> str:
    t = (raw or "").replace("\r\n", "\n").replace("\r", "\n")
    for ch in ("\u00a0", "\u3000", "\u200b", "\ufeff"):
        t = t.replace(ch, " ")
    return unicodedata.normalize("NFKC", t)


def to_oneline(t: str) -> str:
    return re.sub(r"\s+", "", t or "")


def to_lines(t: str) -> list:
    out = []
    for ln in (t or "").split("\n"):
        ln = re.sub(r"[ \t]+", "", ln).strip()
        if ln:
            out.append(ln)
    return out


MONEY_STR = r"-?\d[\d,]*(?:\.\d{1,2})?"


def to_money(s):
    if s in (None, ""):
        return None
    s = str(s).replace(",", "").replace("¥", "").replace("￥", "").strip()
    try:
        return round(float(s), 2)
    except ValueError:
        return None


# ==========================================================================
# 2.1 中文大写金额 → 数字
#     用途：与「价税合计（小写）」交叉勾稽。这是财务票据最有价值的一条
#     互相印证关系 —— 大小写不一致几乎必然是识别错误（OCR 认错字 / 字体映射错），
#     而单字段校验完全抓不到。规则见 R17。
# ==========================================================================
_CN_DIGITS = {
    "零": 0, "壹": 1, "贰": 2, "叁": 3, "肆": 4, "伍": 5, "陆": 6, "柒": 7,
    "捌": 8, "玖": 9, "〇": 0, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
    "六": 6, "七": 7, "八": 8, "九": 9, "两": 2,
}
_CN_UNITS = {"拾": 10, "佰": 100, "仟": 1000, "十": 10, "百": 100, "千": 1000}
_CN_SECTIONS = {"万": 10 ** 4, "亿": 10 ** 8}


def cn_amount_to_number(s):
    """
    把「壹仟零陆拾圆整」这类中文大写金额转成数字；无法解析返回 None。

    宽容度：圆/圓/圜/元 都认、整/正 忽略、可带空格与全角符号、
    也接受 一二三 这类小写汉字。分/角单独处理。
    """
    if not s:
        return None
    t = str(s)
    t = re.sub(r"[（(][^)）]*[)）]", "", t)          # 去掉「(大写)」这类括注
    t = t.replace("整", "").replace("正", "")
    t = t.replace("圆", "元").replace("圓", "元").replace("圜", "元")
    t = re.sub(r"[￥¥,，\s]", "", t)
    if not t or not any(ch in t for ch in _CN_DIGITS) \
            and "元" not in t and "角" not in t and "分" not in t:
        return None

    jiao = fen = 0
    m = re.search(r"([零壹贰叁肆伍陆柒捌玖〇一二三四五六七八九])角", t)
    if m:
        jiao = _CN_DIGITS[m.group(1)]
    m = re.search(r"([零壹贰叁肆伍陆柒捌玖〇一二三四五六七八九])分", t)
    if m:
        fen = _CN_DIGITS[m.group(1)]

    main = re.split(r"[元角分]", t)[0]
    total = section = num = 0
    for ch in main:
        if ch in _CN_DIGITS:
            num = _CN_DIGITS[ch]
        elif ch in _CN_UNITS:
            section += (num or 1) * _CN_UNITS[ch]
            num = 0
        elif ch in _CN_SECTIONS:
            section = (section + num) * _CN_SECTIONS[ch]
            total += section
            section = num = 0
    value = total + section + num + jiao / 10.0 + fen / 100.0
    return round(value, 2)


def money_matches(s) -> bool:
    if s in (None, ""):
        return True
    try:
        return abs(float(s)) < 1e-9
    except (TypeError, ValueError):
        return False


# ==========================================================================
# 3. 文本提取：级联 + 质量闸门
#    —— 借鉴 invoice2data 的 cascading backend 与 OCR 前处理实践
# ==========================================================================
def _have(mod: str) -> bool:
    import importlib.util

    try:
        return importlib.util.find_spec(mod) is not None
    except (ImportError, ValueError):
        return False


# ------------------------- 3.1 纯标准库 PDF 文本提取 -------------------------
_ESC = {"n": "\n", "r": "\r", "t": "\t", "b": "\b", "f": "\f",
        "(": "(", ")": ")", "\\": "\\"}


def _pdf_unescape(raw: bytes) -> str:
    out = bytearray()
    i, n = 0, len(raw)
    while i < n:
        c = raw[i]
        if c == 0x5C and i + 1 < n:               # 反斜杠
            nxt = raw[i + 1:i + 2]
            if nxt in (b"n", b"r", b"t", b"b", b"f", b"(", b")", b"\\"):
                out += _ESC[nxt.decode("latin-1")].encode("latin-1")
                i += 2
                continue
            m = re.match(rb"[0-7]{1,3}", raw[i + 1:i + 4])
            if m:
                out.append(int(m.group(0), 8) & 0xFF)
                i += 1 + len(m.group(0))
                continue
            i += 2
            continue
        out.append(c)
        i += 1
    return out.decode("latin-1", "replace")


def _parse_tounicode(cmap_bytes: bytes) -> dict:
    """解析 ToUnicode CMap：把 CID 映射回真正的 Unicode 字符。"""
    text = cmap_bytes.decode("latin-1", "replace")
    mapping = {}

    def to_uni(hexs: str) -> str:
        hexs = re.sub(r"[^0-9A-Fa-f]", "", hexs)
        try:
            b = bytes.fromhex(hexs)
        except ValueError:
            return ""
        for enc in ("utf-16-be", "latin-1"):
            try:
                return b.decode(enc)
            except UnicodeDecodeError:
                continue
        return ""

    for blk in re.findall(r"beginbfchar(.*?)endbfchar", text, re.S):
        for src, dst in re.findall(r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", blk):
            mapping[int(src, 16)] = to_uni(dst)
    for blk in re.findall(r"beginbfrange(.*?)endbfrange", text, re.S):
        for lo, hi, dst in re.findall(
                r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", blk):
            start, end = int(lo, 16), int(hi, 16)
            base = to_uni(dst)
            if not base:
                continue
            if len(base) == 1 and end - start < 0x10000:
                for k in range(end - start + 1):
                    mapping[start + k] = chr(ord(base[0]) + k)
            else:
                mapping[start] = base
    return mapping


def _stream_filters(head: bytes) -> list:
    """解析 /Filter 字典项，返回滤镜名列表（按应用顺序）。"""
    m = re.search(rb"/Filter\s*\[([^\]]*)\]", head)
    if m:
        return re.findall(rb"/([A-Za-z0-9]+)", m.group(1))
    m = re.search(rb"/Filter\s*/([A-Za-z0-9]+)", head)
    return [m.group(1)] if m else []


# 只可能是图像的滤镜 / 子类型：这类流里不可能有 ToUnicode CMap 或文本绘制指令，
# 直接跳过解码。真实扫描件 PDF 里图像流往往占 95% 以上体积。
_IMAGE_FILTERS = {b"DCTDecode", b"JPXDecode", b"JBIG2Decode", b"CCITTFaxDecode"}
_IMAGE_SUBTYPE = re.compile(rb"/Subtype\s*/Image")


def is_image_stream(head: bytes) -> bool:
    if _IMAGE_SUBTYPE.search(head):
        return True
    return any(f in _IMAGE_FILTERS for f in _stream_filters(head))


def stream_reader(objs: dict, cmap_limit: int, content_limit: int):
    """
    带缓存与体量护栏的流解码器（性能护栏，v2.1 新增）。

    为什么需要它（实测数据）：
      · 旧实现把每个对象解码两遍（第一遍找 CMap、第三遍取内容流），
        44 个对象的样本实测调用 88 次；
      · 更要命的是与票面无关的大流（扫描件的图像流）也会被完整解压 ——
        合成测试里 8 个 2MB 的 Flate 流（共 16MB）被解出 256MB 数据、耗时 628ms，
        而文件里真正有用的文本流只有几十字节。

    做法：图像流永久跳过；其余流按「本次请求的上限」决定是否解码，并缓存结果。
    上限放宽后的重试仍然允许（第一遍 512KB 没解，第三遍放宽到 16MB 会重试）。
    """
    dec, skip = {}, {}

    def get(num: int, limit: int):
        if num in skip:
            return None
        prev = dec.get(num)
        if prev is not None and (prev[1] is not None or prev[0] >= limit):
            return prev[1]           # 已成功解出，或已用不小于本次的上限试过
        body = objs.get(num, b"")
        head = body.split(b"stream", 1)[0] if b"stream" in body else body
        if is_image_stream(head):
            skip[num] = True         # 图像流永远不需要解码
            return None
        if len(body) > limit:        # body 含流字典，作为体积的保守上界
            dec[num] = (limit, None)
            return None
        out = _decode_stream(body)
        dec[num] = (limit, out)
        return out

    return get


def _apply_one_filter(data: bytes, name: bytes):
    if name == b"FlateDecode":
        for fn in (lambda d: zlib.decompress(d),
                   lambda d: zlib.decompressobj().decompress(d),
                   lambda d: zlib.decompressobj(-15).decompress(d)):
            try:
                return fn(data)
            except Exception:                         # noqa: BLE001
                continue
        return None
    if name == b"ASCII85Decode":
        import base64

        s = data.strip()
        if s.startswith(b"<~"):
            s = s[2:]
        if s.endswith(b"~>"):
            s = s[:-2]
        for kwargs in ({"adobe": False}, {"adobe": True}):
            try:
                return base64.a85decode(
                    s, ignorechars=b" \t\r\n\v\f", **kwargs)
            except Exception:                         # noqa: BLE001
                continue
        return None
    if name == b"ASCIIHexDecode":
        try:
            return bytes.fromhex(re.sub(rb"[^0-9A-Fa-f]", b"", data).decode("latin-1"))
        except ValueError:
            return None
    return None                                       # 其他滤镜（LZW/DCT…）不支持


def _decode_stream(body: bytes):
    """
    取出并解码一个 PDF 流对象。

    支持滤镜链 —— reportlab 生成的页面内容常写成
        /Filter [ /ASCII85Decode /FlateDecode ]
    只处理 FlateDecode 是不够的（这正是 v2 初版读不出 reportlab 样票的原因）。
    """
    m = re.search(rb"stream(\r\n|\r|\n)", body)
    if not m:
        return None
    start = m.end()
    end = body.find(b"endstream", start)
    raw = body[start:end] if end > 0 else body[start:]
    head = body[:m.start()]

    filters = _stream_filters(head)
    if not filters:
        return raw
    data = raw
    for f in filters:
        data = _apply_one_filter(data, f)
        if data is None:
            return None
    return data


def _pdf_objs(data: bytes) -> dict:
    objs = {}
    for m in re.finditer(rb"(\d+)\s+\d+\s+obj\b(.*?)\bendobj", data, re.S):
        objs[int(m.group(1))] = m.group(2)
    return objs


# 预定义 CMap 名称 → 解码方式
#   UTF-16 系列：字符码本身就是 UTF-16BE（PyMuPDF 内置 Heiti/china-s 等走这条，
#                这类字体通常没有 ToUnicode，只能靠编码名判断）
#   GBK 系列：字符码是 GBK 双字节
_ENC_UTF16 = {
    b"UniGB-UCS2-H", b"UniGB-UTF16-H", b"UniGB-UCS2-V", b"UniGB-UTF16-V",
    b"UniCNS-UCS2-H", b"UniCNS-UTF16-H", b"UniCNS-UTF16-V",
    b"UniJIS-UCS2-H", b"UniJIS-UTF16-H", b"UniJIS-UCS2-V",
    b"UniKS-UCS2-H", b"UniKS-UTF16-H", b"UniKS-UCS2-V",
}
_ENC_GBK = {b"GBK-EUC-H", b"GBK-EUC-V", b"GBpc-EUC-H", b"B5pc-H", b"ETen-B5-H"}


def pdf_text_stdlib(path, cfg: dict | None = None) -> tuple:
    """
    纯标准库的 PDF 文本提取（best-effort，零依赖）。

    覆盖到的三种字体映射方式（实测覆盖本项目全部样本）：
      1. 字体自带 /ToUnicode CMap（PyMuPDF / Word 导出常见）→ 按 CMap 映射
      2. /Encoding 是预定义 UTF-16 CMap（UniGB-UTF16-H 等，PyMuPDF 内置 CJK 字体走这条）
         → 直接把双字节码当 UTF-16BE 解
      3. /Encoding 是 GBK-EUC-H 等 → 按 GBK 解
    另外支持滤镜链（/ASCII85Decode /FlateDecode，reportlab 常见）。

    不覆盖：加密 PDF、非 ASCII85/Flate 的滤镜、纯图片扫描件（本就要 OCR）。

    性能：流解码走 `stream_reader`（缓存 + 图像流跳过 + 体量上限），
    所以扫描件里的大图像流不会再被白白解压。
    """
    cfg = cfg or {}
    data = Path(path).read_bytes()
    objs = _pdf_objs(data)
    if not objs:
        return "", 0

    cmap_limit = max(64, int(cfg.get("cmap_probe_max_kb", 512))) * 1024
    content_limit = max(1024, int(cfg.get("pdf_stream_max_mb", 16))) * 1024 * 1024
    get_stream = stream_reader(objs, cmap_limit, content_limit)

    # 1) 收集所有 ToUnicode CMap（obj 号 → CID→字符）
    cmaps = {}
    for num in objs:
        raw = get_stream(num, cmap_limit)
        if raw and (b"beginbfchar" in raw or b"beginbfrange" in raw):
            cmaps[num] = _parse_tounicode(raw)

    # 2) 字体对象 → 解码方式 {obj 号: {"cmap":…, "enc":…, "two_byte":bool}}
    fonts = {}
    for num, body in objs.items():
        if b"/Font" not in body and b"/BaseFont" not in body:
            continue
        info = {"cmap": None, "enc": None, "two_byte": False}
        m = re.search(rb"/ToUnicode\s+(\d+)\s+\d+\s+R", body)
        if m:
            info["cmap"] = cmaps.get(int(m.group(1)))
        sub = re.search(rb"/Subtype\s*/([A-Za-z0-9]+)", body)
        if sub:
            # Type0 = 复合字体，字符码是 2 字节 CID；
            # TrueType / Type1 = 简单字体，字符码是 1 字节（reportlab 的
            # 子集字体就属于这类，CID 即字形序号，必须按单字节查 CMap）
            info["two_byte"] = sub.group(1) == b"Type0"
        enc = re.search(rb"/Encoding\s*/([A-Za-z0-9\-]+)", body)
        if enc:
            name = enc.group(1)
            if name in _ENC_UTF16:
                info["enc"] = "utf16"
            elif name in _ENC_GBK:
                info["enc"] = "gbk"
        if info["cmap"] or info["enc"]:
            fonts[num] = info

    # 3) 资源字典里 /Fx -> 字体对象号
    #    两种写法都要认：
    #      · 内联   /Resources << /Font << /F1 2 0 R >> >>
    #      · 间接   /Resources << /Font 1 0 R >>  →  obj 1 才是 << /F2+0 16 0 R >>
    #    reportlab 用的是后者。字体资源名可能带 '+'（如 /F2+0），名字里必须允许它。
    _NAME = rb"/([A-Za-z0-9#_.\-+]+)\s+(\d+)\s+\d+\s+R"
    name_to_font = {}

    def register(blk: bytes):
        for nm, ref in re.findall(_NAME, blk):
            f = fonts.get(int(ref))
            if f:
                name_to_font[nm.decode("latin-1")] = f

    pending = set()
    for _num, body in objs.items():
        for blk in re.findall(rb"/Font\s*<<(.*?)>>", body, re.S):
            register(blk)
        m = re.search(rb"/Font\s+(\d+)\s+\d+\s+R", body)
        if m:
            pending.add(int(m.group(1)))
    for num in pending:
        body = objs.get(num, b"")
        m = re.search(rb"<<(.*)>>", body, re.S)
        register(m.group(1) if m else body)

    # 4) 页数（粗计：/Type /Page 出现次数）
    pages = len(re.findall(rb"/Type\s*/Page[^s]", data))

    # 5) 遍历内容流，按操作符取文本
    chunks = []
    for num in objs:
        cs = get_stream(num, content_limit)
        if not cs or (b"Tj" not in cs and b"TJ" not in cs):
            continue
        cur = None
        for m in re.finditer(
                rb"/([A-Za-z0-9#_.+\-]+)\s+[\d.\-]+\s+Tf"         # 选字体
                rb"|\[((?:[^\[\]\\]|\\.)*)\]\s*TJ"                # 数组串
                rb"|\(((?:[^()\\]|\\.)*)\)\s*Tj"                  # 简单串
                rb"|<([0-9A-Fa-f\s]*)>\s*Tj"                      # 十六进制串
                rb"|(T\*|Td|TD|Tm|ET)", cs, re.S):
            if m.group(1):
                cur = name_to_font.get(m.group(1).decode("latin-1"))
                continue
            if m.group(5):
                chunks.append("\n")
                continue

            # 取出生字节：字面串按 PDF 转义规则解，十六进制串按 hex 解
            blobs = []
            if m.group(2) is not None:
                for sub in re.finditer(rb"\(((?:[^()\\]|\\.)*)\)|<([0-9A-Fa-f\s]*)>",
                                       m.group(2), re.S):
                    if sub.group(1) is not None:
                        blobs.append(_pdf_unescape(sub.group(1))
                                     .encode("latin-1", "replace"))
                    else:
                        hx = re.sub(rb"\s", b"", sub.group(2))
                        if len(hx) % 2:
                            hx += b"0"
                        try:
                            blobs.append(bytes.fromhex(hx.decode("latin-1")))
                        except ValueError:
                            pass
            elif m.group(3) is not None:
                blobs.append(_pdf_unescape(m.group(3)).encode("latin-1", "replace"))
            else:
                hx = re.sub(rb"\s", b"", m.group(4) or b"")
                if len(hx) % 2:
                    hx += b"0"
                try:
                    blobs.append(bytes.fromhex(hx.decode("latin-1")))
                except ValueError:
                    pass

            for blob in blobs:
                if not blob:
                    continue
                chunks.append(_decode_pdf_bytes(blob, cur))
    return "".join(chunks), pages


def _cmap_decode(blob: bytes, cmap: dict, width: int) -> str:
    if width == 2:
        return "".join(cmap.get(blob[i] * 256 + blob[i + 1], "")
                       for i in range(0, len(blob) - 1, 2))
    return "".join(cmap.get(b, "") for b in blob)


def _readable_ratio(s: str) -> float:
    """可读字符占比：用于在"单字节/双字节"两种解之间择优。"""
    if not s:
        return -1.0
    good = sum(1 for c in s
               if c.isalnum() or "\u4e00" <= c <= "\u9fff"
               or c in "¥￥.,:-/()（）年月日 　")
    return good / len(s)


def _decode_pdf_bytes(blob: bytes, font) -> str:
    """
    按字体的映射方式把 PDF 原始字节还原成可读文本。

    单字节 / 双字节两种走法都试一遍，按"可读字符占比"择优 —— 这一步很关键：
    同一个 ToUnicode 表，用错字宽会得到"部分字符恰好命中"的乱码（非空！），
    只判断"结果非空"就会把乱码当成成功。
    """
    if font:
        cmap = font.get("cmap")
        if cmap:
            declared = 2 if font.get("two_byte") else 1
            best, best_score = "", -2.0
            for width in (declared, 3 - declared):
                s = _cmap_decode(blob, cmap, width)
                score = _readable_ratio(s) + (0.15 if width == declared else 0.0)
                if score > best_score:
                    best, best_score = s, score
            if best:
                return best
        enc = font.get("enc")
        if enc == "utf16":
            return blob.decode("utf-16-be", "replace")
        if enc == "gbk":
            try:
                return blob.decode("gbk", "replace")
            except Exception:                             # noqa: BLE001
                return ""
    return blob.decode("latin-1", "replace")


# ------------------------- 3.2 PyMuPDF 提取 -------------------------
def pdf_text_pymupdf(path) -> tuple:
    try:
        import pymupdf as fitz                      # type: ignore
    except ImportError:
        import fitz                                 # type: ignore
    doc = fitz.open(str(path))
    try:
        return "\n".join(page.get_text("text") or "" for page in doc), doc.page_count
    finally:
        doc.close()


# ------------------------- 3.3 OCR（可选） -------------------------
_OCR_PROBE = None


def ocr_provider() -> str | None:
    """低成本探测（只查包在不在，不加载模型）。"""
    global _OCR_PROBE
    if _OCR_PROBE is not None:
        return _OCR_PROBE or None
    if _have("paddleocr"):
        _OCR_PROBE = "PaddleOCR"
    elif _have("rapidocr") or _have("rapidocr_onnxruntime"):
        _OCR_PROBE = "RapidOCR"
    elif _have("pytesseract"):
        _OCR_PROBE = "Tesseract"
    else:
        _OCR_PROBE = ""
    return _OCR_PROBE or None


OCR_HINT = (
    "未检测到 OCR 引擎（不影响电子版 PDF 发票）。如需识别图片/扫描件，任选其一：\n"
    "  A. RapidOCR（推荐，纯 pip、免编译）：pip install rapidocr onnxruntime\n"
    "  B. PaddleOCR（中文票据最准，但 Windows 上较难装）：pip install paddleocr paddlepaddle\n"
    "  C. Tesseract：先装主程序（勾 chi_sim），再 pip install pytesseract\n"
    "注意：必须装到「你运行本脚本的同一个 Python」里。未安装时图片会被登记为「待人工录入」。"
)


def _ocr_worker_main(paths: list) -> int:
    """
    OCR 子进程入口（由 --ocr-worker 调用）。

    为什么放子进程：onnxruntime 等原生推理库在老 Windows 上可能 import 即原生崩溃
    （0xC0000005），Python 层拦不住。隔离后崩溃只影响当前文件，整批不受影响。
    """
    import numpy as np                              # type: ignore
    from PIL import Image                           # type: ignore

    name, engine = None, None
    prov = ocr_provider()
    try:
        if prov == "PaddleOCR":
            from paddleocr import PaddleOCR         # type: ignore

            engine = PaddleOCR(use_angle_cls=True, lang="ch", show_log=False)
        elif prov == "RapidOCR":
            try:
                from rapidocr import RapidOCR       # type: ignore
            except ImportError:
                from rapidocr_onnxruntime import RapidOCR  # type: ignore
            engine = RapidOCR()
        elif prov == "Tesseract":
            import pytesseract                      # type: ignore

            engine = pytesseract
        name = prov
    except Exception as exc:                        # noqa: BLE001
        print(json.dumps({"ok": False, "error": f"引擎加载失败：{exc}"},
                         ensure_ascii=True))
        return 0

    if not name:
        print(json.dumps({"ok": False, "error": "没有可用的 OCR 引擎"},
                         ensure_ascii=True))
        return 0

    texts = []
    for p in paths:
        try:
            img = Image.open(p)
            img.load()
            if name == "PaddleOCR":
                res = engine.ocr(np.array(img.convert("RGB")), cls=True)
                lines = []
                for page in res or []:
                    for item in page or []:
                        try:
                            lines.append(item[1][0])
                        except Exception:               # noqa: BLE001
                            pass
                texts.append("\n".join(lines))
            elif name == "RapidOCR":
                out = engine(np.array(img.convert("RGB")))
                got = getattr(out, "txts", None)
                if got:
                    texts.append("\n".join(str(t) for t in got if t))
                    continue
                if isinstance(out, tuple):
                    out = out[0]
                lines = []
                for item in out or []:
                    try:
                        t = item[1]
                        if isinstance(t, (list, tuple)):
                            t = t[0]
                        if t:
                            lines.append(str(t))
                    except (IndexError, TypeError):
                        pass
                texts.append("\n".join(lines))
            else:
                texts.append(engine.image_to_string(img, lang="chi_sim+eng"))
        except Exception as exc:                    # noqa: BLE001
            texts.append("")
    print(json.dumps({"ok": True, "engine": name, "texts": texts},
                     ensure_ascii=True))
    return 0


_CRASH_CODES = {
    -1073741819: "0xC0000005 访问违例（原生库崩溃）",
    3221225477: "0xC0000005 访问违例（原生库崩溃）",
    3221226505: "0xC0000409 栈溢出/保护性终止",
    -1073741510: "0xC000013A 被强制中断",
}


def _ocr_images(images: list, cfg: dict | None = None) -> str:
    """把一组 PIL 图片丢给 OCR 子进程，返回合并文本。超时取配置的 ocr_timeout。"""
    import subprocess
    import tempfile

    timeout = int((cfg or {}).get("ocr_timeout", 600) or 600)

    if not ocr_provider():
        raise RuntimeError("未安装 OCR 引擎。" + OCR_HINT)

    tmpdir = Path(tempfile.mkdtemp(prefix="inv_ocr_"))
    try:
        paths = []
        for i, img in enumerate(images):
            fp = tmpdir / f"p{i}.png"
            img.convert("RGB").save(fp, "PNG")
            paths.append(str(fp))
        cmd = [sys.executable, "-u", str(Path(__file__).resolve()),
               "--ocr-worker", *paths]
        try:
            proc = subprocess.run(cmd, capture_output=True, timeout=timeout,
                                  text=True, encoding="utf-8", errors="replace")
        except Exception:
            proc = subprocess.run(cmd, capture_output=True, timeout=timeout)
        payload = None
        for line in reversed((proc.stdout or "").splitlines()):
            line = line.strip()
            if line.startswith("{"):
                try:
                    payload = json.loads(line)
                    break
                except ValueError:
                    continue
        if payload is None:
            reason = _CRASH_CODES.get(proc.returncode, f"退出码 {proc.returncode}")
            raise RuntimeError(
                f"OCR 子进程未返回结果：{reason}。通常是 OCR 引擎与本机系统不兼容。\n"
                f"{('子进程输出：' + (proc.stderr or '').strip()[-200:]) if proc.stderr else ''}")
        if not payload.get("ok"):
            raise RuntimeError(f"OCR 失败：{payload.get('error')}")
        return "\n".join(payload.get("texts") or [])
    finally:
        import shutil

        shutil.rmtree(tmpdir, ignore_errors=True)


def ocr_preprocess(img):
    """
    灰度高对比 + 锐化 + 放大 —— 多份 GitHub 方案都靠这一步把 OCR 准确率抬上去。
    返回若干"候选图"，逐份识别后取质量最好的那份。
    """
    from PIL import Image, ImageEnhance, ImageFilter           # type: ignore

    out = []
    g = img.convert("L")
    if min(g.width, g.height) < 900:
        scale = max(1.0, 1600 / max(min(g.width, g.height), 1))
        g = g.resize((int(g.width * scale), int(g.height * scale)), Image.LANCZOS)
    out.append(g)
    out.append(ImageEnhance.Contrast(g).enhance(2.0).filter(ImageFilter.SHARPEN))
    try:
        import numpy as np                                     # type: ignore

        arr = np.asarray(g).astype("float32")
        if arr.mean() < 128:                                   # 偏暗 → 反相（白底黑字更利识别）
            out.append(Image.fromarray((255 - arr).astype("uint8")))
    except Exception:                                          # noqa: BLE001
        pass
    return out


# ------------------------- 3.4 文本质量闸门 -------------------------
# 关键字段标签探针 —— 用于判断"这份文本能不能用来解析发票"
LABEL_PROBE = ("发票号码", "开票日期", "价税合计", "购买方", "销售方",
               "税额", "发票代码", "校验码", "合计")


def text_quality(text: str) -> dict:
    """
    给"提取出来的文本"打分，用来决定是否需要升级提取方式。
    只看字符构成，不需要任何外部依赖。

    评分看重四件事：够不够长、可打印比例、汉字占比、数字占比。
    典型反面例子：OCR 把票据识别成一堆 `|||| ~~~` 或全体方块字。

    另外做一次**关键标签探针**：这一步能抓住"字符全都正常、可打印比例也很高，
    但标签与取值被拆散在不同文本块里"的文本 —— 光看可打印比例是发现不了的。
    （真实案例：PyMuPDF 按区块排序输出时，12 份样本的标签与数值被分开，
      文本质量仍判 1.0，但解析器一个字段都取不到。）
    """
    t = (text or "").strip()
    n = len(t)
    if n == 0:
        return {"score": 0.0, "len": 0, "labels": 0, "reasons": ["空文本"]}

    printable = sum(1 for c in t if c.isprintable() or c in "\n\t")
    cjk = sum(1 for c in t if "\u4e00" <= c <= "\u9fff")
    digit = sum(1 for c in t if c.isdigit())
    punct_noise = sum(1 for c in t if c in "|~^`_{}[]<>\\")

    ratio_print = printable / n
    ratio_cjk = cjk / n
    ratio_digit = digit / n
    ratio_noise = punct_noise / n

    score = 0.0
    reasons = []
    score += 0.35 * min(n / 200.0, 1.0)
    if n < 30:
        reasons.append(f"文本过短({n}字符)")
    score += 0.25 * ratio_print
    if ratio_print < 0.95:
        reasons.append(f"含不可打印字符({ratio_print:.0%}可打印)")
    # 中文票：汉字或数字总得占一定比例
    meaningful = ratio_cjk + ratio_digit
    score += 0.40 * min(meaningful / 0.35, 1.0)
    if meaningful < 0.15:
        reasons.append(f"有效字符占比过低(汉字+数字={meaningful:.0%})")
    score -= min(ratio_noise * 1.5, 0.25)
    if ratio_noise > 0.08:
        reasons.append(f"噪声符号偏多({ratio_noise:.0%})")
    # 出现替换符 / 方块字，说明字体映射没解出来
    if t.count("\ufffd") or t.count("\u25a1"):
        score -= 0.20
        reasons.append("存在替换符/方块字（字体映射失败）")

    label_hits = sum(1 for k in LABEL_PROBE if k in t)
    if label_hits < 2:
        score -= 0.35
        reasons.append(f"关键字段标签缺失（命中 {label_hits}/{len(LABEL_PROBE)}）")

    return {"score": round(max(0.0, min(1.0, score)), 3), "len": n,
            "cjk": cjk, "digit": digit, "noise": round(ratio_noise, 3),
            "labels": label_hits, "reasons": reasons}


# ------------------------- 3.5 统一入口 -------------------------
def _pdf_text_views(path, cfg: dict | None = None) -> tuple:
    """
    收集所有可用的 PDF 文本视图，返回 (views, pages)，views 为 [(来源, 文本), ...]。

    为什么要多个视图：PyMuPDF 的 `get_text("text")` 按文本**区块**排序输出，
    遇到"标签在左列、取值在右列"的版式时，标签与数值会被拆进不同区块、在正文里
    相距很远，解析器的"标签紧贴取值"假设就失效了；而本模块内置的标准库解析器按
    内容流**绘制顺序**输出，标签紧邻数值。两者对不同版式各有胜负。
    真实案例：某批样本在 PyMuPDF 视图下 12 张全部解析失败（置信度掉到 29%），
    换成标准库视图全部正常 —— 所以两个都抽出来，交给上层解析后择优。
    """
    views, pages = [], 0
    if _have("pymupdf") or _have("fitz"):
        try:
            t, p = pdf_text_pymupdf(path)
            pages = max(pages, p)
            if t.strip():
                views.append(("pdf-text(pymupdf)", t))
        except Exception:                                  # noqa: BLE001
            pass
    try:
        t, p = pdf_text_stdlib(path, cfg)
        pages = max(pages, p)
        if t.strip():
            views.append(("pdf-text(stdlib)", t))
    except Exception:                                      # noqa: BLE001
        pass

    # 两个抽取器给出的文本完全相同时只保留一份，避免无意义的重复解析
    seen, uniq = set(), []
    for src, t in views:
        k = hash(t)
        if k not in seen:
            seen.add(k)
            uniq.append((src, t))
    return uniq, pages


def extract_text(path, cfg: dict) -> dict:
    """
    返回 {text, alts, source, pages, quality, notes}；失败抛 RuntimeError。

    `text` 是按"文本质量"选出的主视图，`alts` 是其余可用视图 ——
    上层会把所有视图都解析一遍再择优（见 _pdf_text_views 的说明）。
    """
    path = Path(path)
    if not path.exists():
        raise RuntimeError(f"文件不存在：{path}")
    ext = path.suffix.lower()
    if ext not in SUPPORTED_EXTS:
        raise RuntimeError(f"不支持的文件类型：{ext}")

    notes = []

    # ---- 图片：直接 OCR ----
    if ext in IMAGE_EXTS:
        return _extract_image(path, cfg, notes)

    # ---- PDF：先收集全部文本视图 ----
    views, pages = _pdf_text_views(path, cfg)

    def _ret(txt, source, quality, extra_alts):
        return {"text": txt, "alts": extra_alts, "source": source, "pages": pages,
                "quality": quality, "notes": "；".join(notes)}

    if not views:
        # 完全没有文本层 → 扫描件分支
        if cfg.get("ocr_enabled") == "off":
            raise RuntimeError("该 PDF 没有文本层（疑似扫描件），且配置里关闭了 OCR。")
        if not ocr_provider():
            raise RuntimeError("该 PDF 没有可用文本层（疑似扫描件）。" + OCR_HINT)
        ocr_text, ocr_pages = _pdf_ocr(path, cfg)
        pages = max(pages, ocr_pages)
        return _ret(ocr_text, f"ocr({ocr_provider()})",
                    text_quality(ocr_text)["score"], [])

    # 按文本质量排序定主视图；同分时优先 PyMuPDF（对复杂版式更稳）
    ranked = sorted(
        ((text_quality(t)["score"] + (0.01 if "pymupdf" in s else 0.0), i, s, t)
         for i, (s, t) in enumerate(views)),
        key=lambda x: (-x[0], x[1]))
    src, text = ranked[0][2], ranked[0][3]
    alts = [(s, t) for _sc, _i, s, t in ranked[1:]]
    q = text_quality(text)
    thick = len(text.strip()) >= int(cfg["min_text_chars"])

    if thick and q["score"] >= float(cfg["text_quality_min"]):
        return _ret(text, src, q["score"], alts)

    # ---- 文本层不可用或质量差 → 尝试 OCR ----
    if cfg.get("ocr_enabled") == "off":
        if not thick:
            raise RuntimeError(
                "PDF 没有可用文本层，且配置里关闭了 OCR。"
                + (OCR_HINT if not ocr_provider() else "可将 ocr_enabled 改回 auto。"))
        notes.append(f"文本质量偏低({q['score']}：{'、'.join(q['reasons'] or ['-'])}），"
                     f"但已按配置跳过 OCR")
        return _ret(text, src, q["score"], alts)

    if not ocr_provider():
        if not thick:
            raise RuntimeError("该 PDF 没有可用文本层（疑似扫描件）。" + OCR_HINT)
        notes.append(f"文本质量偏低({q['score']}：{'、'.join(q['reasons'] or ['-'])}），"
                     f"未装 OCR，沿用文本层结果")
        return _ret(text, src, q["score"], alts)

    try:
        ocr_text, ocr_pages = _pdf_ocr(path, cfg)
        oq = text_quality(ocr_text)
        pages = max(pages, ocr_pages)
        if not thick or oq["score"] > q["score"]:
            notes.append(f"文本层质量 {q['score']} → OCR 质量 {oq['score']}，已采用 OCR 结果")
            return _ret(ocr_text, f"ocr({ocr_provider()})", oq["score"],
                        [(src, text)] + alts)
        notes.append(f"文本层质量 {q['score']} ≥ OCR 质量 {oq['score']}，保留文本层")
        return _ret(text, src, q["score"], alts + [("ocr", ocr_text)])
    except Exception as exc:                         # noqa: BLE001
        if not thick:
            raise RuntimeError(f"扫描件 OCR 失败：{exc}") from exc
        notes.append(f"OCR 尝试失败（{exc}），沿用文本层结果")
        return _ret(text, src, q["score"], alts)


def _pdf_ocr(path, cfg) -> tuple:
    from PIL import Image                                  # type: ignore
    try:
        import pymupdf as fitz                             # type: ignore
    except ImportError:
        import fitz                                        # type: ignore

    doc = fitz.open(str(path))
    try:
        zoom = float(cfg.get("ocr_zoom", 2.5))
        mat = fitz.Matrix(zoom, zoom)
        raw = [Image.open(__import__("io").BytesIO(
            page.get_pixmap(matrix=mat).tobytes("png")))
            for page in doc]
        pages = doc.page_count
    finally:
        doc.close()

    best, best_q = "", -1.0
    for img in raw:
        for cand in ocr_preprocess(img):
            try:
                t = _ocr_images([cand])
            except Exception:                              # noqa: BLE001
                continue
            q = text_quality(t)["score"]
            if q > best_q:
                best, best_q = t, q
            if best_q >= 0.85:                             # 已经很好，不必穷举
                break
    if not best.strip():
        raise RuntimeError("OCR 未识别到任何文字")
    return best, pages


def _extract_image(path, cfg, notes) -> dict:
    try:
        from PIL import Image                              # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "识别图片需要 Pillow（本机未安装）。安装后即可自动识别：\n"
            "    pip install pillow\n"
            "（图片无法用纯标准库解码，这是唯一的硬依赖；PDF 不需要它。）"
        ) from exc

    try:
        img = Image.open(str(path))
        img.load()
    except Exception as exc:                               # noqa: BLE001
        raise RuntimeError(f"图片打开失败：{exc}") from exc

    best, best_q = "", -1.0
    for cand in ocr_preprocess(img):
        try:
            t = _ocr_images([cand], cfg)
        except Exception as exc:                           # noqa: BLE001
            notes.append(str(exc))
            break
        q = text_quality(t)["score"]
        if q > best_q:
            best, best_q = t, q
        if best_q >= 0.85:
            break
    if not best.strip():
        raise RuntimeError("图片未识别到任何文字（请确认清晰度，或换电子版发票）")
    return {"text": best, "alts": [], "source": f"ocr({ocr_provider()})", "pages": 1,
            "quality": best_q, "notes": "；".join(notes)}


# ==========================================================================
# 4. 字段解析：模板优先 + 内置启发式兜底 + 逐字段置信度
# ==========================================================================
NAME_STOP = (
    r"(?=销售方|购买方|销货方|购货方|统一社会信用代码|纳税人识别号|税号|"
    r"项目名称|货物或应税|规格型号|单位|数量|单价|金额|税率|征收率|价税合计|"
    r"合计|备注|开票人|收款人|复核人|地址|开户行|电话|账号|$)"
)

_RATE_RE = re.compile(r"(?<![\d.])(\d[\d.]*)%")
_MONEY_PREFIX = re.compile(r"^\d[\d,]*\.\d{2}")


def _strip_money_prefix(tok: str) -> str:
    """从粘连 token 左侧逐段剥离"两位小数的金额"，剩下的是税率。"""
    rest = tok
    while True:
        m = _MONEY_PREFIX.match(rest)
        if not m:
            break
        nxt = rest[m.end():]
        if not nxt:
            break
        rest = nxt
    return rest


def extract_rates(one: str) -> list:
    out = []
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


# 这些票种票面没有「销售方名称 / 销售方税号 / 发票代码」栏位
# （承运人信息不在购销方栏，票号本身就是 20 位数电号），
# 因此 R01(销售方名称) / R11(缺发票代码) / R12(销售方税号缺失) 对它们不适用。
NO_SELLER_TYPES = (
    "电子发票(铁路电子客票)",
    "电子发票(航空运输电子客票行程单)",
    "电子发票(通行费)",
)


def _is_edigital(itype) -> bool:
    """是否属于「全面数字化电子发票」体系（数电票及其各专用票种）。"""
    t = str(itype or "")
    return "数电" in t or "全电" in t or t.startswith("电子发票(")


def _relaxed_type(itype) -> bool:
    """是否为「无销售方栏位」的票种。"""
    t = str(itype or "")
    return any(k in t for k in NO_SELLER_TYPES)


def detect_type(one: str) -> str:
    """
    识别票种。

    判断顺序很重要（v2.1 调整两处）：
      · 新增的三个数电票种（铁路电子客票 / 航空运输电子客票行程单 / 通行费）
        标题里同样含「电子发票」，必须排在通用的「电子发票+普通/专用」之前，
        否则会被降级成「电子发票(其他)」并因找不到销售方而误报异常；
      · 通用「数电 / 全电」判断被移到专/普判断之后 —— 旧实现把它放在第一位，
        票面备注里只要出现「数电」二字（例如红冲说明），专票就会退化成裸的
        「数电发票」，丢掉专/普区分（已实测复现）。

    只匹配**标题级**写法（如「电子发票(铁路电子客票)」），不匹配孤立的
    「通行费 / 行程单」等词，避免把明细行或备注里的字样误判成票种。
    """
    if "电子发票(铁路电子客票)" in one or "铁路电子客票" in one:
        return "电子发票(铁路电子客票)"
    if "电子发票(航空运输电子客票行程单)" in one \
            or "航空运输电子客票行程单" in one:
        return "电子发票(航空运输电子客票行程单)"
    if "通行费电子发票" in one or "电子发票(通行费)" in one \
            or "收费公路通行费" in one:
        return "电子发票(通行费)"
    if "电子发票" in one and "专用发票" in one:
        return "数电发票(增值税专用发票)"
    if "电子发票" in one and "普通发票" in one:
        return "数电发票(普通发票)"
    if "数电" in one or "全电" in one:
        return "数电发票"
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


def _clean_name(v) -> str:
    if not v:
        return ""
    v = re.sub(r"^[:：\s]+", "", v)
    v = re.split(r"统一社会信用代码|纳税人识别号|销售方|购买方|销货方|购货方|"
                 r"项目名称|货物或应税|规格型号|地址|开户行|电话|账号|备注|开票人", v)[0]
    v = re.sub(r"[:：\s]+$", "", v)
    v = re.sub(r"^[（(].*?[)）]", "", v)
    return v.strip(" :：,，、")


def _scan_section(section: list, pattern: str):
    for ln in section:
        m = re.search(pattern, ln)
        if m:
            return m.group(1), ln
    return "", ""


def _party(one: str, lines: list, who: str):
    """购销双方双通道解析：整行视图（对竖直排版最稳）+ 按行视图。"""
    is_buyer = who == "购买方"
    other = "销售方" if is_buyer else "购买方"
    aliases = (["购买方", "购货方", "付款方"] if is_buyer
               else ["销售方", "销货方", "收款方"])

    name = tax = ""
    ev_name = ev_tax = ""

    for al in aliases:
        if name:
            break
        m = re.search(al + r"(?:信息)?名称[:：]?(.{2,50}?)" + NAME_STOP, one)
        if m:
            name = _clean_name(m.group(1))
            ev_name = m.group(0)
    if not name:
        for al in aliases:
            m = re.search(al + r"(?:信息)?[:：]?名称[:：]?(.{2,50}?)" + NAME_STOP, one)
            if m:
                name = _clean_name(m.group(1))
                ev_name = m.group(0)
                break

    if not name or not tax:
        b_i = s_i = None
        for i, ln in enumerate(lines):
            if b_i is None and any(a in ln for a in ("购买方", "购货方", "付款方")):
                b_i = i
            if s_i is None and any(a in ln for a in ("销售方", "销货方", "收款方")):
                s_i = i
        start = b_i if is_buyer else s_i
        if start is not None:
            end = s_i if (is_buyer and s_i is not None and s_i > start) \
                else min(len(lines), start + 8)
            section = lines[start:end]
            if not name:
                nm, ev = _scan_section(section, r"名称[:：]?(.+)$")
                if nm:
                    name = _clean_name(nm)
                    ev_name = ev
            if not tax:
                tx, ev = _scan_section(
                    section,
                    r"(?:统一社会信用代码|纳税人识别号|税号)[/]?[:：]?([0-9A-Z]{15,20})")
                if tx:
                    tax, ev_tax = tx, ev

    if not tax:
        tax_re = re.compile(
            re.escape(who) + r"(?:信息)?.{0,120}?(?:统一社会信用代码|纳税人识别号|税号)"
                            r"[/]?[:：]?([0-9A-Z]{15,20})")
        m = tax_re.search(one)
        if m:
            cand = m.group(1)
            pos, opp = one.find(cand), one.find(other)
            if not (opp != -1 and pos > opp and one.find(who) < opp):
                tax, ev_tax = cand, m.group(0)

    return name, tax, ev_name, ev_tax


# 一张新发票的起始行：必须「以票据编号标签开头」且标签后紧跟票据号数字。
# 这一条必须严格 —— 旧实现只判 `"发票号码" in ln`，于是
# 「备注：原发票号码 2631… 已作废」这种行也会触发切分，
# 把一张完整发票切出一个只有备注的碎片，凭空多出一行台账（实测已复现）。
_INV_START_RE = re.compile(r"^[（(\[【]*(?:发票号码|票据号码|发票No)[:：]?\s*\d{8,20}")


def _looks_like_title(ln: str) -> bool:
    return bool(re.match(r"^(电子发票|全电发票|增值税|机动车销售|二手车销售)", ln)) \
        and "发票" in ln and "号码" not in ln and "代码" not in ln


def _is_invoice_start(ln: str) -> bool:
    return bool(_INV_START_RE.match(ln))


def _split_multi_invoice(t: str) -> list:
    """
    一个文件含多张发票时切段，保证"一张发票一行"。

    两层防误切（v2.1 加固）：
      1. 切分锚点从严 —— 只有「以编号标签开头」的行才算新票开始（见 _INV_START_RE）；
      2. 碎片回收 —— 切出来的段落若既无票据号也无金额要素，判定为备注/页脚碎片，
         并回相邻段落，绝不单独成行。
    """
    per_line = re.sub(r"\s+", "", t)
    hits = [m.start() for m in re.finditer(r"发票号码", per_line)]
    titles = [m.start() for m in re.finditer(r"(电子发票|全电发票|增值税)", per_line)]
    if len(hits) <= 1 and len(titles) <= 1:
        return [t]

    segs, cur = [], []
    for ln in to_lines(t):
        joined = "".join(cur)
        if cur and (("价税合计" in joined or "税额" in joined)
                    and "开票日期" in joined) \
                and (_is_invoice_start(ln) or _looks_like_title(ln)):
            segs.append(cur)
            cur = []
        cur.append(ln)
    if cur:
        segs.append(cur)

    segs = _merge_fragments(segs)
    return ["\n".join(s) for s in segs] if len(segs) > 1 else [t]


def _merge_fragments(segs: list) -> list:
    """把不含票据要素的碎片段落并回相邻段落（前导碎片并入下一段）。"""
    merged, pending = [], []
    for s in segs:
        joined = "".join(s)
        if "发票号码" not in joined and "价税合计" not in joined \
                and "税额" not in joined:
            pending.extend(s)                       # 碎片：先挂起
            continue
        merged.append(pending + list(s))
        pending = []
    if pending:
        if merged:
            merged[-1].extend(pending)
        else:
            merged.append(pending)
    return merged


def _builtin_parse(t: str, conf: dict, ev: dict) -> dict:
    """内置启发式解析（v1 的成熟规则，作为模板未覆盖时的兜底）。"""
    lines = to_lines(t)
    one = to_oneline(t)
    rec = {k: "" for k, _, _ in FIELDS}

    def put(key, val, c, evidence):
        if val in ("", None):
            return
        rec[key] = val
        conf[key] = max(conf.get(key, 0.0), c)
        if evidence:
            ev[key] = evidence

    m = re.search(r"发票号码[:：]?\s*(\d{8,20})", one)
    if m:
        put("invoice_no", m.group(1), 0.95, m.group(0))
    m = re.search(r"发票代码[:：]?\s*(\d{10,12})", one)
    if m:
        put("invoice_code", m.group(1), 0.93, m.group(0))

    m = re.search(r"开票日期[:：]?\s*(\d{4})[年\-/.]?(\d{1,2})[月\-/.]?(\d{1,2})日?", one)
    if m:
        y, mo, d = (int(g) for g in m.groups())
        try:
            put("issue_date", date(y, mo, d).isoformat(), 0.95, m.group(0))
        except ValueError:
            put("issue_date", f"{y:04d}-{mo:02d}-{d:02d}", 0.35, m.group(0))
            rec["_bad_date"] = True

    m = re.search(r"校验码[:：]?\s*(\d{15,25})", one)
    if m:
        put("check_code", m.group(1), 0.9, m.group(0))

    for key, who in (("buyer", "购买方"), ("seller", "销售方")):
        nm, tx, ev_nm, ev_tx = _party(one, lines, who)
        put(f"{key}_name", nm, 0.85, ev_nm)
        put(f"{key}_tax", tx, 0.88, ev_tx)

    amount = tax = total = None
    m = re.search(r"价税合计[^¥￥]{0,60}?[¥￥]?(" + MONEY_STR + r")", one)
    if not m:
        m = re.search(r"[(（]小写[)）][¥￥]?(" + MONEY_STR + r")", one)
    if not m:
        m = re.search(r"价税合计[^0-9]{0,60}(" + MONEY_STR + r")", one)
    if not m:
        # 铁路电子客票 / 航空行程单用「票价」而不是「价税合计」（v2.1 新增）
        m = re.search(r"票价[^0-9¥￥]{0,40}?[¥￥]?(" + MONEY_STR + r")", one)
    if m:
        total = to_money(m.group(1))
        put("total", total, 0.92, m.group(0))

    # 票面大写金额（用于 R17 与「小写」交叉勾稽）
    m = re.search(r"价税合计[（(]大写[)）][:：]?([^0-9¥￥]{1,40}?)[（(]小写", one)
    if not m:
        m = re.search(r"价税合计[（(]大写[)）][:：]?([零壹贰叁肆伍陆柒捌玖〇一二三四五六七八九"
                      r"拾佰仟万亿元角分整正]{2,40})", one)
    if m:
        cn_val = cn_amount_to_number(m.group(1))
        if cn_val is not None:
            rec["_cn_amount"] = cn_val
            ev["_cn_amount"] = m.group(0)[:80]

    m = re.search(r"合计[¥￥](" + MONEY_STR + r")[¥￥](" + MONEY_STR + r")", one)
    if m:
        amount, tax = to_money(m.group(1)), to_money(m.group(2))
        put("amount", amount, 0.9, m.group(0))
        put("tax", tax, 0.9, m.group(0))
    else:
        m = re.search(r"合计金额[¥￥]?(" + MONEY_STR + r")", one)
        if m:
            amount = to_money(m.group(1))
            put("amount", amount, 0.85, m.group(0))
        m = re.search(r"(?:合计税额|税额合计)[¥￥]?(" + MONEY_STR + r")", one)
        if m:
            tax = to_money(m.group(1))
            put("tax", tax, 0.85, m.group(0))
        if amount is None and tax is None:
            pairs = re.findall(r"[¥￥](" + MONEY_STR + r")[¥￥](" + MONEY_STR + r")", one)
            if pairs:
                amount, tax = to_money(pairs[-1][0]), to_money(pairs[-1][1])
                put("amount", amount, 0.55, "兜底匹配相邻两金额")
                put("tax", tax, 0.55, "兜底匹配相邻两金额")

    # 用价税合计反推缺失项 —— 置信度必须压低：这是"算出来的"不是"读出来的"
    if total is not None and tax is not None and amount is None:
        put("amount", round(total - tax, 2), 0.6, "由 价税合计−税额 推导")
    if total is not None and amount is not None and tax is None:
        put("tax", round(total - amount, 2), 0.6, "由 价税合计−金额 推导")

    rates = extract_rates(one)
    if not rates:
        for kw in ("免税", "不征税"):
            if kw in one:
                rates.append(kw)
    if rates:
        put("tax_rate", "/".join(rates), 0.85, "、".join(rates))

    items = []
    for ln in lines:
        for it in re.findall(r"\*[^*]{1,20}\*[^*0-9]{0,30}", ln):
            it = it.strip()
            if it.endswith("*"):
                it = it[:-1]
            if it and it not in items:
                items.append(it)
    if items:
        put("item_name", "; ".join(items)[:120], 0.9, items[0])
    else:
        m = re.search(r"(?:货物或应税劳务、?服务名称|项目名称)[:：]?([^\d¥￥]{2,30})"
                      + NAME_STOP, one)
        if m:
            put("item_name", m.group(1), 0.7, m.group(0))

    m = re.search(r"开票人[:：]?([\u4e00-\u9fa5A-Za-z]{2,8})", one)
    if not m:
        m = re.search(r"开票人[:：]?([\u4e00-\u9fa5A-Za-z]{2,8})(?=收款人|复核人|$)", one)
    if m:
        put("drawer", m.group(1), 0.85, m.group(0))

    # 新数电票种没有「货物或应税劳务」明细行，用「车次 / 航班 + 区间」充当
    # 『主要项目/货物名称』，让台账这一列对它们同样有意义（v2.1 新增）。
    if not rec.get("item_name"):
        kind = detect_type(one)
        if kind == "电子发票(铁路电子客票)":
            tr = re.search(r"车次[:：]?([GDCZTKY]?\d{1,5})", one)
            stations = re.findall(r"([\u4e00-\u9fa5]{2,8}站)", one)
            parts = ([(tr.group(1) + "次")] if tr else [])
            if len(stations) >= 2:
                parts.append("→".join(stations[:2]))
            if parts:
                put("item_name", " ".join(parts), 0.8, "车次/区间")
        elif kind == "电子发票(航空运输电子客票行程单)":
            fl = re.search(r"航班号[:：]?([A-Z]{2}\d{3,4})", one)
            if fl:
                put("item_name", fl.group(1), 0.8, fl.group(0))

    return rec


def _score_template(tpl: dict, one: str) -> int:
    """模板匹配度：命中 keywords 越多越优先（借鉴 invoice2data 的 keywords 选择）。"""
    if any(ex and ex in one for ex in (tpl.get("exclude_keywords") or [])):
        return -1
    hits = sum(1 for kw in (tpl.get("keywords") or []) if kw and kw in one)
    return hits


def _apply_template(tpl: dict, one: str, rec: dict, conf: dict, ev: dict) -> None:
    """模板里的字段正则命中就覆盖（模板是"更懂这种版式"的先验知识）。"""
    for key, patterns in (tpl.get("fields") or {}).items():
        if key not in rec:
            continue
        for pat in patterns or []:
            try:
                m = re.search(pat, one)
            except re.error as exc:
                ev[f"!{key}"] = f"正则错误：{exc}"
                continue
            if not m:
                continue
            val = m.group(1) if m.groups() else m.group(0)
            if key in MONEY_KEYS:
                val = to_money(val)
                if val is None:
                    continue
            if key == "issue_date":
                d = _parse_date_flex(val)
                if not d:
                    continue
                val = d
            if val in ("", None):
                continue
            rec[key] = val
            conf[key] = 0.93
            ev[key] = f"[模板:{tpl.get('name')}] {m.group(0)[:80]}"
            break
    for key, val in (tpl.get("static") or {}).items():
        if key in rec and val not in ("", None):
            rec[key] = val
            conf[key] = max(conf.get(key, 0.0), 0.9)
            ev[key] = f"[模板静态] {val}"


def _parse_date_flex(s: str) -> str:
    s = re.sub(r"\s+", "", str(s or ""))
    m = re.match(r"^(\d{4})[年\-/.]?(\d{1,2})[月\-/.]?(\d{1,2})日?$", s)
    if not m:
        return ""
    y, mo, d = (int(g) for g in m.groups())
    try:
        return date(y, mo, d).isoformat()
    except ValueError:
        return f"{y:04d}-{mo:02d}-{d:02d}"


def parse_invoice(text: str, file_name: str, templates: list) -> list:
    """返回发票记录列表（一个文件可能含多张）。"""
    normalized = norm_text(text)
    segments = _split_multi_invoice(normalized)
    doc_type = detect_type(to_oneline(normalized))

    records, seen = [], set()
    for idx, seg in enumerate(segments):
        one = to_oneline(seg)
        conf, ev = {}, {}

        # (1) 内置启发式打底
        rec = _builtin_parse(seg, conf, ev)

        # (2) 模板覆盖（按 keywords 命中数排序，取最匹配的一个）
        scored = sorted(((_score_template(tpl, one), i, tpl)
                         for i, tpl in enumerate(templates)),
                        key=lambda x: (-x[0], x[1]))
        for s, _i, tpl in scored:
            if s <= 0:
                continue
            _apply_template(tpl, one, rec, conf, ev)
            break

        # (3) 类型回退
        if rec.get("invoice_type") in ("", "未知类型", "电子发票(其他)"):
            if doc_type != "未知类型":
                rec["invoice_type"] = doc_type
        if not rec.get("invoice_type"):
            rec["invoice_type"] = doc_type
        conf["invoice_type"] = 0.9 if rec["invoice_type"] != "未知类型" else 0.2
        ev["invoice_type"] = rec["invoice_type"]

        no = rec.get("invoice_no") or ""
        if no and no in seen:
            continue                                     # 同文件内的重复副本
        if no:
            seen.add(no)

        rec["file_name"] = file_name
        rec["_confidence_map"] = conf
        rec["_evidence"] = ev
        rec["_segment"] = idx + 1
        rec["_segment_total"] = len(segments)
        records.append(rec)

    if len(records) > 1:
        for i, r in enumerate(records, 1):
            r["file_name"] = f"{file_name} [第{i}张/共{len(records)}张]"
    return records


# ==========================================================================
# 5. 校验规则注册表（可开关、可扩展）
# ==========================================================================
RULES = []


def rule(rid: str, level: str, title: str):
    def deco(fn):
        RULES.append({"id": rid, "level": level, "title": title, "fn": fn})
        return fn
    return deco


USCC_CHARS = "0123456789ABCDEFGHJKLMNPQRTUWXY"
USCC_WEIGHTS = [1, 3, 9, 27, 19, 26, 16, 17, 20, 29, 25, 13, 8, 24, 10, 30, 28]


def uscc_check(code: str):
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


# 20 位数电发票号码：前 2 位年份 + 第 3-4 位省级区域码
PROVINCE_CODES = {
    "11", "12", "13", "14", "15", "21", "22", "23", "31", "32", "33", "34",
    "35", "36", "37", "41", "42", "43", "44", "45", "46", "50", "51", "52",
    "53", "54", "61", "62", "63", "64", "65", "71",
}


@rule("R01", "严重", "必填字段缺失")
def r01(ctx):
    out = []
    # 铁路客票 / 航空行程单 / 通行费票面本就没有「销售方」栏位（承运人信息不在
    # 购销方栏），对它们要求销售方名称只会制造必然的误报。
    relaxed = _relaxed_type(ctx["rec"].get("invoice_type"))
    for key, label, suggest in (
        ("invoice_no", "发票号码", "确认 PDF 是否完整原件，必要时人工补录"),
        ("issue_date", "开票日期", "检查票据是否清晰、是否被裁切"),
        ("total", "价税合计", "核对发票『价税合计(小写)』字样"),
        ("seller_name", "销售方名称", "销售方名称未识别，请人工补录"),
        ("buyer_name", "购买方名称", "购买方名称未识别（个人抬头可能为空）"),
    ):
        if relaxed and key == "seller_name":
            continue
        if ctx["rec"].get(key) in (None, "", []):
            out.append({"level": "严重", "field": label,
                        "message": f"未识别到「{label}」", "suggest": suggest})
    return out


@rule("R02", "严重", "发票号码含非数字")
def r02(ctx):
    no = str(ctx["rec"].get("invoice_no") or "")
    if no and not no.isdigit():
        return [{"level": "严重", "field": "发票号码",
                 "message": f"发票号码含非数字字符：{no}", "suggest": "人工核对"}]
    return []


@rule("R03", "警告", "发票号码位数与版式不符")
def r03(ctx):
    no = str(ctx["rec"].get("invoice_no") or "")
    itype = ctx["rec"].get("invoice_type") or ""
    if not no:
        return []
    if _is_edigital(itype) or len(no) == 20:
        if len(no) != 20:
            return [{"level": "警告", "field": "发票号码",
                     "message": f"数电发票号码应为 20 位，当前 {len(no)} 位",
                     "suggest": "核对号码是否被截断"}]
        return []
    if len(no) not in (8, 9, 10, 12):
        return [{"level": "警告", "field": "发票号码",
                 "message": f"发票号码为 {len(no)} 位，与常见版式不符",
                 "suggest": "人工核对"}]
    return []


@rule("R04", "严重", "开票日期非法或超出合理范围")
def r04(ctx):
    ds = str(ctx["rec"].get("issue_date") or "")
    if not ds:
        return []
    if ctx["rec"].get("_bad_date"):
        return [{"level": "严重", "field": "开票日期",
                 "message": f"开票日期不是合法日期：{ds}", "suggest": "人工核对"}]
    try:
        d = date.fromisoformat(ds)
    except ValueError:
        return [{"level": "严重", "field": "开票日期",
                 "message": f"开票日期无法解析：{ds}", "suggest": "人工核对"}]
    grace = int(ctx["cfg"].get("future_date_grace_days", 0))
    if (d - ctx["today"]).days > grace:
        return [{"level": "严重", "field": "开票日期",
                 "message": f"开票日期 {ds} 晚于当前日期", "suggest": "核对是否识别错误"}]
    if d < date(2000, 1, 1):
        return [{"level": "警告", "field": "开票日期",
                 "message": f"开票日期 {ds} 过早，疑似识别错误", "suggest": "人工核对"}]
    return []


@rule("R05", "严重", "金额勾稽不平衡")
def r05(ctx):
    rec, cfg = ctx["rec"], ctx["cfg"]
    a, t, tot = rec.get("amount"), rec.get("tax"), rec.get("total")
    if isinstance(a, (int, float)) and isinstance(t, (int, float)) \
            and isinstance(tot, (int, float)):
        diff = round(a + t - tot, 2)
        if abs(diff) > float(cfg["balance_tolerance"]):
            return [{"level": "严重", "field": "金额勾稽",
                     "message": f"金额(不含税){a:,.2f} + 税额{t:,.2f} = {a + t:,.2f}，"
                                f"与价税合计{tot:,.2f} 相差 {diff:,.2f}",
                     "suggest": "核对是否漏识别明细行或合计数被遮挡"}]
        return []
    if isinstance(tot, (int, float)) and (a is None or t is None):
        return [{"level": "警告", "field": "金额勾稽",
                 "message": "金额(不含税)或税额缺失，无法完成勾稽校验",
                 "suggest": "人工补充后核对"}]
    return []


@rule("R06", "提示", "红字/零金额发票")
def r06(ctx):
    tot = ctx["rec"].get("total")
    if not isinstance(tot, (int, float)):
        return []
    if tot < 0:
        return [{"level": "提示", "field": "价税合计",
                 "message": f"价税合计为负数（{tot:,.2f}），属于红字发票",
                 "suggest": "按红字发票入账，注意冲减方向"}]
    if tot == 0:
        return [{"level": "警告", "field": "价税合计", "message": "价税合计为 0.00",
                 "suggest": "核对是否为零金额发票或识别异常"}]
    return []


@rule("R07", "严重", "税号格式非法")
def r07(ctx):
    out = []
    for key, label in (("buyer_tax", "购买方税号"), ("seller_tax", "销售方税号")):
        code = str(ctx["rec"].get(key) or "").strip().upper()
        if not code:
            continue
        if not re.fullmatch(r"[0-9A-Z]{15,20}", code):
            out.append({"level": "严重", "field": label,
                        "message": f"税号格式非法：{code}", "suggest": "人工核对"})
        elif len(code) not in (15, 17, 18, 20):
            out.append({"level": "警告", "field": label,
                        "message": f"税号长度 {len(code)} 位较为罕见",
                        "suggest": "人工核对"})
    return out


@rule("R08", "警告", "统一社会信用代码校验位不通过")
def r08(ctx):
    out = []
    for key, label in (("buyer_tax", "购买方税号"), ("seller_tax", "销售方税号")):
        code = str(ctx["rec"].get(key) or "").strip().upper()
        if code and uscc_check(code) is False:
            out.append({"level": "警告", "field": label,
                        "message": f"统一社会信用代码校验位不通过：{code}",
                        "suggest": "疑似识别错误（0/O、1/I 混淆），建议人工核对"})
    return out


@rule("R09", "警告", "购销双方同名")
def r09(ctx):
    bn = (ctx["rec"].get("buyer_name") or "").strip()
    sn = (ctx["rec"].get("seller_name") or "").strip()
    if bn and sn and bn == sn:
        return [{"level": "警告", "field": "购销双方",
                 "message": "购买方与销售方名称完全相同",
                 "suggest": "核对是否填错，或属于自开自抵情形"}]
    return []


@rule("R10", "提示", "税率不在常用集合")
def r10(ctx):
    rate = str(ctx["rec"].get("tax_rate") or "")
    if not rate:
        return []
    for r in rate.split("/"):
        if r and r not in VALID_TAX_RATES:
            return [{"level": "提示", "field": "税率",
                     "message": f"税率 {r} 不在常用税率集合内",
                     "suggest": "确认是否为特殊征收率/减免政策"}]
    return []


@rule("R11", "警告", "非数电发票缺少发票代码")
def r11(ctx):
    itype = ctx["rec"].get("invoice_type") or ""
    if not _is_edigital(itype) and not ctx["rec"].get("invoice_code") \
            and ctx["rec"].get("invoice_no"):
        return [{"level": "警告", "field": "发票代码",
                 "message": "未识别到发票代码（非数电发票通常应有 12 位发票代码）",
                 "suggest": "人工核对"}]
    return []


@rule("R12", "警告", "销售方税号缺失")
def r12(ctx):
    if _relaxed_type(ctx["rec"].get("invoice_type")):
        return []           # 铁路客票 / 航空行程单 / 通行费票面无销售方税号栏位
    if not str(ctx["rec"].get("seller_tax") or "").strip():
        return [{"level": "警告", "field": "销售方税号",
                 "message": "销售方纳税人识别号缺失", "suggest": "人工补录"}]
    return []


# ---------- 以下 4 条是本版新增的"交叉校验"，专门抓组合型识别错误 ----------
@rule("R13", "严重", "税额与「金额×税率」不勾稽")
def r13(ctx):
    rec, cfg = ctx["rec"], ctx["cfg"]
    a, t = rec.get("amount"), rec.get("tax")
    rate_s = str(rec.get("tax_rate") or "")
    if not isinstance(a, (int, float)) or not isinstance(t, (int, float)):
        return []
    if a == 0 or not rate_s:
        return []
    rates = [r for r in rate_s.split("/") if r.endswith("%")]
    if len(rates) != 1:                     # 多税率/免税 → 本条不适用
        return []
    try:
        pct = float(rates[0].rstrip("%"))
    except ValueError:
        return []
    expect = abs(a) * pct / 100.0
    tol = max(float(cfg["rate_abs_floor"]), abs(a) * float(cfg["rate_rel_tolerance"]))
    if abs(abs(t) - expect) > tol:
        return [{"level": "严重", "field": "税率×金额",
                 "message": f"按税率 {rates[0]} 应有税额 {expect:,.2f}，"
                            f"实际税额 {t:,.2f}（差 {abs(t) - expect:,.2f}）",
                 "suggest": "税率或金额/税额识别有误，三者需自洽"}]
    return []


@rule("R14", "警告", "数电发票号码结构异常")
def r14(ctx):
    no = str(ctx["rec"].get("invoice_no") or "")
    if len(no) != 20 or not no.isdigit():
        return []
    # 注：旧实现此处有一行 `if "数电" not in itype and not no.startswith("0"): pass`
    # —— 一个空分支，等于没有判断（死代码）。为不改变既有检出行为，直接删除，
    # 规则对全部 20 位号码生效。
    out = []
    prov = no[2:4]
    if prov not in PROVINCE_CODES:
        out.append({"level": "警告", "field": "发票号码",
                    "message": f"20 位号码第 3-4 位省级区域码 {prov} 不在编码表内",
                    "suggest": "疑似号码识别错误"})
    ds = str(ctx["rec"].get("issue_date") or "")
    if len(ds) >= 4 and ds[:4].isdigit():
        if no[:2] != ds[2:4]:
            out.append({"level": "警告", "field": "发票号码",
                        "message": f"号码前两位 {no[:2]} 与开票年份 {ds[:4]} 不一致",
                        "suggest": "号码或开票日期识别有误（数电号码前两位=年份后两位）"})
    return out


@rule("R15", "警告", "发票代码结构异常")
def r15(ctx):
    code = str(ctx["rec"].get("invoice_code") or "")
    if len(code) != 12 or not code.isdigit():
        return []
    out = []
    if code[0] != "0":
        out.append({"level": "警告", "field": "发票代码",
                    "message": f"12 位发票代码首位应为 0，当前为 {code[0]}",
                    "suggest": "疑似代码识别错误"})
    if code[10:] not in ("11", "13"):
        out.append({"level": "提示", "field": "发票代码",
                    "message": f"发票代码第 11-12 位票种码为 {code[10:]}（常见 11/13）",
                    "suggest": "确认票种码是否为该地区特殊码制"})
    return out


@rule("R16", "警告", "购销方名称高度相似但税号不同")
def r16(ctx):
    bn = (ctx["rec"].get("buyer_name") or "").strip()
    sn = (ctx["rec"].get("seller_name") or "").strip()
    bt = str(ctx["rec"].get("buyer_tax") or "").strip().upper()
    st = str(ctx["rec"].get("seller_tax") or "").strip().upper()
    if not (bn and sn and bt and st) or bn == sn:
        return []
    sim = _similarity(bn, sn)
    if sim >= float(ctx["cfg"]["vendor_similarity"]) and bt != st:
        return [{"level": "警告", "field": "购销双方",
                 "message": f"购销双方名称相似度 {sim:.0%}（{bn} / {sn}）但税号不同",
                 "suggest": "核对是否名称识别缺字，或确为关联方"}]
    return []


@rule("R17", "警告", "大写金额与小写金额不一致")
def r17(ctx):
    """
    票面「价税合计（大写）」与「（小写）」必须相等。

    这是票据里信息量最大的一条互相印证关系：大小写同时对上的概率极高，而一旦不一致
    几乎必然是识别错误（OCR 认错字、字体映射错位、跨页丢字）。单字段校验永远抓不到
    这类错误 —— 它既不是缺字段，也不违反任何单字段合法性。
    """
    cn = ctx["rec"].get("_cn_amount")
    total = ctx["rec"].get("total")
    if cn is None or not isinstance(total, (int, float)):
        return []
    if abs(float(cn) - float(total)) > float(ctx["cfg"]["balance_tolerance"]):
        return [{"level": "警告", "field": "价税合计",
                 "message": f"票面大写金额 {float(cn):,.2f} 与小写价税合计 "
                            f"{float(total):,.2f} 不一致",
                 "suggest": "大小写金额必须相等，请人工核对（多为识别错误）"}]
    return []


def _similarity(a: str, b: str) -> float:
    """字符级相似度（不需要第三方库）。"""
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    sa, sb = set(a), set(b)
    jac = len(sa & sb) / max(len(sa | sb), 1)
    import difflib

    return max(jac, difflib.SequenceMatcher(None, a, b).ratio())


def validate(rec: dict, cfg: dict, today: date | None = None) -> list:
    """按注册表逐条跑规则；返回分级问题清单。"""
    today = today or date.today()
    disabled = set(cfg.get("rules_disabled") or [])
    ctx = {"rec": rec, "cfg": cfg, "today": today}
    issues = []
    for r in RULES:
        if r["id"] in disabled:
            continue
        try:
            issues.extend(r["fn"](ctx) or [])
        except Exception as exc:                            # noqa: BLE001
            issues.append({"level": "警告", "field": f"规则{r['id']}",
                           "message": f"规则执行异常：{exc}",
                           "suggest": "反馈给流程维护者"})
    return issues


# 规则给出的「涉及字段」是中文标签，置信度表用的却是英文 key。
# 旧实现直接做字符串互含匹配（`field in k or k in field`），中文标签与英文 key
# **永远匹配不上** —— 于是一段"严重问题 → 字段打折"的逻辑成了死代码。
# 实测佐证：一份勾稽不平（严重）的票，其 amount/tax/total 置信度仍是
# 0.9 / 0.9 / 0.93，一点没被压低。下面这张表把两者对上。
FIELD_ALIAS = {
    "发票号码": ["invoice_no"], "发票代码": ["invoice_code"],
    "开票日期": ["issue_date"], "校验码": ["check_code"],
    "购买方名称": ["buyer_name"], "销售方名称": ["seller_name"],
    "购买方税号": ["buyer_tax"], "销售方税号": ["seller_tax"],
    "税号": ["buyer_tax", "seller_tax"],
    "价税合计": ["total"], "金额(不含税)": ["amount"], "税额": ["tax"],
    "税率": ["tax_rate"], "税率×金额": ["amount", "tax", "tax_rate"],
    "金额勾稽": ["amount", "tax", "total"],
    "购销双方": ["buyer_name", "seller_name"],
}


def _penalty_targets(field: str, conf: dict) -> list:
    """把规则的中文字段标签翻译成置信度表里的 key；未知标签退回子串匹配。"""
    keys = FIELD_ALIAS.get(str(field or ""))
    if keys:
        return [k for k in keys if k in conf]
    return [k for k in conf if field and (field in k or k in field)]


def adjust_confidence(rec: dict, issues: list, cfg: dict) -> tuple:
    """
    用"校验结论"反过来修字段置信度 —— 交叉校验通过会加分，冲突会减分。
    这是把"校验"和"识别"打通的关键：v1 里两者互不相干。
    """
    conf = dict(rec.get("_confidence_map") or {})
    ev = dict(rec.get("_evidence") or {})
    notes = []

    # 勾稽通过 → 三个金额字段互相印证，加分
    if not any(i["field"] in ("金额勾稽", "税率×金额") for i in issues):
        if all(isinstance(rec.get(k), (int, float)) for k in ("amount", "tax", "total")):
            for k in ("amount", "tax", "total"):
                conf[k] = min(0.99, conf.get(k, 0.5) + 0.05)
            notes.append("金额三值勾稽通过 +0.05")

    # 校验位通过 → 税号加分
    for k in ("buyer_tax", "seller_tax"):
        code = str(rec.get(k) or "").strip().upper()
        if code and uscc_check(code) is True:
            conf[k] = min(0.99, conf.get(k, 0.6) + 0.07)
    if any(uscc_check(str(rec.get(k) or "").strip().upper()) is True
           for k in ("buyer_tax", "seller_tax")):
        notes.append("信用代码校验位通过 +0.07")

    # 号码结构对得上 → 号码与日期加分
    if not any(i["field"] == "发票号码" and "结构" in i["message"] for i in issues):
        if rec.get("invoice_no") and rec.get("issue_date"):
            conf["invoice_no"] = min(0.99, conf.get("invoice_no", 0.6) + 0.03)

    # 有严重问题 → 相关字段整体打折（说明"读到的值"和"规则"打架了）
    for i in issues:
        if i["level"] != "严重":
            continue
        for k in _penalty_targets(i["field"], conf):
            conf[k] = round(conf[k] * 0.6, 3)
            notes.append(f"因「{i['field']}」严重问题，{k} 置信度 ×0.6")

    rec["_confidence_map"] = conf
    rec["_evidence"] = ev
    return conf, notes


CRITICAL_FIELDS = ("invoice_no", "issue_date", "total")
CONF_WEIGHTS = {
    "invoice_no": 3.0, "issue_date": 2.0, "total": 3.0,
    "amount": 2.0, "tax": 2.0, "seller_name": 2.0, "buyer_name": 1.5,
    "seller_tax": 1.5, "buyer_tax": 1.0, "invoice_type": 0.5,
    "invoice_code": 1.0, "tax_rate": 1.0, "item_name": 0.5, "drawer": 0.3,
}


def overall_confidence(rec: dict) -> float:
    """记录级加权置信度；字段缺失按其权重计 0，所以"缺字段"会被自然惩罚。"""
    conf = rec.get("_confidence_map") or {}
    num = den = 0.0
    for k, w in CONF_WEIGHTS.items():
        num += conf.get(k, 0.0) * w
        den += w
    return round(num / den, 3) if den else 0.0


def _records_score(records: list) -> float:
    """
    给一组解析结果打分，用于"多视图择优"：取各记录加权置信度的平均值。
    关键字段取不到时分数会明显偏低（缺失字段按其权重计 0），
    所以它能区分"解析成功"和"文本拿到了但字段没抽出来"。
    """
    if not records:
        return -1.0
    return round(sum(overall_confidence(r) for r in records) / len(records), 4)


def decide_status(issues: list, conf_overall: float, rec: dict, cfg: dict) -> str:
    if any(i["level"] == "严重" for i in issues):
        return ST_ERROR
    conf = rec.get("_confidence_map") or {}
    low_critical = [k for k in CRITICAL_FIELDS
                    if rec.get(k) in ("", None)
                    or conf.get(k, 0.0) < float(cfg["min_critical_confidence"])]
    if conf_overall < float(cfg["min_confidence"]) or low_critical:
        return ST_REVIEW
    if any(i["level"] in ("警告", "提示") for i in issues):
        return ST_WARN
    return ST_OK


def summarize_issues(issues: list, limit: int = 6) -> str:
    if not issues:
        return ""
    order = {"严重": 0, "警告": 1, "重复": 2, "提示": 3}
    srt = sorted(issues, key=lambda i: order.get(i["level"], 9))
    parts = [f"[{i['level']}] {i['field']}: {i['message']}" for i in srt[:limit]]
    if len(srt) > limit:
        parts.append(f"……等共 {len(srt)} 条，详见「异常记录」")
    return "； ".join(parts)


# ==========================================================================
# 6. LLM 兜底（默认关闭；只补低置信度字段，且必须过同一套校验）
# ==========================================================================
def llm_extract(text: str, cfg: dict) -> dict:
    lc = cfg.get("llm_fallback") or {}
    if not lc.get("enabled"):
        return {}
    key = os.environ.get(lc.get("api_key_env", "INVOICE_LLM_API_KEY"), "").strip()
    if not key:
        return {}
    import urllib.request

    prompt = (
        "你是发票字段抽取器。只输出 JSON，不要解释。键：\n"
        'invoice_no, invoice_code, issue_date(YYYY-MM-DD), buyer_name, buyer_tax,'
        ' seller_name, seller_tax, amount, tax, total, tax_rate, item_name\n'
        "缺失写 null。数值不带千分位。\n\n票面文本：\n"
        + text[:int(lc.get("max_text_chars", 6000))]
    )
    body = json.dumps({
        "model": lc.get("model", "gpt-4o-mini"),
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0,
        "response_format": {"type": "json_object"},
    }).encode("utf-8")
    req = urllib.request.Request(
        lc["endpoint"], data=body,
        headers={"Content-Type": "application/json",
                 "Authorization": f"Bearer {key}"})
    try:
        with urllib.request.urlopen(req, timeout=int(lc.get("timeout", 30))) as resp:
            data = json.loads(resp.read().decode("utf-8", "replace"))
        content = data["choices"][0]["message"]["content"]
        return json.loads(content)
    except Exception as exc:                                # noqa: BLE001
        return {"_error": str(exc)}


# ==========================================================================
# 7. 台账层（Excel 可选 / CSV / JSON / 审计 JSONL）
# ==========================================================================
def file_fingerprint(path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()[:16]


_LONG_DIGITS = re.compile(r"^\d{8,}$")


def csv_text_cell(v):
    """长数字串包成 Excel 文本公式，避免丢前导 0 / 变科学计数法。"""
    if isinstance(v, str) and _LONG_DIGITS.match(v):
        return f'="{v}"'
    return "" if v is None else v


def _csv_escape(v) -> str:
    s = "" if v is None else str(v)
    if s.startswith('="') and s.endswith('"') and len(s) > 3:
        return s                                        # 已是文本公式，别再转义
    if any(ch in s for ch in ',"\r\n'):
        return '"' + s.replace('"', '""') + '"'
    return s


def write_csv(path, header, rows, encoding="utf-8-sig"):
    with open(path, "w", newline="", encoding=encoding) as fh:
        fh.write(",".join(_csv_escape(h) for h in header) + "\r\n")
        for row in rows:
            fh.write(",".join(_csv_escape(v) for v in row) + "\r\n")


def dedupe_key(rec: dict) -> str:
    code = str(rec.get("invoice_code") or "").strip()
    no = str(rec.get("invoice_no") or "").strip()
    if not no:
        return ""
    return f"{code}-{no}" if code else no


class Ledger:
    """Excel 台账 = 唯一数据源；openpyxl 缺失时自动降级为纯 CSV/JSON。"""

    CLR_HEADER_BG = "1F4E79"
    CLR_HEADER_FG = "FFFFFF"

    IGNORE_SHEETS = {"异常记录", "汇总统计", "待复核"}

    def __init__(self, cfg: dict):
        self.cfg = cfg
        self.out_dir = Path(cfg["output_dir"])
        self.out_dir.mkdir(parents=True, exist_ok=True)
        self.ledger_path = self.out_dir / cfg["ledger_name"]
        self.rows = []
        self.issues = []
        self._issue_keys = set()
        self.by_key = {}
        self.by_hash = {}
        self.audit = []
        self.stats = defaultdict(int)
        self.stats.update({"sum_amount": 0.0, "sum_tax": 0.0, "sum_total": 0.0})
        self._excel_ok = _have("openpyxl")
        self._load()

    # ---------------- 读回已有台账 ----------------
    def _load(self):
        """
        恢复上一次的状态。两条路，缺一不可：

          1. `发票台账.xlsx`（装了 openpyxl 且文件存在）—— 这是「Excel 即唯一数据源」的正式路径，
             人工在 Excel 里改过的值会被读回并保留。
          2. `台账索引.json`（内部状态文件）—— **零依赖路径的生命线**。
             没装 openpyxl 时读不了 xlsx，若不读这个文件，每次运行都会从空台账开始，
             **上一次的结果会被整体覆盖**（实测踩到过，属于数据丢失级缺陷）。
        """
        loaded = False
        if self._excel_ok and self.ledger_path.exists():
            loaded = self._load_from_excel()
        if not loaded:
            self._load_from_state()
        self._recalc()

    def _load_from_excel(self) -> bool:
        from openpyxl import load_workbook

        try:
            wb = load_workbook(self.ledger_path)
        except Exception:                                  # noqa: BLE001
            return False
        got = False
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            heads = [c.value for c in ws[1]]
            if sheet_name not in self.IGNORE_SHEETS:
                idx = {h: i for i, h in enumerate(heads) if h}
                for r in ws.iter_rows(min_row=2, values_only=True):
                    if r is None or all(v in (None, "") for v in r):
                        continue
                    rec = {k: (r[idx[h]] if idx.get(h) is not None and idx[h] < len(r)
                               else "") for k, h, _t in FIELDS}
                    try:
                        rec["seq"] = int(rec.get("seq"))
                    except (TypeError, ValueError):
                        rec["seq"] = len(self.rows) + 1
                    self._append_row(rec)
                    got = True
            elif sheet_name == "异常记录":
                for r in ws.iter_rows(min_row=2, values_only=True):
                    if r is None or all(v in (None, "") for v in r):
                        continue
                    d = dict(zip(heads, r))
                    self.issues.append(d)
                    self._issue_keys.add(self._issue_key(d))
        return got

    def _load_from_state(self):
        p = self.out_dir / self.cfg["state_name"]
        if not p.exists():
            return
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
        except Exception:                                  # noqa: BLE001
            return
        for rec in (data.get("rows") or []):
            if isinstance(rec, dict):
                rec.setdefault("_confidence_map", {})
                rec.setdefault("_evidence", {})
                self._append_row(rec)
        for d in (data.get("issues") or []):
            if isinstance(d, dict):
                self.issues.append(d)
                self._issue_keys.add(self._issue_key(d))
        self._recalc()

    def _append_row(self, rec: dict):
        self.rows.append(rec)
        k = dedupe_key(rec)
        if k and k not in self.by_key:
            self.by_key[k] = rec
        h = str(rec.get("file_hash") or "").strip()
        if h and h not in self.by_hash:
            self.by_hash[h] = rec

    @staticmethod
    def _issue_key(d: dict) -> tuple:
        return (str(d.get("源文件名") or ""), str(d.get("发票号码") or ""),
                str(d.get("问题级别") or ""), str(d.get("涉及字段") or ""),
                str(d.get("问题描述") or ""))

    def _recalc(self):
        counted = [r for r in self.rows
                   if str(r.get("status") or "") not in (ST_DUP, ST_TODO)]
        self.stats["sum_amount"] = round(sum(float(r.get("amount") or 0)
                                             for r in counted), 2)
        self.stats["sum_tax"] = round(sum(float(r.get("tax") or 0)
                                          for r in counted), 2)
        self.stats["sum_total"] = round(sum(float(r.get("total") or 0)
                                            for r in counted), 2)
        self.stats["ok"] = sum(1 for r in self.rows if r.get("status") == ST_OK)
        self.stats["warn"] = sum(1 for r in self.rows if r.get("status") == ST_WARN)
        self.stats["error"] = sum(1 for r in self.rows if r.get("status") == ST_ERROR)
        self.stats["review"] = sum(1 for r in self.rows if r.get("status") == ST_REVIEW)
        self.stats["ocr_todo"] = sum(1 for r in self.rows if r.get("status") == ST_TODO)
        self.stats["duplicate"] = sum(1 for r in self.issues
                                      if str(r.get("问题级别")) == "重复")

    # ---------------- 写入 ----------------
    def add_issue(self, file_name, invoice_no, level, field, message, suggest=""):
        key = (str(file_name or ""), str(invoice_no or ""), level, field, message)
        if key in self._issue_keys:
            return
        self._issue_keys.add(key)
        self.issues.append({
            "序号": len(self.issues) + 1, "源文件名": file_name,
            "发票号码": invoice_no or "—", "问题级别": level, "涉及字段": field,
            "问题描述": message, "建议处理": suggest or "人工核对",
            "发生时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})

    def add_record(self, rec: dict):
        rec["seq"] = len(self.rows) + 1
        self.rows.append(rec)
        k = dedupe_key(rec)
        if k and k not in self.by_key:
            self.by_key[k] = rec
        h = str(rec.get("file_hash") or "")
        if h and h not in self.by_hash:
            self.by_hash[h] = rec

    def find_duplicate(self, rec: dict, file_hash: str):
        k = dedupe_key(rec)
        if k and k in self.by_key:
            return "发票级重复", self.by_key[k]
        if file_hash and file_hash in self.by_hash:
            return "文件级重复", self.by_hash[file_hash]
        return "", None

    def remove_row(self, rec: dict):
        """从台账移除一行（供 --force 强制重算某个已入账文件时用）。"""
        try:
            self.rows.remove(rec)
        except ValueError:
            return
        k = dedupe_key(rec)
        if k and self.by_key.get(k) is rec:
            self.by_key.pop(k, None)
        h = str(rec.get("file_hash") or "")
        if h and self.by_hash.get(h) is rec:
            self.by_hash.pop(h, None)

    def purge_issues(self, file_name, fields=("文本提取", "流程", "字段解析")):
        keep = [d for d in self.issues
                if not (str(d.get("源文件名") or "") == file_name
                        and str(d.get("涉及字段") or "") in fields)]
        if len(keep) != len(self.issues):
            self.issues = keep
            self._issue_keys = {self._issue_key(d) for d in self.issues}
            for i, d in enumerate(self.issues, 1):
                d["序号"] = i

    # ---------------- 落盘 ----------------
    def save(self):
        paths = {}
        # Excel
        if self._excel_ok:
            from openpyxl import Workbook
            from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
            from openpyxl.utils import get_column_letter

            thin = Side(style="thin", color="D0D7E5")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)
            wb = Workbook()

            ws = wb.active
            ws.title = "发票明细"
            for j, (key, header, kind) in enumerate(FIELDS, 1):
                c = ws.cell(row=1, column=j, value=header)
                c.font = Font(bold=True, color=self.CLR_HEADER_FG, size=11)
                c.fill = PatternFill("solid", fgColor=self.CLR_HEADER_BG)
                c.alignment = Alignment(horizontal="center", vertical="center",
                                        wrap_text=True)
                c.border = border
                ws.column_dimensions[get_column_letter(j)].width = (
                    22 if kind == "money" else
                    32 if header in ("源文件名", "购买方名称", "销售方名称",
                                     "主要项目/货物名称", "异常/提示说明") else 14)
            ws.row_dimensions[1].height = 30

            for i, rec in enumerate(self.rows, start=2):
                status = str(rec.get("status") or "")
                fill = None
                if status == ST_ERROR:
                    fill = PatternFill("solid", fgColor="FCE4E4")
                elif status in (ST_REVIEW, ST_TODO):
                    fill = PatternFill("solid", fgColor="EDE7F6")
                elif status == ST_DUP:
                    fill = PatternFill("solid", fgColor="FFF4D6")
                for j, (key, _h, kind) in enumerate(FIELDS, 1):
                    val = rec.get(key, "")
                    if kind == "money":
                        val = None if val in ("", None) else val
                    c = ws.cell(row=i, column=j, value=val)
                    c.border = border
                    if fill:
                        c.fill = fill
                    if kind == "money":
                        c.number_format = "#,##0.00"
                        c.alignment = Alignment(horizontal="right",
                                                vertical="center")
                    elif key == "confidence":
                        c.alignment = Alignment(horizontal="center")
                        if val not in ("", None):
                            try:
                                fv = float(str(val).rstrip("%"))
                                c.font = Font(color=("1E7145" if fv >= 90 else
                                                     "BF8F00" if fv >= 75 else "C00000"))
                            except ValueError:
                                pass
                    elif key in ("status", "invoice_type", "tax_rate", "issue_date"):
                        c.alignment = Alignment(horizontal="center", vertical="center",
                                                wrap_text=True)
                    else:
                        c.alignment = Alignment(horizontal="left", vertical="center")
            ws.freeze_panes = "C2"
            if self.rows:
                ws.auto_filter.ref = f"A1:{get_column_letter(len(FIELDS))}{len(self.rows)+1}"

            # 异常记录
            ws2 = wb.create_sheet("异常记录")
            heads = ["序号", "源文件名", "发票号码", "问题级别", "涉及字段",
                     "问题描述", "建议处理", "发生时间"]
            widths = [6, 30, 22, 10, 14, 54, 32, 20]
            for j, (h, w) in enumerate(zip(heads, widths), 1):
                c = ws2.cell(row=1, column=j, value=h)
                c.font = Font(bold=True, color=self.CLR_HEADER_FG)
                c.fill = PatternFill("solid", fgColor=self.CLR_HEADER_BG)
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = border
                ws2.column_dimensions[get_column_letter(j)].width = w
            for i, d in enumerate(self.issues, start=2):
                for j, h in enumerate(heads, 1):
                    c = ws2.cell(row=i, column=j, value=d.get(h, ""))
                    c.border = border
                    c.alignment = Alignment(vertical="center",
                                            wrap_text=h in ("问题描述", "建议处理"))
            ws2.freeze_panes = "A2"

            # 汇总统计
            ws3 = wb.create_sheet("汇总统计")
            agg = defaultdict(lambda: [0, 0.0, 0.0, 0.0])
            for r in self.rows:
                k = str(r.get("seller_name") or "（未识别）")
                agg[k][0] += 1
                agg[k][1] += float(r.get("amount") or 0)
                agg[k][2] += float(r.get("tax") or 0)
                agg[k][3] += float(r.get("total") or 0)
            ws3.column_dimensions["A"].width = 26
            for col in "BCDE":
                ws3.column_dimensions[col].width = 18
            summary_rows = [
                ("统计项", "数值", "说明"),
                ("台账发票行数", len(self.rows), "一张发票一行"),
                ("正常", self.stats["ok"], "校验全部通过"),
                ("存在提示", self.stats["warn"], "有警告或提示"),
                ("异常", self.stats["error"], "存在严重问题"),
                ("待复核", self.stats["review"], "置信度不足，需人工确认"),
                ("待人工录入", self.stats["ocr_todo"], "图片/扫描件未识别"),
                ("重复拦截", self.stats["duplicate"], "已阻止重复入账"),
                ("", "", ""),
                ("合计金额(不含税)", self.stats["sum_amount"], ""),
                ("合计税额", self.stats["sum_tax"], ""),
                ("合计价税合计", self.stats["sum_total"], ""),
                ("", "", ""),
                ("生成时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S"), ""),
            ]
            for i, (a, b, cc) in enumerate(summary_rows, start=1):
                ws3.cell(row=i, column=1, value=a).font = Font(
                    bold=(i == 1 or str(a).startswith("合计")))
                cell = ws3.cell(row=i, column=2, value=b)
                if isinstance(b, float):
                    cell.number_format = "#,##0.00"
                ws3.cell(row=i, column=3, value=cc).font = Font(size=9, color="808080")

            try:
                wb.save(self.ledger_path)
                paths["ledger"] = self.ledger_path
            except PermissionError:
                raise

        # CSV / JSON
        vis = [(k, h, t) for k, h, t in FIELDS if k != "file_hash"]
        json_path = self.out_dir / self.cfg["json_name"]
        payload = []
        for r in self.rows:
            item = {h: r.get(k, "") for k, h, _t in vis}
            item["识别证据"] = json.dumps(r.get("_evidence") or {},
                                          ensure_ascii=False)
            payload.append(item)
        json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2),
                             encoding="utf-8")
        paths["json"] = json_path

        csv_path = self.out_dir / self.cfg["csv_name"]
        write_csv(csv_path, [h for _k, h, _t in vis],
                  [[csv_text_cell(r.get(k, "")) for k, _h, _t in vis]
                   for r in self.rows])
        paths["csv"] = csv_path

        issue_path = self.out_dir / self.cfg["issue_csv_name"]
        if self.issues:
            heads = list(self.issues[0].keys())
            write_csv(issue_path, heads,
                      [[csv_text_cell(d.get(h, "")) for h in heads]
                       for d in self.issues])
            paths["issue_csv"] = issue_path

        # 待复核清单（只挑需要人看的，人工工作量最小化）
        review = [r for r in self.rows if r.get("status") in (ST_REVIEW, ST_TODO)]
        review_path = self.out_dir / self.cfg["review_csv_name"]
        cols = [("file_name", "源文件名"), ("invoice_no", "发票号码"),
                ("issue_date", "开票日期"), ("seller_name", "销售方名称"),
                ("total", "价税合计"), ("confidence", "识别置信度"),
                ("status", "处理状态"), ("issue", "异常/提示说明")]
        write_csv(review_path,
                  [h for _k, h in cols],
                  [[csv_text_cell(r.get(k, "")) for k, _h in cols] for r in review])
        paths["review_csv"] = review_path

        # 内部状态：下批运行靠它恢复台账（零依赖路径的生命线）
        state_path = self.out_dir / self.cfg["state_name"]
        state_path.write_text(json.dumps(
            {"version": __version__,
             "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
             "rows": self.rows, "issues": self.issues},
            ensure_ascii=False, indent=1), encoding="utf-8")
        paths["state"] = state_path

        # 审计轨迹
        audit_path = self.out_dir / self.cfg["audit_name"]
        with open(audit_path, "w", encoding="utf-8") as fh:
            for line in self.audit:
                fh.write(json.dumps(line, ensure_ascii=False) + "\n")
        paths["audit"] = audit_path
        return paths


# ==========================================================================
# 8. HTML 报告
# ==========================================================================
def _e(v) -> str:
    return _html.escape(str(v if v is not None else ""))


def _money(v) -> str:
    if v in ("", None):
        return "—"
    try:
        return f"{float(v):,.2f}"
    except (TypeError, ValueError):
        return _e(v)


def build_report(ledger: Ledger, summary: dict) -> Path:
    out = Path(ledger.cfg["output_dir"]) / ledger.cfg["report_name"]
    s = ledger.stats

    cards = [
        ("台账发票行数", len(ledger.rows), "一张发票一行", ""),
        ("本次新增发票", s["new_invoices"], f"新增文件 {s['new_files']} 个", ""),
        ("异常", s["error"], "需人工核对", "err"),
        ("待复核", s["review"], "置信度不足", "rev"),
        ("存在提示", s["warn"], "建议复核", "warn"),
        ("重复拦截", s["duplicate"], "已阻止重复入账", "dup"),
        ("待人工录入", s["ocr_todo"], "图片/扫描件", "todo"),
        ("合计价税合计", f"¥{s['sum_total']:,.2f}", "", "money"),
    ]
    card_html = "".join(
        f'<div class="card {c[3]}"><div class="c-label">{_e(c[0])}</div>'
        f'<div class="c-value">{_e(c[1])}</div>'
        f'<div class="c-sub">{_e(c[2])}</div></div>' for c in cards)

    order = {"严重": 0, "警告": 1, "重复": 2, "提示": 3}
    issues = sorted(ledger.issues, key=lambda d: order.get(str(d.get("问题级别")), 9))
    if issues:
        rows_issue = "".join(
            '<tr><td>{i}</td><td class="mono">{f}</td><td class="mono">{n}</td>'
            '<td><span class="badge b-{k}">{lv}</span></td><td>{fd}</td>'
            '<td>{msg}</td><td class="dim">{sg}</td></tr>'.format(
                i=i, f=_e(d.get("源文件名")), n=_e(d.get("发票号码")),
                k={"严重": "err", "警告": "warn", "重复": "dup",
                   "提示": "hint"}.get(str(d.get("问题级别")), "hint"),
                lv=_e(d.get("问题级别")), fd=_e(d.get("涉及字段")),
                msg=_e(d.get("问题描述")), sg=_e(d.get("建议处理")))
            for i, d in enumerate(issues, 1))
    else:
        rows_issue = ('<tr><td colspan="7" class="empty">'
                      '本次没有发现问题，全部发票校验通过 ✓</td></tr>')

    show = [(k, h, t) for k, h, t in FIELDS
            if k not in ("issue", "file_hash", "seq")]
    thead = "".join(f"<th>{_e(h)}</th>" for _k, h, _t in show)
    body = []
    for r in ledger.rows:
        tds = []
        for k, _h, kind in show:
            v = r.get(k)
            if k == "status":
                cls = {"正常": "b-ok", "异常": "b-err", "待复核": "b-rev",
                       "待人工录入": "b-todo", "重复未入账": "b-dup"}.get(str(v), "b-warn")
                tds.append(f'<td class="ctr"><span class="badge {cls}">{_e(v)}</span></td>')
            elif k == "confidence":
                try:
                    fv = float(str(v).rstrip("%"))
                except (TypeError, ValueError):
                    fv = 0.0
                color = "#1E7145" if fv >= 90 else "#BF8F00" if fv >= 75 else "#C00000"
                tds.append(f'<td class="ctr" style="color:{color};font-weight:600">'
                           f'{_e(v)}</td>')
            elif kind == "money":
                tds.append(f'<td class="num">{_money(v)}</td>')
            else:
                cls = "mono" if k in ("invoice_no", "invoice_code", "buyer_tax",
                                      "seller_tax") else ""
                tds.append(f'<td class="{cls}">{_e(v if v not in ("", None) else "—")}</td>')
        row_cls = {"异常": "row-err", "待复核": "row-rev", "待人工录入": "row-todo",
                   "重复未入账": "row-dup"}.get(str(r.get("status")), "")
        body.append(f'<tr class="{row_cls}">{"".join(tds)}</tr>')
    if not body:
        body = [f'<tr><td colspan="{len(show)}" class="empty">台账中暂无记录</td></tr>']

    banner = ""
    if not summary.get("ocr_available"):
        banner = ('<div class="banner"><b>⚠ 未检测到 OCR 引擎</b>：电子版 PDF 已正常识别；'
                  '图片与扫描件 PDF 已登记为「待人工录入」。'
                  '<pre>' + _e(OCR_HINT) + '</pre></div>')

    doc = f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>发票处理报告 · {_e(summary.get('finished_at', datetime.now()).strftime('%Y-%m-%d %H:%M'))}</title>
<style>
 :root{{--bg:#f5f7fb;--panel:#fff;--line:#e3e8f0;--text:#1e2a3a;--dim:#7b8794;
   --blue:#1f4e79;--red:#c0392b;--amber:#b7791f;--green:#1e7145;--purple:#5b4b8a;}}
 *{{box-sizing:border-box}}
 body{{margin:0;background:var(--bg);color:var(--text);font-size:14px;line-height:1.6;
   font-family:"Microsoft YaHei","PingFang SC","Segoe UI",system-ui,sans-serif}}
 .wrap{{max-width:1560px;margin:0 auto;padding:26px 22px 60px}}
 header{{background:linear-gradient(135deg,#1f4e79,#2e6da4);color:#fff;border-radius:14px;
   padding:24px 28px;box-shadow:0 6px 22px rgba(31,78,121,.18)}}
 header h1{{margin:0 0 6px;font-size:23px}}
 header .meta{{opacity:.9;font-size:13px}}
 header .meta span{{margin-right:16px;white-space:nowrap}}
 .cards{{display:grid;grid-template-columns:repeat(auto-fit,minmax(158px,1fr));
   gap:13px;margin:20px 0 6px}}
 .card{{background:var(--panel);border:1px solid var(--line);border-radius:12px;
   padding:15px 17px;box-shadow:0 2px 8px rgba(16,24,40,.04)}}
 .card .c-label{{font-size:12px;color:var(--dim)}}
 .card .c-value{{font-size:23px;font-weight:700;margin:3px 0 2px;color:var(--blue)}}
 .card .c-sub{{font-size:11px;color:var(--dim)}}
 .card.err .c-value{{color:var(--red)}} .card.warn .c-value{{color:var(--amber)}}
 .card.rev .c-value{{color:var(--purple)}}
 .card.dup .c-value,.card.todo .c-value{{color:#8a6d3b}}
 .card.money .c-value{{font-size:19px}}
 .banner{{background:#fff7e6;border:1px solid #f0d9a8;border-left:4px solid var(--amber);
   border-radius:10px;padding:13px 17px;margin:16px 0;font-size:13px}}
 .banner pre{{margin:7px 0 0;white-space:pre-wrap;font-size:12px;color:#5c4a1f}}
 h2{{font-size:16px;margin:28px 0 11px;padding-left:10px;border-left:4px solid var(--blue)}}
 .panel{{background:var(--panel);border:1px solid var(--line);border-radius:12px;
   overflow:auto;max-height:640px;box-shadow:0 2px 8px rgba(16,24,40,.04)}}
 table{{border-collapse:separate;border-spacing:0;width:100%;font-size:12.5px}}
 th{{position:sticky;top:0;background:#eef3f9;color:var(--blue);font-weight:600;
   text-align:left;padding:9px;border-bottom:1px solid var(--line);white-space:nowrap;z-index:2}}
 td{{padding:7px 9px;border-bottom:1px solid #f0f3f8;vertical-align:top}}
 tbody tr:hover{{background:#f8fbff}}
 tr.row-err{{background:#fdecea}} tr.row-rev{{background:#f0ebfa}}
 tr.row-dup{{background:#fff7e6}} tr.row-todo{{background:#f2eefb}}
 td.num{{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}}
 td.ctr{{text-align:center;white-space:nowrap}}
 td.mono{{font-family:Consolas,monospace;font-size:12px;white-space:nowrap}}
 td.dim{{color:var(--dim);font-size:12px}}
 td.empty{{text-align:center;color:var(--green);padding:24px;font-weight:600}}
 .badge{{display:inline-block;padding:2px 9px;border-radius:20px;font-size:11.5px;font-weight:600}}
 .b-ok{{background:#e8f5ee;color:var(--green)}} .b-err{{background:#f8d7d3;color:var(--red)}}
 .b-warn{{background:#fdeec7;color:var(--amber)}} .b-dup{{background:#fdeec7;color:#8a6d3b}}
 .b-todo{{background:#e6e0f5;color:#5b4b8a}} .b-rev{{background:#ece5f8;color:var(--purple)}}
 .b-hint{{background:#e9eef5;color:#5a6b7d}}
 footer{{margin-top:30px;color:var(--dim);font-size:12px;line-height:1.9}}
 code{{background:#eef2f7;padding:1px 6px;border-radius:5px;font-family:Consolas,monospace;font-size:12px}}
</style></head><body><div class="wrap">
<header><h1>发票自动化处理报告 · v2</h1>
<div class="meta">
 <span>生成时间：{_e(summary.get('finished_at', datetime.now()).strftime('%Y-%m-%d %H:%M:%S'))}</span>
 <span>批次耗时：{summary.get('elapsed', 0):.2f}s</span>
 <span>输入目录：{_e(summary.get('input_dir', ''))}</span>
 <span>扫描文件：{summary.get('files', 0)} 个</span>
 <span>文本提取器：{_e(summary.get('extractor_mix', ''))}</span>
</div></header>
{banner}
<div class="cards">{card_html}</div>
<h2>异常与提示清单（{len(issues)} 条）</h2>
<div class="panel"><table><thead><tr><th>#</th><th>源文件名</th><th>发票号码</th>
<th>级别</th><th>涉及字段</th><th>问题描述</th><th>建议处理</th></tr></thead>
<tbody>{rows_issue}</tbody></table></div>
<h2>发票明细台账（{len(ledger.rows)} 行 · 一张发票一行）</h2>
<div class="panel"><table><thead><tr>{thead}</tr></thead>
<tbody>{''.join(body)}</tbody></table></div>
<footer>
台账：<code>{_e(str(ledger.ledger_path))}</code>　
待复核清单：<code>{_e(str(Path(ledger.cfg['output_dir']) / ledger.cfg['review_csv_name']))}</code>　
审计轨迹：<code>{_e(str(Path(ledger.cfg['output_dir']) / ledger.cfg['audit_name']))}</code><br>
说明：合计金额只统计「真正入账」的行；「重复未入账」「待人工录入」不计入。<br>
「待复核」= 加权置信度低于 {ledger.cfg['min_confidence']} 或关键字段置信度低于
{ledger.cfg['min_critical_confidence']}，建议优先人工确认这些行。
</footer></div></body></html>"""
    out.write_text(doc, encoding="utf-8")
    return out


# ==========================================================================
# 9. 编排
# ==========================================================================
def discover_files(input_dir, recursive=True) -> list:
    root = Path(input_dir)
    if not root.exists():
        return []
    it = root.rglob("*") if recursive else root.glob("*")
    files = [p for p in it if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS
             and not p.name.startswith("~$")]
    return sorted(files, key=lambda p: (p.stat().st_mtime, p.name))


def _empty_record(name: str, fp: str, note: str) -> dict:
    rec = {k: "" for k, _h, _t in FIELDS}
    rec.update({"file_name": name, "file_hash": fp, "status": ST_TODO,
                "issue": f"[严重] 文本提取: {note}",
                "processed_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "confidence": "0.0%", "extractor": "none",
                "_confidence_map": {}, "_evidence": {}})
    return rec


def process_file(path: Path, ledger: Ledger, cfg: dict, templates: list,
                 log=print, verbose=True) -> dict:
    delta = defaultdict(int)
    fp = file_fingerprint(path)
    name = path.name

    # ---- 文件指纹预检（幂等的关键） ----
    old = ledger.by_hash.get(fp)
    if old is not None and cfg.get("force"):
        # --force：把已入账的那一行摘掉再重算。
        # 典型用途：① 装好 OCR 后重算此前的「待人工录入」文件；
        #           ② 改了版式模板后重算受影响的票。
        ledger.remove_row(old)
        old = None
    if old is not None:
        old_name = str(old.get("file_name") or "").split(" [第")[0]
        if old_name == name:
            delta["skipped"] += 1
            log(f"  [已处理] {name} —— 跳过")
            return delta
        delta["duplicate"] += 1
        ledger.add_issue(name, old.get("invoice_no", ""), "重复", "文件指纹",
                         f"内容与台账中「{old_name}」完全一致（指纹 {fp}），已阻止重复入账",
                         "确认为重复提交时忽略；如需重解请删除台账对应行")
        log(f"  [重复提交] {name} —— 与「{old_name}」内容相同，已拦截")
        return delta

    t0 = time.time()
    # ---- 文本提取 ----
    try:
        info = extract_text(path, cfg)
    except Exception as exc:                                # noqa: BLE001
        msg = str(exc)
        rec = _empty_record(name, fp, msg.splitlines()[0])
        ledger.add_record(rec)
        ledger.add_issue(name, "", "严重", "文本提取", msg,
                         "改为人工录入，或在装好 OCR 后重跑")
        ledger.audit.append({"file": name, "hash": fp, "extractor": "none",
                             "quality": 0.0, "records": 0, "error": msg.splitlines()[0],
                             "elapsed": round(time.time() - t0, 3)})
        delta["ocr_todo"] += 1
        log(f"  [待人工] {name} —— {msg.splitlines()[0][:60]}")
        return delta

    delta["new_files"] += 1

    # ---- 字段解析：多视图择优 ----
    # 同一份 PDF 可能有多个文本视图（PyMuPDF / 标准库 / OCR），不同视图对不同版式
    # 各有胜负：PyMuPDF 按文本区块排序，遇到"标签在左列、取值在右列"的版式会把两者
    # 拆散；标准库解析器按内容流绘制顺序，标签紧邻取值。所以这里把每个视图都解析
    # 一遍，用"记录级加权置信度"挑最好的一组 —— 与"文本质量闸门"是同一套比质量取优的思路。
    views_all = [(info["source"], info["text"])] + list(info.get("alts") or [])
    candidates = []                      # [(来源, 记录列表, 得分)]
    for src_i, txt_i in views_all:
        try:
            recs_i = parse_invoice(txt_i, name, templates)
        except Exception as exc:                            # noqa: BLE001
            ledger.add_issue(name, "", "严重", "字段解析",
                             f"{src_i} 解析异常：{exc}", "请把该文件反馈给流程维护者")
            log(f"  [异常] {name} 解析失败（{src_i}）：{exc}")
            if verbose:
                import traceback
                traceback.print_exc()
            continue
        candidates.append((src_i, recs_i, _records_score(recs_i)))

    if not candidates:
        return delta

    candidates.sort(key=lambda c: -c[2])
    src_note, records, best_score = candidates[0]
    if records and len(candidates) > 1 and best_score > candidates[-1][2] + 0.02:
        src_note = f"{src_note}(多视图择优，共 {len(candidates)} 个视图)"
        log(f"  [择优] {name} → 采用 {src_note.split('(')[0]}"
            f"（候选得分 {', '.join(f'{s}={sc:.2f}' for s, _r, sc in candidates)}）")

    if not records:
        ledger.add_issue(name, "", "严重", "字段解析", "未从文件中解析出任何发票记录",
                         "确认是否发票文件（非发票 PDF 会被忽略）")
        return delta

    # 同一文件含多张发票时，它们共享同一个文件指纹；
    # 第 1 张入账后必须停止用"文件指纹"去判重，否则第 2 张起会被误判为「文件级重复」。
    committed_in_this_file = False

    for rec in records:
        rec["file_hash"] = fp
        rec["extractor"] = src_note
        rec["processed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # ---- LLM 兜底（仅低置信度 + 显式开启）----
        llm_note = ""
        lc = cfg.get("llm_fallback") or {}
        if lc.get("enabled"):
            pre_conf = overall_confidence(rec)
            if pre_conf < float(cfg["min_confidence"]):
                got = llm_extract(info["text"], cfg)
                if got and not got.get("_error"):
                    merged = 0
                    for key in ("invoice_no", "invoice_code", "issue_date",
                                "buyer_name", "buyer_tax", "seller_name",
                                "seller_tax", "amount", "tax", "total",
                                "tax_rate", "item_name"):
                        cur = rec.get(key)
                        if cur in ("", None) and got.get(key) not in (None, ""):
                            rec[key] = got[key]
                            rec["_confidence_map"][key] = 0.7
                            rec["_evidence"][key] = "[LLM 兜底]"
                            merged += 1
                    llm_note = f"；LLM 兜底补全 {merged} 个字段"
                elif got.get("_error"):
                    llm_note = f"；LLM 兜底失败（{got['_error']}）"

        # ---- 校验（先校验后置信度调整，形成闭环）----
        issues = validate(rec, cfg)
        conf_map, conf_notes = adjust_confidence(rec, issues, cfg)
        conf_overall = overall_confidence(rec)

        rec["confidence"] = f"{conf_overall * 100:.1f}%"
        rec["status"] = decide_status(issues, conf_overall, rec, cfg)
        rec["issue"] = summarize_issues(issues) + llm_note

        for it in issues:
            ledger.add_issue(str(rec["file_name"]).split(" [第")[0],
                             rec.get("invoice_no", ""), it["level"],
                             it["field"], it["message"], it["suggest"])

        # ---- 重复校验 ----
        dup_kind, old_row = ledger.find_duplicate(
            rec, "" if committed_in_this_file else fp)
        if dup_kind:
            delta["duplicate"] += 1
            where = old_row.get("file_name", "台账已有行") if old_row else "台账已有行"
            ledger.add_issue(str(rec["file_name"]).split(" [第")[0],
                             rec.get("invoice_no", ""), "重复", "发票唯一性",
                             f"{dup_kind}：与台账「{where}」第 {old_row.get('seq')} 行重复，未入账",
                             "确认为重复报销时忽略；需保留可改 on_duplicate=mark")
            log(f"  [重复] {rec.get('invoice_no') or name} —— {dup_kind}，未入账")
            if cfg["on_duplicate"] == "mark":
                rec["status"] = ST_DUP
                ledger.add_record(rec)
                committed_in_this_file = True
                delta["new_invoices"] += 1
            continue

        ledger.add_record(rec)
        committed_in_this_file = True
        delta["new_invoices"] += 1
        flag = "" if rec["status"] == ST_OK else f"  ← {rec['status']}"
        log(f"  [入账] {rec.get('invoice_no') or '(无号码)'} "
            f"{str(rec.get('seller_name') or '')[:16]} "
            f"¥{rec.get('total') if rec.get('total') is not None else '?'} "
            f"置信 {rec['confidence']} ({src_note}){flag}")

        ledger.audit.append({
            "file": rec["file_name"], "hash": fp, "invoice_no": rec.get("invoice_no"),
            "extractor": src_note, "quality": round(info["quality"], 3),
            "pages": info["pages"], "confidence": conf_overall,
            "confidence_map": conf_map, "evidence": rec.get("_evidence"),
            "confidence_notes": conf_notes, "notes": info.get("notes", ""),
            "issues": [f"{i['level']}/{i['field']}: {i['message']}" for i in issues],
            "status": rec["status"], "elapsed": round(time.time() - t0, 3),
        })

    ledger.purge_issues(name)
    return delta


def run(cfg: dict, log=print, verbose=True) -> dict:
    t0 = datetime.now()
    templates = load_templates()
    input_dir = Path(cfg["input_dir"])
    input_dir.mkdir(parents=True, exist_ok=True)

    files = discover_files(input_dir, cfg["recursive"])
    log("=" * 72)
    log(f"发票识别工作流机器人 v{__version__}   {t0:%Y-%m-%d %H:%M:%S}")
    log(f"输入目录：{input_dir}")
    log(f"输出目录：{Path(cfg['output_dir'])}")
    log(f"待处理文件：{len(files)} 个")
    log(f"PDF 引擎：{'PyMuPDF' if (_have('pymupdf') or _have('fitz')) else '标准库(内置)'}"
        f"    Excel：{'openpyxl' if _have('openpyxl') else '不可用→仅 CSV/JSON'}"
        f"    OCR：{ocr_provider() or '未安装'}")
    if not ocr_provider():
        log("提示：未检测到 OCR 引擎，图片/扫描件将登记为『待人工录入』")
    log("-" * 72)

    ledger = Ledger(cfg)
    for p in files:
        log(f"处理：{p.name}")
        try:
            d = process_file(p, ledger, cfg, templates, log=log, verbose=verbose)
        except Exception as exc:                            # noqa: BLE001
            ledger.add_issue(p.name, "", "严重", "流程", f"未预期异常：{exc}",
                             "请保留文件并反馈")
            log(f"  [异常] {p.name}：{exc}")
            if verbose:
                import traceback
                traceback.print_exc()
            continue
        for k, v in d.items():
            ledger.stats[k] += v

    ledger._recalc()
    try:
        paths = ledger.save()
    except PermissionError:
        log("")
        log("!" * 72)
        log(f"[无法写入] {ledger.ledger_path}")
        log("该文件正被其他程序占用（最常见是 Excel 开着它）。")
        log("请先关闭 Excel / WPS 再重跑；已入账的文件会自动跳过。")
        log("!" * 72)
        return {"failed": True, "reason": "ledger_locked"}

    mix = defaultdict(int)
    for a in ledger.audit:
        mix[a.get("extractor", "?")] += 1

    summary = {
        "started_at": t0, "finished_at": datetime.now(),
        "elapsed": (datetime.now() - t0).total_seconds(),
        "input_dir": str(input_dir), "files": len(files),
        "extractor_mix": "、".join(f"{k}×{v}" for k, v in mix.items()),
        "paths": {k: str(v) for k, v in paths.items()},
        "ocr_available": ocr_provider() is not None,
    }
    build_report(ledger, summary)

    s = ledger.stats
    log("-" * 72)
    log(f"处理完成，用时 {summary['elapsed']:.2f}s")
    log(f"台账行数 {len(ledger.rows)}  |  正常 {s['ok']}  |  提示 {s['warn']}  |  "
        f"异常 {s['error']}  |  待复核 {s['review']}  |  重复拦截 {s['duplicate']}  |  "
        f"待人工 {s['ocr_todo']}")
    log(f"合计金额 {s['sum_amount']:,.2f}  |  合计税额 {s['sum_tax']:,.2f}  |  "
        f"价税合计 {s['sum_total']:,.2f}")
    for k, v in paths.items():
        log(f"  {k:12s} → {v}")
    if s["review"]:
        log(f"⚠ 有 {s['review']} 行置信度不足，已列入「{cfg['review_csv_name']}」，"
            f"建议优先人工确认")
    log("=" * 72)
    return {"failed": False, "stats": dict(s), "summary": summary}


def dry_run(cfg: dict, templates: list, log=print) -> int:
    log("=" * 72)
    log("[干跑模式] 只解析校验，不写任何文件")
    files = discover_files(cfg["input_dir"], cfg["recursive"])
    log(f"发现 {len(files)} 个文件")
    log("-" * 72)
    ok = bad = 0
    for p in files:
        try:
            info = extract_text(p, cfg)
            # 与正式流程一致：所有视图都解析一遍，按记录级置信度择优
            views_all = [(info["source"], info["text"])] + list(info.get("alts") or [])
            picked, view_note = None, ""
            for src_i, txt_i in views_all:
                recs_i = parse_invoice(txt_i, p.name, templates)
                sc_i = _records_score(recs_i)
                if picked is None or sc_i > picked[1]:
                    picked = (src_i, sc_i, recs_i)
            src_show, _sc, recs = picked
            if src_show != info["source"]:
                view_note = f"（多视图择优：{info['source']}→{src_show}）"
        except Exception as exc:                            # noqa: BLE001
            log(f"✗ {p.name}\n   提取失败：{exc}")
            bad += 1
            continue
        for i, r in enumerate(recs, 1):
            issues = validate(r, cfg)
            adjust_confidence(r, issues, cfg)
            conf = overall_confidence(r)
            st = decide_status(issues, conf, r, cfg)
            tag = f"（第 {i}/{len(recs)} 张）" if len(recs) > 1 else ""
            log(f"✓ {p.name}{tag}  [{src_show} 质量{info['quality']}]{view_note}  "
                f"置信 {conf * 100:.1f}%  状态 {st}")
            log(f"   类型 {r.get('invoice_type') or '—'} | "
                f"号码 {r.get('invoice_no') or '—'} | 日期 {r.get('issue_date') or '—'}")
            log(f"   购 {r.get('buyer_name') or '—'} ({r.get('buyer_tax') or '—'})")
            log(f"   销 {r.get('seller_name') or '—'} ({r.get('seller_tax') or '—'})")
            log(f"   金额/税额/合计 {r.get('amount')} / {r.get('tax')} / {r.get('total')}"
                f"  税率 {r.get('tax_rate') or '—'}")
            if issues:
                log(f"   问题 {summarize_issues(issues)}")
                bad += 1
            else:
                log("   问题 无 ✓")
                ok += 1
            log("")
    log("=" * 72)
    log(f"干跑结束：{ok} 张无问题，{bad} 张有问题（未写任何文件）")
    return 0


# ==========================================================================
# 10. 自检（不依赖任何文件 / 任何第三方包）
# ==========================================================================
SELFTEST_CASES = [
    {
        "name": "数电普票-正常",
        "text": """电子发票（普通发票）
发票号码：26312000000012345678     发票代码：031002600111
开票日期：2026年03月12日
购买方信息 名称：某某科技有限公司  统一社会信用代码：91310115MA1H8WXYQ4
销售方信息 名称：虚拟信息技术服务有限公司  统一社会信用代码：91110108MA01C2XY3P
项目名称 *信息技术服务*技术服务费
合计 ¥10000.00 ¥600.00
价税合计（小写）¥10600.00
税率 6%
开票人：张三""",
        "expect": {"invoice_no": "26312000000012345678", "total": 10600.0,
                   "amount": 10000.0, "tax": 600.0},
    },
    {
        "name": "金额勾稽不平-应报严重",
        "text": """增值税电子普通发票
发票号码：12345678  发票代码：031002600111
开票日期：2026年03月12日
购买方信息 名称：甲公司  纳税人识别号：91310115MA1H8WXYQ4
销售方信息 名称：乙公司  纳税人识别号：91110108MA01C2XY3P
合计金额 ¥1000.00 合计税额 ¥60.00
价税合计（小写）¥1120.00
税率 6%""",
        "expect_exception": "金额勾稽",
    },
    {
        "name": "税率与税额矛盾-应报严重",
        "text": """增值税电子普通发票
发票号码：87654321  发票代码：031002600111
开票日期：2026年03月12日
购买方信息 名称：甲公司  纳税人识别号：91310115MA1H8WXYQ4
销售方信息 名称：乙公司  纳税人识别号：91110108MA01C2XY3P
合计金额 ¥1000.00 合计税额 ¥130.00
价税合计（小写）¥1130.00
税率 6%""",
        "expect_exception": "税率×金额",
    },
    # ---- v2.1 新增用例：把这次修掉的四个缺陷固化成回归护栏 ----
    {
        "name": "整数金额（无小数点）-v2.1",
        "text": """电子发票（普通发票）
发票号码：26312000000012345678
开票日期：2026年03月12日
购买方信息 名称：某某科技有限公司  统一社会信用代码：91310115MA1H8WXYQ4
销售方信息 名称：虚拟信息技术服务有限公司  统一社会信用代码：91110108MA01C2XY3P
合计 ¥10000 ¥600
价税合计（小写）¥10600
税率 6%""",
        "expect": {"amount": 10000.0, "tax": 600.0, "total": 10600.0},
    },
    {
        "name": "铁路电子客票-票种识别-v2.1",
        "text": """电子发票（铁路电子客票）
发票号码：26312000000012345678
开票日期：2026年03月12日
购买方信息 名称：某某科技有限公司  统一社会信用代码：91310115MA1H8WXYQ4
车次 G1234  北京南站 上海虹桥站
票价 ¥106.00
价税合计（小写）¥106.00
税率 9%""",
        "expect": {"invoice_type": "电子发票(铁路电子客票)", "total": 106.0},
        "expect_no_exception": "销售方名称",
    },
    {
        "name": "金额与发票号码矛盾-严重问题应压低置信度-v2.1",
        "text": """电子发票（普通发票）
发票号码：26312000000012345678
开票日期：2026年03月12日
购买方信息 名称：某某科技有限公司  统一社会信用代码：91310115MA1H8WXYQ4
销售方信息 名称：虚拟信息技术服务有限公司  统一社会信用代码：91110108MA01C2XY3P
合计 ¥1000.00 ¥60.00
价税合计（小写）¥1120.00
税率 6%""",
        "expect_max_conf": 0.85,          # 严重问题必须把 amount/tax/total 打折
    },
]


def selftest() -> int:
    print("=" * 72)
    print(f"自检 · idle_invoice_bot v{__version__}")
    print("=" * 72)
    cfg = load_config()
    templates = load_templates()
    failed = 0
    for case in SELFTEST_CASES:
        recs = parse_invoice(case["text"], "selftest", templates)
        if not recs:
            print(f"✗ {case['name']}：解析不出记录")
            failed += 1
            continue
        r = recs[0]
        issues = validate(r, cfg)
        adjust_confidence(r, issues, cfg)
        conf = overall_confidence(r)
        status = decide_status(issues, conf, r, cfg)

        problems = []
        for k, v in (case.get("expect") or {}).items():
            got = r.get(k)
            if isinstance(v, float):
                if not isinstance(got, (int, float)) or abs(got - v) > 1e-6:
                    problems.append(f"{k} 期望 {v} 实际 {got}")
            elif str(got) != str(v):
                problems.append(f"{k} 期望 {v} 实际 {got}")
        if case.get("expect_exception"):
            if not any(case["expect_exception"] in i["field"] for i in issues):
                problems.append(f"未命中预期问题「{case['expect_exception']}」")
        if case.get("expect_no_exception"):
            needle = case["expect_no_exception"]
            if any(needle in i["field"] for i in issues):
                problems.append(f"不该报出的问题「{needle}」被误报")
        if case.get("expect_max_conf") is not None:
            if conf > float(case["expect_max_conf"]):
                problems.append(f"置信度 {conf:.3f} 高于上限 "
                                f"{case['expect_max_conf']}（严重问题未有效压低置信度）")

        tag = "✓" if not problems else "✗"
        print(f"{tag} {case['name']}   置信 {conf*100:.1f}%  状态 {status}")
        for p in problems:
            print(f"     - {p}")
        failed += bool(problems)

    # 文本质量闸门自检
    good = text_quality(SELFTEST_CASES[0]["text"])
    bad = text_quality("||| ~~~ ^^^ ___\ufffd\ufffd\ufffd")
    print(f"{'✓' if good['score'] > bad['score'] else '✗'} 文本质量闸门："
          f"正常票面 {good['score']} > 噪声文本 {bad['score']}")
    failed += not (good["score"] > bad["score"])

    # 规则注册表自检
    ids = [r["id"] for r in RULES]
    ok_ids = len(ids) == len(set(ids))
    print(f"{'✓' if ok_ids else '✗'} 规则注册表：{len(RULES)} 条规则，id 唯一 = {ok_ids}")
    failed += not ok_ids

    # 多视图择优判据：模拟"标签与取值被拆散"的文本视图（PyMuPDF 按区块排序时会这样），
    # 断言它的得分明显低于正常视图 —— 这是"多视图择优"能救回来的前提。
    good_recs = parse_invoice(SELFTEST_CASES[0]["text"], "selftest", templates)
    split_view = (
        "26312000000012345678\n2026-03-12\n10000.00\n600.00\n10600.00\n6%\n"
        "某某科技有限公司\n虚拟信息技术服务有限公司\n"
        "91310115MA1H8WXYQ4\n91110108MA01C2XY3P\n"
        "发票号码 开票日期 合计金额 合计税额 价税合计 税率 购买方信息 名称 "
        "销售方信息 名称 纳税人识别号 项目名称 开票人"
    )
    split_recs = parse_invoice(split_view, "selftest", templates)
    g, sp = _records_score(good_recs), _records_score(split_recs)
    ok_split = g > sp and sp < 0.6
    print(f"{'✓' if ok_split else '✗'} 多视图择优判据：正常视图 {g:.3f} > 拆散视图 {sp:.3f}"
          f"（且拆散视图 < 0.6）")
    failed += not ok_split

    # 关键标签探针：标签缺失的文本必须被判低分
    q_label_ok = text_quality(SELFTEST_CASES[0]["text"])
    q_label_bad = text_quality("12345678 2026-03-12 10600.00 10000.00 600.00 6%")
    ok_label = q_label_ok["labels"] >= 3 and q_label_bad["labels"] < 2 \
        and q_label_ok["score"] > q_label_bad["score"]
    print(f"{'✓' if ok_label else '✗'} 标签探针：正常票面命中 {q_label_ok['labels']} 个标签、"
          f"无标签文本命中 {q_label_bad['labels']} 个，得分 {q_label_ok['score']} > "
          f"{q_label_bad['score']}")
    failed += not ok_label

    print("-" * 72)
    print(f"自检结果：{'全部通过' if not failed else f'{failed} 项未通过'}")
    print("=" * 72)
    return 1 if failed else 0


def explain() -> int:
    cfg = load_config()
    print("=" * 72)
    print("当前生效配置")
    print("=" * 72)
    print(json.dumps(cfg, ensure_ascii=False, indent=2))
    print("-" * 72)
    print(f"已注册校验规则 {len(RULES)} 条：")
    for r in RULES:
        mark = "（已禁用）" if r["id"] in (cfg.get("rules_disabled") or []) else ""
        print(f"  {r['id']}  [{r['level']}]  {r['title']}{mark}")
    print("-" * 72)
    tpls = load_templates()
    print(f"已加载版式模板 {len(tpls)} 个：")
    for t in tpls:
        print(f"  · {t.get('name')}  keywords={t.get('keywords')}  "
              f"字段规则 {len((t.get('fields') or {}))} 组")
    print("-" * 72)
    print(f"运行环境：PyMuPDF={_have('pymupdf') or _have('fitz')}  "
          f"openpyxl={_have('openpyxl')}  Pillow={_have('PIL')}  "
          f"OCR={ocr_provider() or '无'}")
    print("=" * 72)
    return 0


# ==========================================================================
# 11. CLI
# ==========================================================================
def main(argv=None) -> int:
    setup_console()
    ap = argparse.ArgumentParser(
        description="发票识别工作流机器人（纯 Python 单文件版）",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-i", "--input", help="发票目录")
    ap.add_argument("-o", "--output-dir", help="输出目录")
    ap.add_argument("--dry-run", action="store_true", help="只解析不写文件")
    ap.add_argument("--selftest", action="store_true", help="内置自检，不读任何文件")
    ap.add_argument("--explain", action="store_true", help="打印当前配置、规则与模板")
    ap.add_argument("--force", action="store_true",
                    help="忽略文件指纹，强制重新解析并替换台账中对应行")
    ap.add_argument("--quiet", action="store_true", help="精简输出")
    ap.add_argument("--ocr-worker", nargs="*", default=None,
                    help=argparse.SUPPRESS)     # 内部用：OCR 子进程入口
    # IDLE 的 sys.argv 可能带 -c 等参数，用 parse_known_args 避免直接报错
    args, _unknown = ap.parse_known_args(argv)

    # OCR 子进程分支：必须在任何配置/IO 之前返回
    if args.ocr_worker is not None:
        return _ocr_worker_main(args.ocr_worker)

    if args.selftest:
        return selftest()
    if args.explain:
        return explain()

    cfg = load_config()
    if args.input:
        cfg["input_dir"] = str(Path(args.input).resolve())
    if args.output_dir:
        cfg["output_dir"] = str(Path(args.output_dir).resolve())
    if args.force:
        cfg["force"] = True

    if args.dry_run:
        return dry_run(cfg, load_templates())

    log = print
    if args.quiet:
        def log(*a, **k):                               # noqa: ANN001
            pass

    result = run(cfg, log=log, verbose=not args.quiet)
    if result.get("failed"):
        return 3
    st = result.get("stats") or {}
    return 2 if (st.get("error") or st.get("ocr_todo")) else 0


if __name__ == "__main__":
    raise SystemExit(main())
