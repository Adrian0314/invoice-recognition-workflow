# -*- coding: utf-8 -*-
"""
CSV 导出时保住「长数字串」的原样
=================================
问题：发票代码（12 位）、发票号码（20 位）、税号这类值都是纯数字，写进 CSV 后
用 Excel 打开会被当成数字处理 —— 前导 0 被丢掉（011002000001 → 11002000001），
20 位号码还会被显示成 2.4312E+19 科学计数法。

做法：把这类字段改写成 Excel 的文本写法 ="..."
      —— Excel / WPS 打开时按文本显示，位数与内容原样保留。

⚠ 不能用标准 csv 模块写这类字段：
   csv.writer 看到字段里含引号，会再套一层引号并转义，落盘就带上了多余的引号，
   Excel 只会把它当普通文本，单元格里显示成 ="011002000001" ，
   比丢前导 0 更糟。所以这里自己控制落盘：已按文本公式处理的字段原样写出，
   不做二次转义。

其他程序若需要原始值，用 parse_cell() 反解，或直接读 发票明细.json（那边是原值）。
"""
from __future__ import annotations

import re

# 只保护「纯数字且长度 >= 8」的值：
#   · 8 位以上才可能丢信息（短数字 Excel 显示无损）
#   · 含字母的税号（如 91310115MA1H8WXYQ4）Excel 本来就按文本处理，无需包装
_LONG_DIGITS = re.compile(r"^\d{8,}$")


def text_cell(value):
    """返回单元格的「落盘字符串」：长数字串包一层 ="" 以便 Excel 按文本处理。"""
    if isinstance(value, str) and _LONG_DIGITS.match(value):
        return f'="{value}"'
    if value is None:
        return ""
    return value


def text_row(row) -> list:
    """对整行逐格处理。"""
    return [text_cell(v) for v in row]


def parse_cell(value: str) -> str:
    """把 ="" 写法还原成原始值（供需要读回 CSV 的程序使用）。"""
    if isinstance(value, str) and len(value) > 3 \
            and value.startswith('="') and value.endswith('"'):
        return value[2:-1]
    return value


def _escape(value) -> str:
    """单元格落盘转义；已包成 ="..." 的字段原样写出，避免二次转义。"""
    s = "" if value is None else str(value)
    if s.startswith('="') and s.endswith('"') and len(s) > 3:
        return s
    if any(ch in s for ch in ',"\r\n'):
        return '"' + s.replace('"', '""') + '"'
    return s


def write_csv(path, header, rows, encoding: str = "utf-8-sig") -> None:
    """
    写 CSV：表头 + 数据行。

    与 csv.writer 的差别只有一处：`="..."` 文本公式原样落盘。
    行尾用 CRLF，符合 RFC 4180，也是 Excel 最省心的形式。
    """
    with open(path, "w", newline="", encoding=encoding) as fh:
        fh.write(",".join(_escape(h) for h in header) + "\r\n")
        for row in rows:
            fh.write(",".join(_escape(v) for v in row) + "\r\n")
