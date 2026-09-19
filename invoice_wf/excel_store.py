# -*- coding: utf-8 -*-
"""
Excel 台账层
------------
设计原则：**Excel 台账本身就是唯一数据源**。
- 每处理一批，先读回台账已有行 → 建立「发票号码」与「文件指纹」索引
- 新记录去重后追加，序号重排，样式重刷 → 保证「一张发票一行」且可跨批次去重
- 用户手工修改过的单元格值会被读回并保留（不覆盖人工结果）
"""
from __future__ import annotations

import json
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .config import (FIELDS, KEYS, SETTINGS, STATUS_DUPLICATE, STATUS_ERROR,
                     STATUS_OCR_TODO, STATUS_WARN)

# 配色（浅色主题友好）
CLR_HEADER_BG = "1F4E79"
CLR_HEADER_FG = "FFFFFF"
CLR_ROW_ERROR = "FCE4E4"
CLR_ROW_DUP = "FFF4D6"
CLR_STATUS_ERROR_FG = "C00000"
CLR_WARN_FG = "BF8F00"
CLR_OK_FG = "1E7145"

THIN = Side(style="thin", color="D0D7E5")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# 这些状态的行属于「未真正入账」，不计入金额合计
COUNTED_EXCLUDE = {STATUS_DUPLICATE, STATUS_OCR_TODO, "重复标记", "重复未入账"}


# 空号码在表中显示为「—」，回读做指纹比对时必须还原成空串，否则会被当成两个不同的值
PLACEHOLDER = "—"


def norm_inv_no(v) -> str:
    v = str(v or "").strip()
    return "" if v == PLACEHOLDER else v


def _issue_key(file_name: str, invoice_no, level: str, field: str, message: str) -> tuple:
    """问题的唯一指纹，用于「同一问题只登记一次」。"""
    return (str(file_name or ""), norm_inv_no(invoice_no),
            str(level or ""), str(field or ""), str(message or ""))


def dedupe_key(rec: dict) -> str:
    """发票唯一键：优先 发票代码+发票号码，其次仅发票号码。"""
    code = str(rec.get("invoice_code") or "").strip()
    no = str(rec.get("invoice_no") or "").strip()
    if not no:
        return ""
    return f"{code}-{no}" if code else no


class Ledger:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.rows: list[dict] = []
        self.issue_rows: list[dict] = []
        self._issue_keys: set = set()
        self.by_key: dict[str, dict] = {}
        self.by_hash: dict[str, dict] = {}
        self.stats = {
            "files_total": 0, "new_files": 0, "new_invoices": 0,
            "ok": 0, "warn": 0, "error": 0, "duplicate": 0,
            "ocr_todo": 0, "skipped_files": 0,
            "sum_amount": 0.0, "sum_tax": 0.0, "sum_total": 0.0,
        }
        self._load()

    # ------------------------------------------------------------------
    # 读取已有台账
    # ------------------------------------------------------------------
    def _load(self):
        if not self.path.exists():
            return
        try:
            wb = load_workbook(self.path)
        except Exception:
            return

        if SETTINGS["sheet_detail"] in wb.sheetnames:
            ws = wb[SETTINGS["sheet_detail"]]
            headers = [c.value for c in ws[1]]
            idx = {h: i for i, h in enumerate(headers) if h}
            for r in ws.iter_rows(min_row=2, values_only=True):
                if r is None or all(v is None or v == "" for v in r):
                    continue
                rec = {}
                for f in FIELDS:
                    j = idx.get(f["header"])
                    rec[f["key"]] = r[j] if (j is not None and j < len(r)) else ""
                # 归一化：序号转 int，金额保持原样
                try:
                    rec["seq"] = int(rec.get("seq"))
                except (TypeError, ValueError):
                    rec["seq"] = len(self.rows) + 1
                self.rows.append(rec)
                k = dedupe_key(rec)
                if k and k not in self.by_key:
                    self.by_key[k] = rec
                h = str(rec.get("file_hash") or "").strip()
                if h and h not in self.by_hash:
                    self.by_hash[h] = rec

        if SETTINGS["sheet_issue"] in wb.sheetnames:
            ws = wb[SETTINGS["sheet_issue"]]
            headers = [c.value for c in ws[1]]
            idx = {h: i for i, h in enumerate(headers) if h}
            for r in ws.iter_rows(min_row=2, values_only=True):
                if r is None or all(v is None or v == "" for v in r):
                    continue
                self.issue_rows.append({
                    h: r[idx[h]] if idx.get(h) is not None and idx[h] < len(r) else ""
                    for h in headers if h
                })
        # 重建「问题指纹」，避免重跑时重复登记同一条问题
        for it in self.issue_rows:
            self._issue_keys.add(_issue_key(
                it.get("源文件名"), it.get("发票号码"),
                it.get("问题级别"), it.get("涉及字段"), it.get("问题描述")))
        self.stats["files_total"] = len({
            str(r.get("file_name") or "").split(" [第")[0] for r in self.rows if r.get("file_name")
        })
        self._recalc()

    def _recalc(self):
        s = self.stats
        # 只有「真正入账」的行才计入金额合计（重复行、待人工录入行不计）
        counted = [r for r in self.rows
                   if str(r.get("status") or "") not in COUNTED_EXCLUDE]
        s["sum_amount"] = round(sum(float(r.get("amount") or 0) for r in counted), 2)
        s["sum_tax"] = round(sum(float(r.get("tax") or 0) for r in counted), 2)
        s["sum_total"] = round(sum(float(r.get("total") or 0) for r in counted), 2)
        s["ok"] = sum(1 for r in self.rows if r.get("status") == "正常")
        s["error"] = sum(1 for r in self.rows if r.get("status") == STATUS_ERROR)
        s["warn"] = sum(1 for r in self.rows if r.get("status") in (STATUS_WARN, "存在提示"))
        s["ocr_todo"] = sum(1 for r in self.rows if r.get("status") == STATUS_OCR_TODO)
        s["duplicate"] = sum(1 for r in self.issue_rows
                             if str(r.get("问题级别")) == "重复")
        s["files_total"] = len({
            str(r.get("file_name") or "").split(" [第")[0]
            for r in self.rows if r.get("file_name")
        })

    def remove_row(self, rec: dict):
        """从台账移除一行（供 --retry-failed 重跑「待人工录入」的文件）。"""
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

    def purge_issues(self, file_name: str,
                     fields: tuple = ("文本提取", "流程", "字段解析")):
        """
        重跑某个文件前，清掉它此前「压根没处理成功」类的历史问题记录。
        这些记录在本次成功识别后已不成立，留着会让异常清单与统计误导使用者。
        """
        keep = []
        for it in self.issue_rows:
            if (str(it.get("源文件名") or "") == file_name
                    and str(it.get("涉及字段") or "") in fields):
                self._issue_keys.discard(_issue_key(
                    it.get("源文件名"), it.get("发票号码"),
                    it.get("问题级别"), it.get("涉及字段"), it.get("问题描述")))
                continue
            keep.append(it)
        if len(keep) != len(self.issue_rows):
            self.issue_rows = keep
            for i, it in enumerate(self.issue_rows, 1):
                it["序号"] = i

    # ------------------------------------------------------------------
    # 去重查询
    # ------------------------------------------------------------------
    def find_duplicate(self, rec: dict, file_hash: str) -> tuple[str, dict | None]:
        """返回 (重复类型, 已存在的那一行)。无重复返回 ('', None)。"""
        k = dedupe_key(rec)
        if k and k in self.by_key:
            return "发票级重复", self.by_key[k]
        if file_hash and file_hash in self.by_hash:
            return "文件级重复", self.by_hash[file_hash]
        return "", None

    # ------------------------------------------------------------------
    # 写入
    # ------------------------------------------------------------------
    def add_record(self, rec: dict):
        rec["seq"] = len(self.rows) + 1
        self.rows.append(rec)
        k = dedupe_key(rec)
        if k and k not in self.by_key:
            self.by_key[k] = rec
        h = str(rec.get("file_hash") or "")
        if h and h not in self.by_hash:
            self.by_hash[h] = rec

    def add_issue(self, file_name: str, invoice_no: str, level: str,
                  field: str, message: str, suggest: str = ""):
        # 同一问题只登记一次：反复重跑同一目录不会让异常记录与统计无限膨胀
        key = _issue_key(file_name, invoice_no, level, field, message)
        if key in self._issue_keys:
            return
        self._issue_keys.add(key)
        self.issue_rows.append({
            "序号": len(self.issue_rows) + 1,
            "源文件名": file_name,
            "发票号码": invoice_no or "—",
            "问题级别": level,
            "涉及字段": field,
            "问题描述": message,
            "建议处理": suggest or "人工核对",
            "发生时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        })

    # ------------------------------------------------------------------
    # 落盘
    # ------------------------------------------------------------------
    def save(self, extra_summary: dict | None = None):
        wb = Workbook()
        ws = wb.active
        ws.title = SETTINGS["sheet_detail"]
        self._write_detail(ws)
        self._write_issues(wb.create_sheet(SETTINGS["sheet_issue"]))
        self._write_summary(wb.create_sheet(SETTINGS["sheet_summary"]), extra_summary or {})
        wb.save(self.path)

    def _write_detail(self, ws):
        # 表头
        for j, f in enumerate(FIELDS, 1):
            c = ws.cell(row=1, column=j, value=f["header"])
            c.font = Font(bold=True, color=CLR_HEADER_FG, size=11)
            c.fill = PatternFill("solid", fgColor=CLR_HEADER_BG)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = BORDER
            letter = get_column_letter(j)
            ws.column_dimensions[letter].width = f["width"]
            if f.get("hidden"):
                ws.column_dimensions[letter].hidden = True
        ws.row_dimensions[1].height = 30

        money_cols = {j for j, f in enumerate(FIELDS, 1) if f["kind"] == "money"}
        int_cols = {j for j, f in enumerate(FIELDS, 1) if f["kind"] == "int"}

        for i, rec in enumerate(self.rows, start=2):
            status = str(rec.get("status") or "")
            row_fill = None
            if status == STATUS_ERROR:
                row_fill = PatternFill("solid", fgColor=CLR_ROW_ERROR)
            elif status == STATUS_DUPLICATE:
                row_fill = PatternFill("solid", fgColor=CLR_ROW_DUP)

            for j, f in enumerate(FIELDS, 1):
                val = rec.get(f["key"], "")
                if f["kind"] == "money" and val in ("", None):
                    val = None
                c = ws.cell(row=i, column=j, value=val)
                c.border = BORDER
                if row_fill:
                    c.fill = row_fill
                if j in money_cols:
                    c.number_format = "#,##0.00"
                    c.alignment = Alignment(horizontal="right", vertical="center")
                elif j in int_cols:
                    c.alignment = Alignment(horizontal="center", vertical="center")
                elif f["key"] in ("status", "invoice_type", "tax_rate", "issue_date"):
                    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                elif f["key"] == "issue":
                    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                else:
                    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)

                if f["key"] == "status":
                    if status == STATUS_ERROR:
                        c.font = Font(bold=True, color=CLR_STATUS_ERROR_FG)
                    elif status in (STATUS_WARN, "存在提示", STATUS_DUPLICATE):
                        c.font = Font(bold=True, color=CLR_WARN_FG)
                    else:
                        c.font = Font(bold=True, color=CLR_OK_FG)
                elif f["key"] == "issue" and status == STATUS_ERROR:
                    c.font = Font(color=CLR_STATUS_ERROR_FG)

        ws.freeze_panes = "C2"
        if self.rows:
            ws.auto_filter.ref = f"A1:{get_column_letter(len(FIELDS))}{len(self.rows) + 1}"

    def _write_issues(self, ws):
        headers = ["序号", "源文件名", "发票号码", "问题级别", "涉及字段",
                   "问题描述", "建议处理", "发生时间"]
        widths = [6, 32, 22, 10, 16, 56, 34, 20]
        for j, (h, w) in enumerate(zip(headers, widths), 1):
            c = ws.cell(row=1, column=j, value=h)
            c.font = Font(bold=True, color=CLR_HEADER_FG, size=11)
            c.fill = PatternFill("solid", fgColor=CLR_HEADER_BG)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = BORDER
            ws.column_dimensions[get_column_letter(j)].width = w
        ws.row_dimensions[1].height = 28

        lvl_font = {"严重": Font(color=CLR_STATUS_ERROR_FG, bold=True),
                    "警告": Font(color=CLR_WARN_FG, bold=True),
                    "提示": Font(color="44546A"),
                    "重复": Font(color="BF8F00", bold=True)}
        lvl_fill = {"严重": PatternFill("solid", fgColor=CLR_ROW_ERROR),
                    "警告": PatternFill("solid", fgColor="FFF4D6"),
                    "重复": PatternFill("solid", fgColor=CLR_ROW_DUP)}

        for i, row in enumerate(self.issue_rows, start=2):
            for j, h in enumerate(headers, 1):
                c = ws.cell(row=i, column=j, value=row.get(h, ""))
                c.border = BORDER
                c.alignment = Alignment(
                    horizontal="left" if h in ("问题描述", "建议处理", "源文件名") else "center",
                    vertical="center", wrap_text=h in ("问题描述", "建议处理"))
                if h == "问题级别":
                    lv = str(row.get(h, ""))
                    c.font = lvl_font.get(lv, Font())
                    if lv in lvl_fill:
                        c.fill = lvl_fill[lv]
        ws.freeze_panes = "A2"
        if self.issue_rows:
            ws.auto_filter.ref = f"A1:H{len(self.issue_rows) + 1}"

    def _write_summary(self, ws, extra: dict):
        ws.column_dimensions["A"].width = 26
        for col, w in zip("BCDEF", [18, 18, 18, 18, 18]):
            ws.column_dimensions[col].width = w

        title = ws.cell(row=1, column=1, value="发票处理汇总统计")
        title.font = Font(bold=True, size=14, color=CLR_HEADER_BG)
        ws.merge_cells("A1:E1")

        s = self.stats
        rows = [
            ("统计项", "数值", "说明"),
            ("处理文件数", s.get("files_total", 0), "台账内累计涉及的发票文件数（含多票文件）"),
            ("本次新增文件数", s.get("new_files", 0), "本次运行新处理的文件数量"),
            ("本次新增发票数", s.get("new_invoices", 0), "本次新写入台账的发票行数"),
            ("台账发票总行数", len(self.rows), "一张发票一行"),
            ("正常", s.get("ok", 0), "校验全部通过"),
            ("存在提示/警告", s.get("warn", 0), "有警告或提示，建议复核"),
            ("异常", s.get("error", 0), "存在严重问题，需人工处理"),
            ("重复/重复提交拦截", s.get("duplicate", 0), "重复发票已阻止重复入账（详见异常记录）"),
            ("待人工录入", s.get("ocr_todo", 0), "图片/扫描件无法自动识别，需人工补录"),
            ("已处理文件跳过", s.get("skipped_files", 0), "同一目录反复扫描时自动跳过，属正常幂等"),
            ("", "", ""),
            ("合计金额(不含税)", s.get("sum_amount", 0.0), "所有台账行金额合计"),
            ("合计税额", s.get("sum_tax", 0.0), "所有台账行税额合计"),
            ("合计价税合计", s.get("sum_total", 0.0), "所有台账行价税合计"),
            ("", "", ""),
            ("生成时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S"), ""),
            ("台账文件", str(self.path), ""),
        ]

        for i, (a, b, c_) in enumerate(rows, start=3):
            ca, cb, cc = (ws.cell(row=i, column=1, value=a),
                          ws.cell(row=i, column=2, value=b),
                          ws.cell(row=i, column=3, value=c_))
            if i == 3:
                for c in (ca, cb, cc):
                    c.font = Font(bold=True, color=CLR_HEADER_FG)
                    c.fill = PatternFill("solid", fgColor=CLR_HEADER_BG)
                    c.alignment = Alignment(horizontal="center", vertical="center")
                continue
            ca.font = Font(bold=(isinstance(b, (int, float)) and b != 0) or a in
                           ("合计金额(不含税)", "合计税额", "合计价税合计"))
            if isinstance(b, float):
                cb.number_format = "#,##0.00"
                cb.alignment = Alignment(horizontal="right", vertical="center")
            elif isinstance(b, int):
                cb.alignment = Alignment(horizontal="center", vertical="center")
            cc.font = Font(size=9, color="808080")

        # 按销售方汇总
        r0 = len(rows) + 5
        ws.cell(row=r0, column=1, value="按销售方汇总（Top 20）").font = Font(
            bold=True, size=12, color=CLR_HEADER_BG)
        heads = ["销售方名称", "张数", "金额(不含税)", "税额", "价税合计"]
        agg = defaultdict(lambda: [0, 0.0, 0.0, 0.0])
        for r in self.rows:
            k = str(r.get("seller_name") or "（未识别）")
            agg[k][0] += 1
            agg[k][1] += float(r.get("amount") or 0)
            agg[k][2] += float(r.get("tax") or 0)
            agg[k][3] += float(r.get("total") or 0)
        top = sorted(agg.items(), key=lambda kv: -kv[1][3])[:20]
        for j, h in enumerate(heads, 1):
            c = ws.cell(row=r0 + 1, column=j, value=h)
            c.font = Font(bold=True, color=CLR_HEADER_FG)
            c.fill = PatternFill("solid", fgColor=CLR_HEADER_BG)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = BORDER
        for i, (name, v) in enumerate(top, start=r0 + 2):
            ws.cell(row=i, column=1, value=name).border = BORDER
            ws.cell(row=i, column=2, value=v[0]).border = BORDER
            for j, val in enumerate(v[1:], start=3):
                c = ws.cell(row=i, column=j, value=round(val, 2))
                c.number_format = "#,##0.00"
                c.border = BORDER
            ws.cell(row=i, column=2).alignment = Alignment(horizontal="center")

    # ------------------------------------------------------------------
    # 旁路导出（便于导入其他系统）
    # ------------------------------------------------------------------
    def export_extras(self, out_dir: str | Path):
        out = Path(out_dir)
        out.mkdir(parents=True, exist_ok=True)

        detail_json = out / "发票明细.json"
        payload = []
        for r in self.rows:
            item = {f["header"]: r.get(f["key"], "") for f in FIELDS if not f.get("hidden")}
            item["识别证据"] = ""
            payload.append(item)
        detail_json.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

        from .csv_util import text_cell, write_csv

        csv_path = out / "发票明细.csv"
        # 发票代码 / 发票号码 / 税号等纯数字长串按文本导出，
        # 否则 Excel 打开会丢前导 0（011002000001 → 11002000001）、
        # 20 位号码变成科学计数法
        write_csv(csv_path,
                  [f["header"] for f in FIELDS if not f.get("hidden")],
                  [[text_cell(r.get(f["key"], ""))
                    for f in FIELDS if not f.get("hidden")] for r in self.rows])

        issue_csv = out / "异常记录.csv"
        if self.issue_rows:
            fields = list(self.issue_rows[0].keys())
            write_csv(issue_csv, fields,
                      [[text_cell(r.get(k, "")) for k in fields]
                       for r in self.issue_rows])
        return {"json": detail_json, "csv": csv_path, "issue_csv": issue_csv}
