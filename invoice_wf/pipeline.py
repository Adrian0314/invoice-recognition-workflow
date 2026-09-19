# -*- coding: utf-8 -*-
"""
流程编排层
----------
输入目录扫描 → 去重预检 → 文本提取(含 OCR) → 字段解析 → 规则校验
→ 重复校验 → 写入 Excel 台账 → 生成异常记录 / 汇总 / HTML 报告

任何一个文件出错都不会中断整批处理，错误会被降级为「异常记录 + 台账行」。
"""
from __future__ import annotations

import traceback
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from . import reader
from .config import (SETTINGS, STATUS_DUPLICATE, STATUS_ERROR, STATUS_OCR_TODO,
                     STATUS_OK, STATUS_WARN)
from .excel_store import Ledger, dedupe_key
from .parser import parse_invoice
from .report import build_report
from .validator import LEVEL_ERROR, LEVEL_WARN, rollup_status, summarize_issues, validate


@dataclass
class Options:
    input_dir: str = "invoices_input"
    ledger: str = "output/发票台账.xlsx"
    report: str = "output/处理报告.html"
    recursive: bool = True
    retry_failed: bool = False      # 重试「待人工录入」的文件
    force: bool = False             # 忽略文件指纹，强制重新解析
    verbose: bool = True
    log: callable = print


def _empty_record(file_name: str, file_hash: str) -> dict:
    rec = {k: "" for k in (
        "invoice_type", "invoice_code", "invoice_no", "issue_date", "buyer_name",
        "buyer_tax", "seller_name", "seller_tax", "amount", "tax", "total",
        "tax_rate", "item_name", "check_code", "drawer")}
    rec.update({"file_name": file_name, "file_hash": file_hash,
                "processed_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "status": STATUS_OCR_TODO, "issue": "",
                "_evidence": {}, "_raw_text": ""})
    return rec


def process_file(path: Path, ledger: Ledger, opts: Options) -> dict:
    """处理单个文件，返回统计增量。"""
    fp = reader.file_fingerprint(path)
    name = path.name
    delta = {"new_files": 0, "new_invoices": 0, "duplicate": 0, "ocr_todo": 0,
             "skipped_files": 0}

    # ---- 文件指纹预检 ----
    old = ledger.by_hash.get(fp)
    if old is not None:
        if opts.retry_failed and str(old.get("status")) == STATUS_OCR_TODO:
            # 移除旧的待人工录入行，并清掉它此前「提取失败」类的历史问题记录
            ledger.remove_row(old)
            ledger.purge_issues(name)
        elif not (opts.force or SETTINGS["reparse_same_file"]):
            delta["skipped_files"] = 1
            old_name = str(old.get("file_name") or "").split(" [第")[0]
            if old_name == name:
                # 同一目录被反复扫描：正常的幂等行为，静默跳过，不计入异常
                opts.log(f"  [已处理] {name} —— 跳过")
            else:
                # 换了文件名再次提交、内容却完全相同 → 真正的重复提交，登记异常
                delta["duplicate"] += 1
                ledger.add_issue(
                    name, old.get("invoice_no", ""), "重复", "文件指纹",
                    f"该文件内容与台账中「{old_name}」完全一致（指纹 {fp}），已阻止重复入账",
                    "确认为重复提交时忽略；如需强制重解，使用 --force")
                opts.log(f"  [重复提交] {name} —— 与「{old_name}」内容完全相同，已拦截")
            return delta

    # ---- 文本提取 ----
    try:
        info = reader.extract_text(path)
    except reader.ExtractionError as exc:
        rec = _empty_record(name, fp)
        rec["issue"] = f"[严重] 文本提取: {exc}"
        ledger.add_record(rec)
        ledger.add_issue(name, "", "严重", "文本提取", str(exc),
                         "改为人工录入，或安装 OCR 引擎后使用 --retry-failed 重试")
        ledger.stats["ocr_todo"] += 1
        delta["new_files"] = delta["ocr_todo"] = 1
        opts.log(f"  [待人工] {name} —— {str(exc).splitlines()[0]}")
        return delta

    # ---- 字段解析 ----
    try:
        records = parse_invoice(info["text"], name)
    except Exception as exc:                                # noqa: BLE001
        ledger.add_issue(name, "", "严重", "字段解析", f"解析异常：{exc}",
                         "请把该文件反馈给流程维护者")
        opts.log(f"  [异常] {name} 解析失败：{exc}")
        if opts.verbose:
            traceback.print_exc()
        return delta

    if not records:
        ledger.add_issue(name, "", "严重", "字段解析", "未从文件中解析出任何发票记录",
                         "确认是否为发票文件（非发票 PDF 会被忽略）")
        return delta

    delta["new_files"] += 1
    src_note = {"pdf-text": "PDF 文本层", "ocr": f"OCR({info.get('ocr_engine')})"}.get(
        info["source"], info["source"])

    # 同一文件含多张发票时，第 2 张起不再做「文件指纹重复」判定
    # （它们本来就共享同一个指纹，否则会被误判为重复文件）
    committed_in_this_file = False

    for rec in records:
        rec["file_hash"] = fp
        rec["processed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # ---- 规则校验 ----
        issues = validate(rec)
        rec["status"] = rollup_status(issues)
        rec["issue"] = summarize_issues(issues)
        for it in issues:
            ledger.add_issue(rec["file_name"].split(" [第")[0], rec.get("invoice_no", ""),
                             it["level"], it["field"], it["message"], it["suggest"])

        # ---- 重复校验 ----
        dup_kind, old_row = ledger.find_duplicate(
            rec, "" if committed_in_this_file else fp)
        if dup_kind:
            delta["duplicate"] += 1
            where = old_row.get("file_name", "台账已有行") if old_row else "台账已有行"
            seq = old_row.get("seq") if old_row else "?"
            ledger.add_issue(
                rec["file_name"].split(" [第")[0], rec.get("invoice_no", ""), "重复", "发票唯一性",
                f"{dup_kind}：与台账中「{where}」第 {seq} 行重复，未重复入账",
                "确认为重复报销时忽略；如需保留可改 config.py 中 on_duplicate=mark")
            opts.log(f"  [重复] {rec.get('invoice_no') or name} —— {dup_kind}，未入账")
            if SETTINGS["on_duplicate"] == "mark":
                rec["status"] = STATUS_DUPLICATE
                rec["issue"] = f"[重复] {dup_kind} —— 与「{where}」重复，需复核"
                ledger.add_record(rec)
                committed_in_this_file = True
                delta["new_invoices"] += 1
            continue

        ledger.add_record(rec)
        committed_in_this_file = True
        delta["new_invoices"] += 1
        flag = "" if rec["status"] == STATUS_OK else f"  ← {rec['status']}"
        total_show = rec.get("total")
        opts.log(f"  [入账] {rec.get('invoice_no') or '(无号码)'} "
                 f"{str(rec.get('seller_name') or '')[:16]} "
                 f"¥{total_show if total_show is not None else '?'} "
                 f"({src_note}){flag}")

    # 该文件本次成功识别 → 清掉它以前「压根没处理成功」类的历史问题记录
    # （例如装 OCR 之前留下的「无文本层」「未安装 OCR 引擎」），否则异常清单会误导使用者
    ledger.purge_issues(name)

    return delta


def run(opts: Options) -> dict:
    t0 = datetime.now()
    input_dir = Path(opts.input_dir)
    if not input_dir.exists():
        input_dir.mkdir(parents=True, exist_ok=True)

    files = reader.discover_files(input_dir, opts.recursive)
    opts.log("=" * 70)
    opts.log(f"发票自动化处理流程启动  {t0:%Y-%m-%d %H:%M:%S}")
    opts.log(f"输入目录：{input_dir.resolve()}")
    opts.log(f"台账文件：{Path(opts.ledger).resolve()}")
    opts.log(f"待处理文件：{len(files)} 个")
    if not reader.ocr_available():
        opts.log("提示：未检测到 OCR 引擎，图片/扫描件将被登记为『待人工录入』(不影响电子发票处理)")
    opts.log("-" * 70)

    ledger = Ledger(opts.ledger)

    for p in files:
        opts.log(f"处理：{p.name}")
        try:
            d = process_file(p, ledger, opts)
        except Exception as exc:                            # noqa: BLE001
            ledger.add_issue(p.name, "", "严重", "流程", f"处理时发生未预期异常：{exc}",
                             "请保留文件并反馈")
            opts.log(f"  [异常] {p.name}：{exc}")
            if opts.verbose:
                traceback.print_exc()
            continue
        for k in ("new_files", "new_invoices"):
            ledger.stats[k] = ledger.stats.get(k, 0) + d.get(k, 0)

    # 汇总（_recalc 会按台账实际状态重算 正常/异常/重复/待人工 与金额合计）
    ledger._recalc()

    # 保存时最常见的意外：台账正被 Excel 打开，文件被独占，写入会 PermissionError。
    # 这里给出可操作的提示，而不是抛一串看不懂的堆栈。
    try:
        ledger.save()
    except PermissionError as exc:
        opts.log("")
        opts.log("!" * 70)
        opts.log(f"[无法写入] {Path(opts.ledger).resolve()}")
        opts.log("该文件正被其他程序占用（最常见的是 Excel 还开着它）。")
        opts.log("请先关闭 Excel / WPS，再重新运行本程序。")
        opts.log("本次已解析的发票尚未写入台账，重跑即可（已入账的文件会自动跳过）。")
        opts.log("!" * 70)
        raise SystemExit(3) from exc

    try:
        extras = ledger.export_extras(Path(opts.ledger).parent)
    except PermissionError:
        opts.log("[提示] 导出 CSV/JSON 失败（文件可能正被 Excel 打开），台账本身已保存成功。")
        extras = {"json": "", "csv": "", "issue_csv": ""}

    summary = {
        "started_at": t0, "finished_at": datetime.now(),
        "elapsed": (datetime.now() - t0).total_seconds(),
        "input_dir": str(input_dir.resolve()),
        "files": len(files),
        "rows": len(ledger.rows),
        "issues": len(ledger.issue_rows),
        "stats": dict(ledger.stats),
        "extras": {k: str(v) for k, v in extras.items()},
        "ledger": str(Path(opts.ledger).resolve()),
        "report": str(Path(opts.report).resolve()),
        "ocr_available": reader.ocr_available(),
    }
    build_report(ledger, summary, opts.report)

    opts.log("-" * 70)
    s = ledger.stats
    opts.log(f"处理完成，用时 {summary['elapsed']:.2f}s")
    opts.log(f"台账总行数 {len(ledger.rows)}  |  正常 {s['ok']}  |  提示 {s['warn']}  |  "
             f"异常 {s['error']}  |  重复拦截 {s['duplicate']}  |  待人工 {s['ocr_todo']}")
    opts.log(f"合计金额 {s['sum_amount']:,.2f}  |  合计税额 {s['sum_tax']:,.2f}  |  "
             f"价税合计 {s['sum_total']:,.2f}")
    opts.log(f"台账：{Path(opts.ledger).resolve()}")
    opts.log(f"报告：{Path(opts.report).resolve()}")
    if ledger.issue_rows:
        opts.log(f"问题清单 {len(ledger.issue_rows)} 条，详见台账『异常记录』页")
    opts.log("=" * 70)
    return summary
