# -*- coding: utf-8 -*-
"""
文件读取层：PDF 文本提取 + 图片/扫描件 OCR（可插拔）
------------------------------------------------------
优先级：
  1) PDF 优先读取内嵌文本层（绝大多数电子发票都是原生 PDF，速度快、准确率 100%）
  2) 文本层过少（扫描件）或输入本身是图片 → 走 OCR
     OCR 引擎按 PaddleOCR → RapidOCR → Tesseract 顺序自动探测；
     一个都没有时不静默失败，而是抛出明确原因，由上层生成「待人工录入」行 + 异常提示。
"""
from __future__ import annotations

import io
import os
from pathlib import Path
from typing import Optional, Tuple

from .config import IMAGE_EXTS, PDF_EXTS, SETTINGS

_ENGINE_CACHE: Optional[Tuple[Optional[str], Optional[object]]] = None
_PROBE_CACHE: Optional[str] = None      # "" 表示探测过且没有可用引擎


class ExtractionError(Exception):
    """无法从文件中取得可用文本。"""


# --------------------------------------------------------------------------
# OCR 引擎探测
# --------------------------------------------------------------------------
# 说明：探测(便宜)与实例化(昂贵)分开。
#   探测只查包在不在，用于启动时打印提示，不会加载模型；
#   实例化只在实际需要识别图片时才发生——否则每次处理纯电子发票都要白等模型加载。
def _module_exists(name: str) -> bool:
    import importlib.util

    try:
        return importlib.util.find_spec(name) is not None
    except (ImportError, ValueError):
        return False


def _tesseract_binary_ok() -> bool:
    """pytesseract 只是壳，真正干活的是 Tesseract-OCR 主程序。"""
    try:
        import pytesseract  # type: ignore

        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def ocr_provider() -> Optional[str]:
    """
    低成本探测可用引擎（不加载模型）。优先级：
      PaddleOCR（中文票据最准） → RapidOCR（纯 pip 免编译，Windows 最省事） → Tesseract
    """
    global _PROBE_CACHE
    if _PROBE_CACHE is not None:
        return _PROBE_CACHE or None

    if _module_exists("paddleocr"):
        _PROBE_CACHE = "PaddleOCR"
    elif _module_exists("rapidocr") or _module_exists("rapidocr_onnxruntime"):
        _PROBE_CACHE = "RapidOCR"
    elif _module_exists("pytesseract") and _tesseract_binary_ok():
        _PROBE_CACHE = "Tesseract"
    else:
        _PROBE_CACHE = ""
    return _PROBE_CACHE or None


def _build_engine():
    """按需实例化引擎；返回 (引擎名, 可调用对象)，不可用时 (None, None)。"""
    name = ocr_provider()
    try:
        if name == "PaddleOCR":
            from paddleocr import PaddleOCR  # type: ignore

            return name, PaddleOCR(use_angle_cls=True, lang="ch", show_log=False)
        if name == "RapidOCR":
            try:
                from rapidocr import RapidOCR  # type: ignore   # 3.x
            except ImportError:
                from rapidocr_onnxruntime import RapidOCR  # type: ignore  # 1.x/2.x
            return name, RapidOCR()
        if name == "Tesseract":
            import pytesseract  # type: ignore

            return name, pytesseract
    except Exception:
        return None, None
    return None, None


def _load_ocr_engine():
    global _ENGINE_CACHE
    if _ENGINE_CACHE is None:
        _ENGINE_CACHE = _build_engine()
    return _ENGINE_CACHE


def ocr_available() -> bool:
    return ocr_provider() is not None


def ocr_hint() -> str:
    return (
        "未检测到可用 OCR 引擎。如需自动识别图片/扫描件，请任选其一安装：\n"
        "  A. RapidOCR（推荐，纯 pip 安装、无需编译，Windows 上最省事）：\n"
        "       pip install rapidocr onnxruntime\n"
        "  B. PaddleOCR（中文票据准确率最高，但 Python 3.9 以下/Windows 易装不上）：\n"
        "       pip install paddleocr paddlepaddle\n"
        "  C. Tesseract：先装 Tesseract-OCR 主程序（勾选 chi_sim 语言包），再 pip install pytesseract\n"
        "注意：必须装到「你运行本程序的同一个 Python 环境」里，装到别的环境不会被用到。\n"
        "未安装时，图片类文件会被登记为「待人工录入」并写入异常提示，不会静默丢失。"
    )


def _rapidocr_text(engine, image) -> str:
    """
    RapidOCR 的返回结构在不同大版本间不一致，两种都兼容：
      3.x     : RapidOCROutput(txts=[...], boxes=..., scores=...)
      1.x/2.x : ([[box, (text, score)], ...], elapse_list)
    """
    import numpy as np  # type: ignore

    out = engine(np.array(image.convert("RGB")))

    txts = getattr(out, "txts", None)
    if txts:
        return "\n".join(str(t) for t in txts if t)

    if isinstance(out, tuple):
        out = out[0]

    lines = []
    for item in out or []:
        try:
            txt = item[1]
            if isinstance(txt, (list, tuple)):
                txt = txt[0]
            if txt:
                lines.append(str(txt))
        except (IndexError, TypeError):
            continue
    return "\n".join(lines)


def _run_ocr_inprocess(image) -> str:
    """
    在**本进程内**执行 OCR。

    仅供 ocr_worker 子进程调用；主流程请一律使用 _run_ocr_many()，
    因为原生推理库崩溃时本进程会直接被操作系统终止，Python 层拦不住。
    """
    name, engine = _load_ocr_engine()
    if not name:
        raise ExtractionError("未安装 OCR 引擎，无法识别图片/扫描件。" + ocr_hint())

    try:
        if name == "PaddleOCR":
            import numpy as np  # type: ignore

            result = engine.ocr(np.array(image.convert("RGB")), cls=True)
            lines = []
            for page in result or []:
                for item in page or []:
                    try:
                        lines.append(item[1][0])
                    except Exception:
                        continue
            return "\n".join(lines)

        if name == "RapidOCR":
            return _rapidocr_text(engine, image)

        # Tesseract
        import pytesseract  # type: ignore

        return pytesseract.image_to_string(image, lang="chi_sim+eng")
    except ExtractionError:
        raise
    except Exception as exc:                                # noqa: BLE001
        raise ExtractionError(
            f"{name} 识别失败：{exc}\n"
            "提示：首次使用 RapidOCR/PaddleOCR 需要联网下载模型文件，"
            "请确认网络可访问 ModelScope 或 HuggingFace。"
        ) from exc


# 原生库崩溃时 Windows 给出的退出码（负数=被信号/异常终止）
_CRASH_CODES = {
    -1073741819: "0xC0000005 访问违例（原生库崩溃）",
    3221225477: "0xC0000005 访问违例（原生库崩溃）",
    3221226505: "0xC0000409 栈溢出/保护性终止",
    -1073741510: "0xC000013A 被强制中断",
}


def _run_ocr_many(images: list) -> list[str]:
    """
    在**独立子进程**中批量执行 OCR，返回与 images 等长的文本列表。

    子进程隔离的意义：ocr 依赖的原生库在个别机器上会直接崩掉整个进程
    （例如 onnxruntime 在老版本 Windows 上 import 即访问违例）。
    放在子进程里，崩溃只会让这一个文件降级为「待人工录入」，批量任务不会整体失败。
    """
    import json
    import subprocess
    import sys
    import tempfile

    if not ocr_provider():
        raise ExtractionError("未安装 OCR 引擎，无法识别图片/扫描件。" + ocr_hint())

    root = Path(__file__).resolve().parent.parent
    timeout = int(SETTINGS.get("ocr_timeout", 600))
    tmpdir = tempfile.mkdtemp(prefix="inv_ocr_")
    try:
        paths = []
        for i, img in enumerate(images):
            fp = Path(tmpdir) / f"page{i}.png"
            img.convert("RGB").save(fp, "PNG")
            paths.append(str(fp))

        cmd = [sys.executable, "-u", "-m", "invoice_wf.ocr_worker", *paths]
        try:
            proc = subprocess.run(cmd, cwd=str(root), capture_output=True,
                                  timeout=timeout, text=True,
                                  encoding="utf-8", errors="replace")
        except subprocess.TimeoutExpired as exc:
            raise ExtractionError(
                f"OCR 超时（超过 {timeout}s）。首次运行需下载模型文件，"
                f"可调大 config.py 里的 ocr_timeout，或先手工跑一次预热模型。"
            ) from exc

        # 从 stdout 里取最后一行 JSON（worker 用的是 ensure_ascii，全 ASCII 转义）
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
            code = proc.returncode
            reason = _CRASH_CODES.get(code, f"退出码 {code}")
            stderr_tail = (proc.stderr or "").strip()[-300:]
            raise ExtractionError(
                f"OCR 子进程未返回结果：{reason}。\n"
                f"这通常说明 OCR 引擎与本机系统/硬件不兼容（不影响电子发票识别）。\n"
                f"{('子进程最后输出：' + stderr_tail) if stderr_tail else ''}"
            )

        if not payload.get("ok"):
            raise ExtractionError(f"OCR 失败：{payload.get('error', '未知原因')}")

        texts = payload.get("texts") or []
        return (texts + [""] * len(images))[:len(images)]
    finally:
        import shutil

        shutil.rmtree(tmpdir, ignore_errors=True)


def _run_ocr(image) -> str:
    """单图 OCR（内部走子进程隔离）。"""
    return _run_ocr_many([image])[0]


# --------------------------------------------------------------------------
# 主入口
# --------------------------------------------------------------------------
def extract_text(path: str | Path) -> dict:
    """
    返回:
      {
        "text":   str,            提取到的原始文本
        "source": "pdf-text" | "ocr" | "none",
        "pages":  int,
        "ocr_engine": str | None,
        "note":   str,
      }
    失败时抛 ExtractionError（含可读原因）。
    """
    path = Path(path)
    ext = path.suffix.lower()
    if not path.exists():
        raise ExtractionError(f"文件不存在：{path}")

    if ext in PDF_EXTS:
        return _extract_pdf(path)
    if ext in IMAGE_EXTS:
        return _extract_image(path)
    raise ExtractionError(f"不支持的文件类型：{ext}")


def _extract_pdf(path: Path) -> dict:
    try:
        import pymupdf  # type: ignore
    except ImportError:  # 兼容旧包名
        import fitz as pymupdf  # type: ignore

    try:
        doc = pymupdf.open(str(path))
    except Exception as exc:
        raise ExtractionError(f"PDF 打开失败：{exc}") from exc

    try:
        texts = []
        for page in doc:
            texts.append(page.get_text("text") or "")
        text = "\n".join(texts)
        pages = doc.page_count

        if len(text.strip()) >= SETTINGS["pdf_min_text_chars"]:
            return {"text": text, "source": "pdf-text", "pages": pages,
                    "ocr_engine": None, "note": ""}

        # 文本层为空 → 扫描件，尝试 OCR（整份 PDF 的所有页一次交给子进程处理）
        if not ocr_available():
            raise ExtractionError(
                "PDF 无文本层（疑似扫描件），且未安装 OCR 引擎。" + ocr_hint()
            )

        from PIL import Image  # type: ignore

        zoom = pymupdf.Matrix(2.5, 2.5)  # 提高分辨率，利于小字号票据识别
        images = []
        for page in doc:
            pix = page.get_pixmap(matrix=zoom)
            images.append(Image.open(io.BytesIO(pix.tobytes("png"))))
        ocr_texts = _run_ocr_many(images)
        return {"text": "\n".join(ocr_texts), "source": "ocr", "pages": pages,
                "ocr_engine": ocr_provider(), "note": "PDF 无文本层，已走 OCR"}
    finally:
        doc.close()


def _extract_image(path: Path) -> dict:
    from PIL import Image  # type: ignore

    try:
        img = Image.open(str(path))
        img.load()
    except Exception as exc:
        raise ExtractionError(f"图片打开失败：{exc}") from exc

    if img.width < 800:  # 小图放大，提升 OCR 命中率
        scale = 1600 / max(img.width, 1)
        img = img.resize((int(img.width * scale), int(img.height * scale)))

    text = _run_ocr(img)
    if not text.strip():
        raise ExtractionError("OCR 未识别到任何文字，请确认图片清晰度或改用电子版发票。")
    return {"text": text, "source": "ocr", "pages": 1,
            "ocr_engine": ocr_provider(), "note": "图片文件，已走 OCR"}


# --------------------------------------------------------------------------
# 目录扫描
# --------------------------------------------------------------------------
def discover_files(input_dir: str | Path, recursive: bool = True) -> list[Path]:
    """按修改时间排序列出目录下所有受支持的文件。"""
    from .config import SUPPORTED_EXTS

    root = Path(input_dir)
    if not root.exists():
        return []
    it = root.rglob("*") if recursive else root.glob("*")
    files = [p for p in it if p.is_file()
             and p.suffix.lower() in SUPPORTED_EXTS
             and not p.name.startswith("~$")]
    return sorted(files, key=lambda p: (p.stat().st_mtime, p.name))


def file_fingerprint(path: str | Path) -> str:
    """文件内容 sha256 前 16 位，用于「同一文件重复提交」判定。"""
    import hashlib

    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()[:16]


def human_size(num: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if num < 1024:
            return f"{num:.0f}{unit}"
        num /= 1024
    return f"{num:.1f}TB"


def rel_display(path: str | Path, base: str | Path) -> str:
    try:
        return os.path.relpath(str(path), str(base))
    except Exception:
        return str(path)
