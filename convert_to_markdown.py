#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
多格式文件轉 Markdown 全功能工具 (Document to Markdown Pro Converter)
=====================================================================
支援功能：
  1. 多格式轉 Markdown：
     - PDF (.pdf)                 : 擷取文字、表格、內嵌圖片，並自動對掃描頁面進行 OCR 辨識
     - Excel (.xlsx, .xls, .csv)  : 將各工作表 (Sheet) 轉換為標準 Markdown 表格
     - Word (.docx, .doc)         : 保留標題階層 (H1~H6)、粗斜體、清單、表格，並匯出內嵌圖片與連結
     - PowerPoint (.pptx, .ppt)   : 擷取投影片標題、階層項目、表格、備忘稿，並匯出投影片圖片與連結
     - 純文字與圖片 (.txt, .json, .png, .jpg, .tiff, .webp) : 自動偵測編碼或執行 OCR 轉為 Markdown

  2. 掃描影像 OCR 辨識 (--ocr)：
     - 支援 RapidOCR (`rapidocr_onnxruntime`，純離線免安裝外部執行檔，最適合打包進單一 .exe)
     - 支援 Tesseract OCR (`pytesseract`，自動偵測同目錄或系統安裝的 `tesseract.exe`，預設 `chi_tra+eng`)
     - 支援 EasyOCR (`easyocr`) 自動備援

  3. 內嵌圖片自動匯出與 Markdown 連結：
     - 自動於輸出目錄建立 `<檔名>_assets/` 資料夾
     - 依原始段落/投影片順序匯出 Word、PPT、PDF 內的圖片並插入 `![圖片說明](<檔名>_assets/xxx.png)`

  4. 四種操作與打包模式：
     - 雙擊執行 / 視窗拖曳桌面介面 (GUI) : `python convert_to_markdown.py` (無參數時預設開啟 GUI) 或 `--gui`
     - 網頁拖曳上傳工具 (Web UI)         : `python convert_to_markdown.py --web --port 8000` (亦可從 GUI 內一鍵開啟)
     - 命令列 / 圖示拖放批次處理 (CLI)   : `python convert_to_markdown.py file1.pdf file2.docx` (支援直接將檔案拖放到 .exe 圖示上)
     - 一鍵打包 Windows 單一執行檔 (.exe): `python convert_to_markdown.py --build-exe`
"""

from __future__ import annotations

import argparse
import base64
import csv
import io
import json
import logging
import os
import platform
import shutil
import subprocess
import sys
import tempfile
import threading
import urllib.parse
import webbrowser
import zipfile
from dataclasses import dataclass, field
from email.parser import BytesParser
from email.policy import default as email_default_policy
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

logging.basicConfig(
    level=logging.INFO,
    format="[%(levelname)s] %(message)s",
)
logger = logging.getLogger("doc2md")


# ============================================================================
# Windows 單一執行檔 (.exe) 執行期環境與路徑自動配置
# ============================================================================

def is_frozen_exe() -> bool:
    """判斷目前是否以 PyInstaller 打包後的單一執行檔 (.exe) 模式執行。"""
    return getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS")


def get_app_base_dir() -> Path:
    """取得應用程式所在目錄（若為 .exe 則回傳 .exe 所在資料夾，否則回傳腳本所在資料夾）。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_bundle_resource_dir() -> Path:
    """取得 PyInstaller 解壓縮暫存資源目錄 (`sys._MEIPASS`) 或腳本目錄。"""
    if is_frozen_exe():
        return Path(getattr(sys, "_MEIPASS")).resolve()
    return get_app_base_dir()


def configure_windows_runtime() -> None:
    """
    針對 Windows 單一執行檔 (.exe) 進行環境最佳化：
      1. 設定終端機 stdout/stderr 編碼為 UTF-8，避免 UnicodeEncodeError
      2. 支援 --windowed 執行檔在命令列執行時自動附加至主控台 (AttachConsole)
      3. 啟用 Windows 高解析度螢幕 (High-DPI) 清晰字體縮放
      4. 自動尋找隨附或系統安裝的 Tesseract-OCR 與 LibreOffice 路徑
    """
    if platform.system() == "Windows":
        try:
            import ctypes
            # 若為 GUI/無主控台模式但由終端機啟動，附加至呼叫端主控台以顯示輸出
            if ctypes.windll.kernel32.AttachConsole(-1):
                sys.stdout = open("CONOUT$", "w", encoding="utf-8", errors="replace")
                sys.stderr = open("CONOUT$", "w", encoding="utf-8", errors="replace")
        except Exception:
            pass

        if hasattr(sys.stdout, "reconfigure"):
            try:
                sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass
        if hasattr(sys.stderr, "reconfigure"):
            try:
                sys.stderr.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass

        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                import ctypes
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

    # 自動偵測可攜版 (Portable) 或系統安裝的 Tesseract-OCR
    candidate_tesseract_dirs = [
        get_app_base_dir() / "Tesseract-OCR",
        get_bundle_resource_dir() / "Tesseract-OCR",
        Path(r"C:\Program Files\Tesseract-OCR"),
        Path(r"C:\Program Files (x86)\Tesseract-OCR"),
        Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Tesseract-OCR",
    ]
    for t_dir in candidate_tesseract_dirs:
        exe_path = t_dir / ("tesseract.exe" if platform.system() == "Windows" else "tesseract")
        if exe_path.exists():
            os.environ["PATH"] = str(t_dir) + os.pathsep + os.environ.get("PATH", "")
            tessdata_dir = t_dir / "tessdata"
            if tessdata_dir.exists() and "TESSDATA_PREFIX" not in os.environ:
                os.environ["TESSDATA_PREFIX"] = str(tessdata_dir)
            try:
                import pytesseract  # type: ignore
                pytesseract.pytesseract.tesseract_cmd = str(exe_path)
            except ImportError:
                pass
            break


# ============================================================================
# 轉換設定資料結構
# ============================================================================

@dataclass
class ConversionOptions:
    """控制轉檔行為的選項設定。"""
    extract_images: bool = True          # 是否匯出 Word/PPT/PDF 內嵌圖片並加上 Markdown 連結
    enable_ocr: bool = True              # 是否對掃描版 PDF 或純圖片檔啟用 OCR
    ocr_on_embedded_images: bool = False # 是否對 Word/PPT/PDF 匯出的內嵌圖片也執行 OCR
    ocr_lang: str = "chi_tra+eng"        # Tesseract OCR 語言包 (預設：繁體中文 + 英文)
    pdf_ocr_min_chars: int = 25          # PDF 單頁文字少於此字數時，自動判定為掃描頁並觸發 OCR
    pdf_render_dpi: int = 200            # 掃描版 PDF 轉圖片供 OCR 辨識的解析度 (DPI)
    engine: str = "custom"               # 'custom' | 'auto' | 'markitdown'


@dataclass
class ConversionResult:
    """單一檔案轉換結果。"""
    input_path: Path
    output_md_path: Path
    markdown_text: str
    assets_dir: Optional[Path] = None
    exported_images: List[Path] = field(default_factory=list)


# ============================================================================
# 輔助函式：Markdown 表格、編碼讀取、圖片儲存與 OCR 引擎
# ============================================================================

def clean_cell(val: object) -> str:
    """清理表格儲存格內容，避免換行與直線符號破壞 Markdown 表格結構。"""
    if val is None:
        return ""
    text = str(val).strip()
    text = text.replace("\r\n", "<br>").replace("\n", "<br>").replace("\r", "<br>")
    text = text.replace("|", "\\|")
    return text


def rows_to_markdown_table(rows: Sequence[Sequence[object]]) -> str:
    """將二維陣列轉換為標準 Markdown 管線表格 (Pipe Table)。"""
    cleaned_rows: List[List[str]] = []
    for row in rows:
        if row is None:
            continue
        cleaned = [clean_cell(c) for c in row]
        if any(cell != "" for cell in cleaned):
            cleaned_rows.append(cleaned)

    if not cleaned_rows:
        return ""

    max_cols = max(len(r) for r in cleaned_rows)
    normalized = [r + [""] * (max_cols - len(r)) for r in cleaned_rows]

    header = normalized[0]
    if all(h == "" for h in header):
        header = [f"Col_{i + 1}" for i in range(max_cols)]
        data_rows = normalized
    else:
        data_rows = normalized[1:]

    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(["---"] * max_cols) + " |",
    ]
    for r in data_rows:
        lines.append("| " + " | ".join(r) + " |")
    return "\n".join(lines)


def read_text_with_fallback(file_path: Path) -> str:
    """自動嘗試常見中文與國際編碼 (UTF-8, UTF-8-SIG, CP950/Big5, GB18030, UTF-16, Latin-1) 讀取文字檔。"""
    encodings = ["utf-8", "utf-8-sig", "cp950", "big5", "gb18030", "utf-16", "latin-1"]
    for enc in encodings:
        try:
            return file_path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    return file_path.read_text(encoding="utf-8", errors="replace")


def save_image_blob(
    blob: bytes,
    assets_dir: Path,
    filename_stem: str,
    ext: str = "png",
) -> Tuple[Path, str]:
    """
    將圖片二進位資料存入 assets_dir，並回傳 (實體檔案路徑, Markdown 相對路徑字串)。
    若圖片為 WMF/EMF/TIFF 等瀏覽器較難直接顯示的格式，嘗試透過 Pillow 轉為 PNG。
    """
    assets_dir.mkdir(parents=True, exist_ok=True)
    clean_ext = ext.lower().lstrip(".") or "png"
    if clean_ext == "jpeg":
        clean_ext = "jpg"

    if clean_ext not in {"png", "jpg", "gif", "webp", "svg"}:
        try:
            from PIL import Image  # type: ignore

            with Image.open(io.BytesIO(blob)) as im:
                out_buf = io.BytesIO()
                im.convert("RGBA" if im.mode in ("RGBA", "LA") else "RGB").save(out_buf, format="PNG")
                blob = out_buf.getvalue()
                clean_ext = "png"
        except Exception:
            pass

    img_path = assets_dir / f"{filename_stem}.{clean_ext}"
    img_path.write_bytes(blob)
    rel_md_path = f"{assets_dir.name}/{img_path.name}"
    return img_path, rel_md_path


_RAPIDOCR_INSTANCE = None


def perform_ocr_on_image_bytes(image_bytes: bytes, lang: str = "chi_tra+eng") -> str:
    """
    對圖片二進位資料進行 OCR 文字辨識。
    依序嘗試：
      1. rapidocr_onnxruntime (RapidOCR，純 Python/ONNX 離線引擎，打包進 .exe 後免安裝外部執行檔即可運作)
      2. pytesseract (Tesseract OCR，支援 chi_tra+eng)
      3. easyocr (EasyOCR)
    """
    global _RAPIDOCR_INSTANCE

    # 1. 嘗試 rapidocr_onnxruntime (最適合單一 .exe 免安裝環境)
    try:
        from rapidocr_onnxruntime import RapidOCR  # type: ignore

        if _RAPIDOCR_INSTANCE is None:
            _RAPIDOCR_INSTANCE = RapidOCR()
        result, _ = _RAPIDOCR_INSTANCE(image_bytes)
        if result:
            lines = [item[1] for item in result if len(item) >= 2 and item[1]]
            if lines:
                return "\n".join(lines).strip()
    except ImportError:
        pass
    except Exception as exc:
        logger.debug("RapidOCR 辨識略過: %s", exc)

    # 2. 嘗試 pytesseract
    try:
        from PIL import Image  # type: ignore
        import pytesseract  # type: ignore

        with Image.open(io.BytesIO(image_bytes)) as img:
            try:
                text = pytesseract.image_to_string(img, lang=lang)
            except Exception:
                text = pytesseract.image_to_string(img)
            if text and text.strip():
                return text.strip()
    except ImportError:
        pass
    except Exception as exc:
        logger.debug("pytesseract 辨識略過: %s", exc)

    # 3. 嘗試 easyocr
    try:
        import easyocr  # type: ignore
        import numpy as np  # type: ignore
        from PIL import Image  # type: ignore

        langs = ["ch_tra", "en"] if "chi_tra" in lang else ["en"]
        reader = easyocr.Reader(langs, gpu=False, verbose=False)
        with Image.open(io.BytesIO(image_bytes)) as img:
            arr = np.array(img.convert("RGB"))
            lines = reader.readtext(arr, detail=0)
            return "\n".join(str(x) for x in lines).strip()
    except ImportError:
        pass
    except Exception as exc:
        logger.debug("EasyOCR 辨識略過: %s", exc)

    return ""


def format_ocr_blockquote(ocr_text: str, title: str = "OCR 圖片文字辨識") -> str:
    """將 OCR 辨識出的文字格式化為 Markdown 引言區塊。"""
    cleaned = ocr_text.strip()
    if not cleaned:
        return ""
    quoted = "\n".join(f"> {line}" if line.strip() else ">" for line in cleaned.splitlines())
    return f"> **{title}：**\n{quoted}"


def convert_legacy_office_via_libreoffice(input_path: Path, target_ext: str) -> Optional[Path]:
    """
    若遇到舊版二進位 Office 格式 (.doc, .ppt, .xls)，優先嘗試呼叫 LibreOffice
    或 Windows 本機的 Microsoft Office COM 介面轉為新版 .docx/.pptx/.xlsx。
    """
    soffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not soffice and platform.system() == "Windows":
        for p in [
            Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
            Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
        ]:
            if p.exists():
                soffice = str(p)
                break

    if soffice:
        temp_dir = Path(tempfile.mkdtemp(prefix="office2md_"))
        try:
            cmd = [
                soffice,
                "--headless",
                "--convert-to",
                target_ext.lstrip("."),
                "--outdir",
                str(temp_dir),
                str(input_path),
            ]
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
            converted = temp_dir / f"{input_path.stem}.{target_ext.lstrip('.')}"
            if converted.exists() and converted.stat().st_size > 0:
                return converted
        except Exception as exc:
            logger.warning("LibreOffice 轉換 %s 失敗: %s", input_path.name, exc)
        shutil.rmtree(temp_dir, ignore_errors=True)

    return None


# ============================================================================
# 各格式專用轉換器 (支援 OCR 與圖片匯出)
# ============================================================================

def _render_pdf_page_to_png_bytes(file_path: Path, page_index_zero_based: int, dpi: int = 200) -> Optional[bytes]:
    """將 PDF 指定頁面渲染為 PNG 影像 bytes，供掃描檔 OCR 或頁面截圖使用。"""
    try:
        import pypdfium2 as pdfium  # type: ignore

        pdf_doc = pdfium.PdfDocument(str(file_path))
        page = pdf_doc[page_index_zero_based]
        bitmap = page.render(scale=dpi / 72.0)
        pil_image = bitmap.to_pil()
        buf = io.BytesIO()
        pil_image.save(buf, format="PNG")
        page.close()
        pdf_doc.close()
        return buf.getvalue()
    except Exception:
        pass

    try:
        from pdf2image import convert_from_path  # type: ignore

        images = convert_from_path(
            str(file_path),
            dpi=dpi,
            first_page=page_index_zero_based + 1,
            last_page=page_index_zero_based + 1,
        )
        if images:
            buf = io.BytesIO()
            images[0].save(buf, format="PNG")
            return buf.getvalue()
    except Exception:
        pass

    return None


def convert_pdf_to_md(
    file_path: Path,
    assets_dir: Path,
    options: ConversionOptions,
    exported_images: List[Path],
) -> str:
    """
    將 PDF 轉為 Markdown：
      1. 擷取各頁段落文字與表格 (pdfplumber / pypdf)
      2. 擷取頁面內嵌圖片至 assets_dir 並插入 Markdown 圖片連結
      3. 若單頁文字少於門檻 (掃描版 PDF) 且開啟 OCR，自動將該頁渲染為影像並執行 OCR
    """
    sections: List[str] = [f"# {file_path.stem}\n"]

    pypdf_reader = None
    try:
        from pypdf import PdfReader  # type: ignore

        pypdf_reader = PdfReader(str(file_path))
    except Exception:
        pass

    try:
        import pdfplumber  # type: ignore

        with pdfplumber.open(file_path) as pdf:
            total_pages = len(pdf.pages)
            for idx, page in enumerate(pdf.pages, start=1):
                page_parts: List[str] = [f"## 第 {idx} 頁 (Page {idx} / {total_pages})"]
                text = (page.extract_text() or "").strip()
                if text:
                    page_parts.append(text)

                tables = page.extract_tables() or []
                for t_idx, table in enumerate(tables, start=1):
                    md_table = rows_to_markdown_table(table)
                    if md_table:
                        page_parts.append(f"### 第 {idx} 頁表格 {t_idx}\n\n{md_table}")

                page_img_count = 0
                if options.extract_images and pypdf_reader and idx <= len(pypdf_reader.pages):
                    try:
                        pypdf_page = pypdf_reader.pages[idx - 1]
                        for img_idx, img_file_obj in enumerate(getattr(pypdf_page, "images", []), start=1):
                            blob = img_file_obj.data
                            ext = Path(img_file_obj.name).suffix.lstrip(".") or "png"
                            stem = f"pdf_p{idx:02d}_img{img_idx:02d}"
                            saved_path, rel_md = save_image_blob(blob, assets_dir, stem, ext)
                            exported_images.append(saved_path)
                            page_img_count += 1
                            alt_label = f"第 {idx} 頁圖片 {img_idx}"
                            page_parts.append(f"![{alt_label}]({rel_md})")

                            if options.ocr_on_embedded_images and options.enable_ocr:
                                ocr_res = perform_ocr_on_image_bytes(blob, lang=options.ocr_lang)
                                if ocr_res:
                                    page_parts.append(format_ocr_blockquote(ocr_res, f"{alt_label} OCR"))
                    except Exception as exc:
                        logger.debug("擷取第 %d 頁內嵌圖片略過: %s", idx, exc)

                if len(text) < options.pdf_ocr_min_chars and options.enable_ocr:
                    png_bytes = _render_pdf_page_to_png_bytes(file_path, idx - 1, dpi=options.pdf_render_dpi)
                    if png_bytes:
                        if options.extract_images and page_img_count == 0:
                            stem = f"pdf_p{idx:02d}_scan"
                            saved_path, rel_md = save_image_blob(png_bytes, assets_dir, stem, "png")
                            exported_images.append(saved_path)
                            page_parts.append(f"![第 {idx} 頁掃描影像]({rel_md})")

                        ocr_text = perform_ocr_on_image_bytes(png_bytes, lang=options.ocr_lang)
                        if ocr_text:
                            page_parts.append(format_ocr_blockquote(ocr_text, f"第 {idx} 頁掃描 OCR 辨識結果"))
                        elif not text:
                            page_parts.append("*(此頁為掃描影像，若需自動辨識文字請確認已安裝 RapidOCR 或 Tesseract OCR)*")

                sections.append("\n\n".join(page_parts))
        return "\n\n".join(sections).strip() + "\n"
    except ImportError:
        pass

    if pypdf_reader is not None:
        total_pages = len(pypdf_reader.pages)
        for idx, page in enumerate(pypdf_reader.pages, start=1):
            page_parts = [f"## 第 {idx} 頁 (Page {idx} / {total_pages})"]
            text = (page.extract_text() or "").strip()
            if text:
                page_parts.append(text)
            if len(text) < options.pdf_ocr_min_chars and options.enable_ocr:
                png_bytes = _render_pdf_page_to_png_bytes(file_path, idx - 1, dpi=options.pdf_render_dpi)
                if png_bytes:
                    ocr_text = perform_ocr_on_image_bytes(png_bytes, lang=options.ocr_lang)
                    if ocr_text:
                        page_parts.append(format_ocr_blockquote(ocr_text, f"第 {idx} 頁掃描 OCR 辨識結果"))
            sections.append("\n\n".join(page_parts))
        return "\n\n".join(sections).strip() + "\n"

    raise RuntimeError("處理 PDF 需要安裝 `pdfplumber` 或 `pypdf`：請執行 `pip install pdfplumber pypdf`")


def convert_excel_to_md(file_path: Path) -> str:
    """將 Excel (.xlsx, .xls) 或 CSV 檔案轉換為 Markdown 表格。"""
    ext = file_path.suffix.lower()
    sections: List[str] = [f"# {file_path.stem}\n"]

    if ext == ".csv":
        content = read_text_with_fallback(file_path)
        reader = csv.reader(content.splitlines())
        md_table = rows_to_markdown_table(list(reader))
        if md_table:
            sections.append(md_table)
        return "\n\n".join(sections).strip() + "\n"

    cleanup_dir: Optional[Path] = None
    target_path = file_path
    if ext == ".xls":
        converted = convert_legacy_office_via_libreoffice(file_path, "xlsx")
        if converted is not None:
            target_path = converted
            cleanup_dir = converted.parent

    try:
        try:
            import openpyxl  # type: ignore

            wb = openpyxl.load_workbook(target_path, data_only=True)
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                rows = list(ws.iter_rows(values_only=True))
                md_table = rows_to_markdown_table(rows)
                if md_table:
                    sections.append(f"## 工作表：{sheet_name}\n\n{md_table}")
                else:
                    sections.append(f"## 工作表：{sheet_name}\n\n*(此工作表無資料)*")
            return "\n\n".join(sections).strip() + "\n"
        except ImportError:
            pass

        try:
            import pandas as pd  # type: ignore

            excel_data = pd.read_excel(target_path, sheet_name=None)
            for sheet_name, df in excel_data.items():
                df = df.dropna(how="all").dropna(axis=1, how="all")
                if df.empty:
                    sections.append(f"## 工作表：{sheet_name}\n\n*(此工作表無資料)*")
                else:
                    rows = [list(df.columns)] + df.fillna("").values.tolist()
                    md_table = rows_to_markdown_table(rows)
                    sections.append(f"## 工作表：{sheet_name}\n\n{md_table}")
            return "\n\n".join(sections).strip() + "\n"
        except ImportError as exc:
            raise RuntimeError(
                "處理 Excel 需要安裝 `openpyxl` 或 `pandas`：請執行 `pip install openpyxl pandas`"
            ) from exc
    finally:
        if cleanup_dir:
            shutil.rmtree(cleanup_dir, ignore_errors=True)


def _extract_docx_element_rids(xml_element) -> List[str]:
    """從 Word XML 節點中找出所有內嵌圖片的 Relationship ID (支援 a:blip 與 v:imagedata)。"""
    rids: List[str] = []
    R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
    R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

    for elem in xml_element.iter():
        tag_local = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
        if tag_local == "blip":
            rid = elem.attrib.get(R_EMBED)
            if rid and rid not in rids:
                rids.append(rid)
        elif tag_local == "imagedata":
            rid = elem.attrib.get(R_ID)
            if rid and rid not in rids:
                rids.append(rid)
    return rids


def _format_docx_paragraph(para) -> str:
    """將 python-docx 的 Paragraph 物件轉為對應的 Markdown 語法。"""
    style_name = (para.style.name if para.style else "").lower()

    runs_text: List[str] = []
    for run in para.runs:
        t = run.text
        if not t:
            continue
        if run.bold and run.italic:
            runs_text.append(f"***{t}***")
        elif run.bold:
            runs_text.append(f"**{t}**")
        elif run.italic:
            runs_text.append(f"*{t}*")
        else:
            runs_text.append(t)

    text = "".join(runs_text).strip() or para.text.strip()
    if not text:
        return ""

    if style_name.startswith("heading"):
        parts = style_name.split()
        level = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 2
        level = max(1, min(level, 6))
        return f"{'#' * level} {para.text.strip()}"
    if style_name == "title":
        return f"# {para.text.strip()}"
    if style_name == "subtitle":
        return f"## {para.text.strip()}"

    if "list bullet" in style_name or "bullet" in style_name:
        return f"- {text}"
    if "list number" in style_name or "number" in style_name:
        return f"1. {text}"

    return text


def convert_word_to_md(
    file_path: Path,
    assets_dir: Path,
    options: ConversionOptions,
    exported_images: List[Path],
) -> str:
    """
    將 Word (.docx, .doc) 轉為 Markdown：
      - 依原始段落與表格順序輸出
      - 自動將段落或表格內的圖片匯出至 assets_dir 並插入 `![Word 圖片 X](...)`
    """
    ext = file_path.suffix.lower()
    cleanup_dir: Optional[Path] = None
    target_path = file_path

    if ext == ".doc":
        converted = convert_legacy_office_via_libreoffice(file_path, "docx")
        if converted is None:
            raise RuntimeError(
                f"檔案 `{file_path.name}` 為舊版二進位 .doc 格式。請先轉存為 .docx，或安裝 LibreOffice 以支援自動轉換。"
            )
        target_path = converted
        cleanup_dir = converted.parent

    try:
        try:
            import docx  # type: ignore
            from docx.oxml.table import CT_Tbl  # type: ignore
            from docx.oxml.text.paragraph import CT_P  # type: ignore
            from docx.table import Table  # type: ignore
            from docx.text.paragraph import Paragraph  # type: ignore
        except ImportError as exc:
            raise RuntimeError(
                "處理 Word 需要安裝 `python-docx`：請執行 `pip install python-docx`"
            ) from exc

        doc = docx.Document(str(target_path))
        blocks: List[str] = []
        img_counter = 0
        rid_to_md_link: Dict[str, Tuple[str, bytes]] = {}

        def process_rids(rids: List[str]) -> List[str]:
            nonlocal img_counter
            img_md_blocks: List[str] = []
            if not options.extract_images:
                return img_md_blocks

            for rid in rids:
                if rid in rid_to_md_link:
                    rel_md, blob = rid_to_md_link[rid]
                    img_md_blocks.append(f"![Word 圖片]({rel_md})")
                    continue

                part = doc.part.related_parts.get(rid)
                if part is None or not hasattr(part, "blob"):
                    continue

                img_counter += 1
                blob = part.blob
                part_ext = Path(str(getattr(part, "partname", "img.png"))).suffix.lstrip(".") or "png"
                stem = f"word_img_{img_counter:03d}"
                saved_path, rel_md = save_image_blob(blob, assets_dir, stem, part_ext)
                exported_images.append(saved_path)
                rid_to_md_link[rid] = (rel_md, blob)

                alt_label = f"Word 圖片 {img_counter}"
                img_md_blocks.append(f"![{alt_label}]({rel_md})")

                if options.ocr_on_embedded_images and options.enable_ocr:
                    ocr_text = perform_ocr_on_image_bytes(blob, lang=options.ocr_lang)
                    if ocr_text:
                        img_md_blocks.append(format_ocr_blockquote(ocr_text, f"{alt_label} OCR"))
            return img_md_blocks

        for child in doc.element.body.iterchildren():
            if isinstance(child, CT_P):
                para = Paragraph(child, doc)
                md_para = _format_docx_paragraph(para)
                if md_para:
                    blocks.append(md_para)
                blocks.extend(process_rids(_extract_docx_element_rids(child)))
            elif isinstance(child, CT_Tbl):
                table = Table(child, doc)
                rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]
                md_table = rows_to_markdown_table(rows)
                if md_table:
                    blocks.append(md_table)
                blocks.extend(process_rids(_extract_docx_element_rids(child)))

        if not blocks:
            return f"# {file_path.stem}\n"
        return "\n\n".join(blocks).strip() + "\n"
    finally:
        if cleanup_dir:
            shutil.rmtree(cleanup_dir, ignore_errors=True)


def _iter_ppt_shapes_recursive(shapes) -> Iterable:
    """遞迴走訪 PowerPoint 投影片內的所有 Shape (包含群組物件 GroupShape 內部的圖片與文字)。"""
    for shape in shapes:
        yield shape
        if hasattr(shape, "shapes"):
            yield from _iter_ppt_shapes_recursive(shape.shapes)


def convert_ppt_to_md(
    file_path: Path,
    assets_dir: Path,
    options: ConversionOptions,
    exported_images: List[Path],
) -> str:
    """
    將 PowerPoint (.pptx, .ppt) 轉為 Markdown：
      - 擷取各張投影片標題、文字階層、表格與講者備忘稿
      - 將投影片內的圖片匯出至 assets_dir 並插入 Markdown 圖片連結
    """
    ext = file_path.suffix.lower()
    cleanup_dir: Optional[Path] = None
    target_path = file_path

    if ext == ".ppt":
        converted = convert_legacy_office_via_libreoffice(file_path, "pptx")
        if converted is None:
            raise RuntimeError(
                f"檔案 `{file_path.name}` 為舊版二進位 .ppt 格式。請先轉存為 .pptx，或安裝 LibreOffice 以支援自動轉換。"
            )
        target_path = converted
        cleanup_dir = converted.parent

    try:
        try:
            from pptx import Presentation  # type: ignore
        except ImportError as exc:
            raise RuntimeError(
                "處理 PowerPoint 需要安裝 `python-pptx`：請執行 `pip install python-pptx`"
            ) from exc

        prs = Presentation(str(target_path))
        sections: List[str] = [f"# {file_path.stem}\n"]

        for idx, slide in enumerate(prs.slides, start=1):
            slide_title = ""
            if slide.shapes.title and slide.shapes.title.text.strip():
                slide_title = slide.shapes.title.text.strip()

            header = f"## 投影片 {idx}：{slide_title}" if slide_title else f"## 投影片 {idx}"
            slide_parts: List[str] = [header]
            slide_img_idx = 0

            for shape in _iter_ppt_shapes_recursive(slide.shapes):
                if slide.shapes.title and shape == slide.shapes.title:
                    continue

                if getattr(shape, "has_text_frame", False):
                    lines: List[str] = []
                    for para in shape.text_frame.paragraphs:
                        text = para.text.strip()
                        if not text:
                            continue
                        indent = "  " * max(0, int(para.level or 0))
                        lines.append(f"{indent}- {text}")
                    if lines:
                        slide_parts.append("\n".join(lines))

                if getattr(shape, "has_table", False):
                    rows = [
                        [cell.text.strip() for cell in row.cells]
                        for row in shape.table.rows
                    ]
                    md_table = rows_to_markdown_table(rows)
                    if md_table:
                        slide_parts.append(md_table)

                if options.extract_images and hasattr(shape, "image"):
                    try:
                        img_obj = shape.image
                        blob = img_obj.blob
                        img_ext = getattr(img_obj, "ext", "png") or "png"
                        slide_img_idx += 1
                        stem = f"slide_{idx:02d}_img_{slide_img_idx:02d}"
                        saved_path, rel_md = save_image_blob(blob, assets_dir, stem, img_ext)
                        exported_images.append(saved_path)

                        alt_label = f"投影片 {idx} 圖片 {slide_img_idx}"
                        slide_parts.append(f"![{alt_label}]({rel_md})")

                        if options.ocr_on_embedded_images and options.enable_ocr:
                            ocr_text = perform_ocr_on_image_bytes(blob, lang=options.ocr_lang)
                            if ocr_text:
                                slide_parts.append(format_ocr_blockquote(ocr_text, f"{alt_label} OCR"))
                    except Exception as exc:
                        logger.debug("投影片 %d 圖片匯出略過: %s", idx, exc)

            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                notes = slide.notes_slide.notes_text_frame.text.strip()
                if notes:
                    quoted_notes = "\n".join(f"> {line}" for line in notes.splitlines())
                    slide_parts.append(f"**講者備忘稿：**\n{quoted_notes}")

            sections.append("\n\n".join(slide_parts))

        return "\n\n".join(sections).strip() + "\n"
    finally:
        if cleanup_dir:
            shutil.rmtree(cleanup_dir, ignore_errors=True)


def convert_image_file_to_md(
    file_path: Path,
    assets_dir: Path,
    options: ConversionOptions,
    exported_images: List[Path],
) -> str:
    """將單獨的掃描圖片檔 (.png, .jpg, .tiff, .bmp, .webp) 複製到 assets 並透過 OCR 轉為 Markdown。"""
    blob = file_path.read_bytes()
    ext = file_path.suffix.lstrip(".") or "png"
    parts: List[str] = [f"# {file_path.stem}"]

    if options.extract_images:
        saved_path, rel_md = save_image_blob(blob, assets_dir, file_path.stem, ext)
        exported_images.append(saved_path)
        parts.append(f"![{file_path.name}]({rel_md})")

    if options.enable_ocr:
        ocr_text = perform_ocr_on_image_bytes(blob, lang=options.ocr_lang)
        if ocr_text:
            parts.append(f"## OCR 辨識文字內容\n\n{ocr_text}")
        else:
            parts.append("*(未偵測到文字，或系統尚未安裝對應的 OCR 套件如 `rapidocr_onnxruntime` / `pytesseract`)*")

    return "\n\n".join(parts).strip() + "\n"


def convert_text_to_md(file_path: Path) -> str:
    """將純文字檔 (.txt, .log, .json, .md) 轉為結構化 Markdown。"""
    ext = file_path.suffix.lower()
    raw_text = read_text_with_fallback(file_path).strip()

    if ext == ".md":
        return raw_text + "\n"

    if ext == ".json":
        try:
            parsed = json.loads(raw_text)
            pretty = json.dumps(parsed, ensure_ascii=False, indent=2)
            return f"# {file_path.stem}\n\n```json\n{pretty}\n```\n"
        except Exception:
            pass

    return f"# {file_path.stem}\n\n{raw_text}\n"


# ============================================================================
# 整合轉換類別
# ============================================================================

class DocumentToMarkdownConverter:
    """統一管理多格式檔案轉 Markdown、圖片匯出與 OCR 的轉換器。"""

    SUPPORTED_EXTENSIONS = {
        ".pdf": "pdf",
        ".xlsx": "excel",
        ".xls": "excel",
        ".csv": "excel",
        ".docx": "word",
        ".doc": "word",
        ".pptx": "ppt",
        ".ppt": "ppt",
        ".txt": "text",
        ".log": "text",
        ".json": "text",
        ".md": "text",
        ".png": "image",
        ".jpg": "image",
        ".jpeg": "image",
        ".bmp": "image",
        ".tiff": "image",
        ".tif": "image",
        ".webp": "image",
    }

    def __init__(self, options: Optional[ConversionOptions] = None) -> None:
        self.options = options or ConversionOptions()

    @classmethod
    def is_supported(cls, file_path: Path) -> bool:
        return file_path.suffix.lower() in cls.SUPPORTED_EXTENSIONS

    def convert_file_detailed(
        self,
        input_path: str | Path,
        output_path: Optional[str | Path] = None,
    ) -> ConversionResult:
        """
        轉換單一檔案並儲存 `.md` 與 `<檔名>_assets/` 圖片資料夾，回傳完整 ConversionResult。
        """
        in_path = Path(input_path).expanduser().resolve()
        if not in_path.exists():
            raise FileNotFoundError(f"找不到指定的檔案：{in_path}")

        ext = in_path.suffix.lower()
        if ext not in self.SUPPORTED_EXTENSIONS:
            supported_list = ", ".join(sorted(self.SUPPORTED_EXTENSIONS.keys()))
            raise ValueError(f"不支援的副檔名 `{ext}`。支援格式：{supported_list}")

        if output_path is None:
            out_md_path = in_path.with_suffix(".md")
        else:
            out_md_path = Path(output_path).expanduser().resolve()
            if out_md_path.is_dir():
                out_md_path = out_md_path / f"{in_path.stem}.md"

        if in_path == out_md_path:
            out_md_path = in_path.with_name(f"{in_path.stem}_converted.md")

        out_md_path.parent.mkdir(parents=True, exist_ok=True)
        assets_dir = out_md_path.parent / f"{out_md_path.stem}_assets"
        exported_images: List[Path] = []

        markdown_text = ""
        if self.options.engine in {"markitdown", "auto"} and not self.options.extract_images:
            try:
                from markitdown import MarkItDown  # type: ignore

                md = MarkItDown()
                res = md.convert(str(in_path))
                markdown_text = (res.text_content or "").strip() + "\n"
            except ImportError:
                if self.options.engine == "markitdown":
                    raise RuntimeError("尚未安裝 `markitdown`，請執行 `pip install markitdown`")

        if not markdown_text:
            doc_type = self.SUPPORTED_EXTENSIONS[ext]
            if doc_type == "pdf":
                markdown_text = convert_pdf_to_md(in_path, assets_dir, self.options, exported_images)
            elif doc_type == "excel":
                markdown_text = convert_excel_to_md(in_path)
            elif doc_type == "word":
                markdown_text = convert_word_to_md(in_path, assets_dir, self.options, exported_images)
            elif doc_type == "ppt":
                markdown_text = convert_ppt_to_md(in_path, assets_dir, self.options, exported_images)
            elif doc_type == "image":
                markdown_text = convert_image_file_to_md(in_path, assets_dir, self.options, exported_images)
            else:
                markdown_text = convert_text_to_md(in_path)

        out_md_path.write_text(markdown_text, encoding="utf-8")
        logger.info(
            "已轉換：%s -> %s (匯出圖片數: %d)",
            in_path.name,
            out_md_path,
            len(exported_images),
        )
        return ConversionResult(
            input_path=in_path,
            output_md_path=out_md_path,
            markdown_text=markdown_text,
            assets_dir=assets_dir if exported_images else None,
            exported_images=exported_images,
        )

    def convert_file(
        self,
        input_path: str | Path,
        output_path: Optional[str | Path] = None,
    ) -> Path:
        """向下相容簡化介面：僅回傳輸出的 Markdown 檔案路徑。"""
        return self.convert_file_detailed(input_path, output_path).output_md_path

    def convert_directory(
        self,
        input_dir: str | Path,
        output_dir: Optional[str | Path] = None,
        recursive: bool = False,
    ) -> List[ConversionResult]:
        """批次轉換資料夾內所有支援的文件格式。"""
        in_dir = Path(input_dir).expanduser().resolve()
        if not in_dir.is_dir():
            raise NotADirectoryError(f"指定的路徑不是資料夾：{in_dir}")

        out_base = Path(output_dir).expanduser().resolve() if output_dir else in_dir
        pattern = "**/*" if recursive else "*"
        results: List[ConversionResult] = []

        for item in sorted(in_dir.glob(pattern)):
            if not item.is_file() or not self.is_supported(item):
                continue
            if item.suffix.lower() == ".md":
                continue
            if any(part.endswith("_assets") for part in item.parts):
                continue

            rel_path = item.relative_to(in_dir).with_suffix(".md")
            target_out = out_base / rel_path
            try:
                res = self.convert_file_detailed(item, target_out)
                results.append(res)
            except Exception as exc:
                logger.error("轉換 %s 時發生錯誤：%s", item, exc)

        return results


# ============================================================================
# 桌面視窗拖曳介面 (Desktop GUI with Drag-and-Drop & One-Click Web UI)
# ============================================================================

def launch_desktop_gui(initial_paths: Optional[Sequence[Path]] = None) -> None:
    """
    啟動具備檔案拖曳 (Drag & Drop) 功能的桌面視窗介面。
    支援：
      - 直接將檔案從 Windows 檔案總管拖入視窗
      - 將檔案直接拖放到 .exe 圖示上啟動時自動載入清單
      - 視窗內一鍵啟動本地端網頁上傳介面 (Web UI)
    """
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, scrolledtext, ttk
    except ImportError as exc:
        raise RuntimeError("目前的環境未包含 Tkinter 視窗模組，請改用 `--web` 啟動網頁介面。") from exc

    has_dnd = False
    try:
        from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore

        root = TkinterDnD.Tk()
        has_dnd = True
    except Exception:
        root = tk.Tk()

    root.title("Document to Markdown Pro - 多格式文件轉 Markdown 工具 (含 OCR 與圖片匯出)")
    root.geometry("1020x700")
    root.minsize(840, 580)

    queued_files: List[Path] = []
    web_server_started = False

    top_frame = ttk.LabelFrame(root, text="轉換設定", padding=10)
    top_frame.pack(fill=tk.X, padx=12, pady=(10, 6))

    var_extract_img = tk.BooleanVar(value=True)
    var_enable_ocr = tk.BooleanVar(value=True)
    var_ocr_embedded = tk.BooleanVar(value=False)
    var_ocr_lang = tk.StringVar(value="chi_tra+eng")
    var_output_dir = tk.StringVar(value="")

    ttk.Checkbutton(
        top_frame,
        text="匯出 Word/PPT/PDF 內嵌圖片並加入 Markdown 連結",
        variable=var_extract_img,
    ).grid(row=0, column=0, sticky=tk.W, padx=6, pady=3)

    ttk.Checkbutton(
        top_frame,
        text="啟用掃描版 PDF / 影像 OCR 文字辨識",
        variable=var_enable_ocr,
    ).grid(row=0, column=1, sticky=tk.W, padx=6, pady=3)

    ttk.Checkbutton(
        top_frame,
        text="對 Word/PPT 內嵌圖片也執行 OCR",
        variable=var_ocr_embedded,
    ).grid(row=1, column=0, sticky=tk.W, padx=6, pady=3)

    lang_frame = ttk.Frame(top_frame)
    lang_frame.grid(row=1, column=1, sticky=tk.W, padx=6, pady=3)
    ttk.Label(lang_frame, text="OCR 語言：").pack(side=tk.LEFT)
    ttk.Entry(lang_frame, textvariable=var_ocr_lang, width=16).pack(side=tk.LEFT, padx=4)
    ttk.Label(lang_frame, text="(RapidOCR 自動中英辨識 / Tesseract 語言包)").pack(side=tk.LEFT)

    out_frame = ttk.Frame(top_frame)
    out_frame.grid(row=2, column=0, columnspan=2, sticky=tk.EW, padx=6, pady=(6, 2))
    ttk.Label(out_frame, text="輸出資料夾 (留空則輸出於原檔案同目錄)：").pack(side=tk.LEFT)
    ttk.Entry(out_frame, textvariable=var_output_dir, width=48).pack(side=tk.LEFT, padx=6, fill=tk.X, expand=True)

    def choose_output_dir() -> None:
        folder = filedialog.askdirectory(title="選擇輸出資料夾")
        if folder:
            var_output_dir.set(folder)

    ttk.Button(out_frame, text="瀏覽...", command=choose_output_dir).pack(side=tk.LEFT)

    paned = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
    paned.pack(fill=tk.BOTH, expand=True, padx=12, pady=6)

    left_frame = ttk.Frame(paned, padding=4)
    right_frame = ttk.Frame(paned, padding=4)
    paned.add(left_frame, weight=1)
    paned.add(right_frame, weight=1)

    dnd_hint = (
        "📥 將 PDF / Excel / Word / PPT / TXT / 圖片 拖曳至此處\n(已啟用原生視窗拖放，亦支援直接將檔案拖放到 .exe 圖示上)"
        if has_dnd
        else "📥 點擊下方「加入檔案」或「加入資料夾」\n(打包前安裝 tkinterdnd2 即可啟用視窗拖放功能)"
    )
    drop_label = tk.Label(
        left_frame,
        text=dnd_hint,
        relief=tk.GROOVE,
        bd=2,
        bg="#f0f4f8",
        fg="#1e293b",
        font=("Microsoft JhengHei", 10, "bold"),
        height=4,
        padx=10,
        pady=10,
    )
    drop_label.pack(fill=tk.X, pady=(0, 6))

    list_frame = ttk.LabelFrame(left_frame, text="待轉換檔案清單", padding=4)
    list_frame.pack(fill=tk.BOTH, expand=True)

    file_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, font=("Consolas", 10))
    file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    list_scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=file_listbox.yview)
    list_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    file_listbox.config(yscrollcommand=list_scroll.set)

    def add_paths(paths: Iterable[Path]) -> None:
        for p in paths:
            p = p.resolve()
            if p.is_dir():
                for sub in sorted(p.rglob("*")):
                    if sub.is_file() and DocumentToMarkdownConverter.is_supported(sub) and sub.suffix.lower() != ".md":
                        if sub not in queued_files:
                            queued_files.append(sub)
                            file_listbox.insert(tk.END, str(sub))
            elif p.is_file() and DocumentToMarkdownConverter.is_supported(p):
                if p not in queued_files:
                    queued_files.append(p)
                    file_listbox.insert(tk.END, str(p))

    if has_dnd:
        def on_drop(event) -> None:
            raw_files = root.tk.splitlist(event.data)
            add_paths(Path(f) for f in raw_files)

        drop_label.drop_target_register(DND_FILES)
        drop_label.dnd_bind("<<Drop>>", on_drop)
        file_listbox.drop_target_register(DND_FILES)
        file_listbox.dnd_bind("<<Drop>>", on_drop)

    btn_bar = ttk.Frame(left_frame)
    btn_bar.pack(fill=tk.X, pady=6)

    def browse_files() -> None:
        files = filedialog.askopenfilenames(
            title="選擇要轉換的檔案",
            filetypes=[
                ("支援的文件與圖片", "*.pdf *.xlsx *.xls *.csv *.docx *.doc *.pptx *.ppt *.txt *.json *.png *.jpg *.jpeg *.tiff *.webp"),
                ("所有檔案", "*.*"),
            ],
        )
        if files:
            add_paths(Path(f) for f in files)

    def browse_folder() -> None:
        folder = filedialog.askdirectory(title="選擇要批次轉換的資料夾")
        if folder:
            add_paths([Path(folder)])

    def clear_list() -> None:
        queued_files.clear()
        file_listbox.delete(0, tk.END)

    def start_web_ui_from_gui() -> None:
        nonlocal web_server_started
        url = "http://127.0.0.1:8000"
        if not web_server_started:
            web_server_started = True
            threading.Thread(
                target=lambda: launch_web_server(host="127.0.0.1", port=8000, open_browser=False),
                daemon=True,
            ).start()
        webbrowser.open(url)

    ttk.Button(btn_bar, text="➕ 加入檔案", command=browse_files).pack(side=tk.LEFT, padx=3)
    ttk.Button(btn_bar, text="📁 加入資料夾", command=browse_folder).pack(side=tk.LEFT, padx=3)
    ttk.Button(btn_bar, text="🗑️ 清空清單", command=clear_list).pack(side=tk.LEFT, padx=3)
    ttk.Button(btn_bar, text="🌐 開啟網頁上傳版", command=start_web_ui_from_gui).pack(side=tk.RIGHT, padx=3)

    preview_group = ttk.LabelFrame(right_frame, text="Markdown 輸出預覽與執行紀錄", padding=4)
    preview_group.pack(fill=tk.BOTH, expand=True)
    preview_box = scrolledtext.ScrolledText(preview_group, wrap=tk.WORD, font=("Consolas", 10))
    preview_box.pack(fill=tk.BOTH, expand=True)

    bottom_frame = ttk.Frame(root, padding=(12, 6))
    bottom_frame.pack(fill=tk.X)

    progress_var = tk.DoubleVar(value=0.0)
    progress_bar = ttk.Progressbar(bottom_frame, variable=progress_var, maximum=100.0)
    progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

    convert_btn = ttk.Button(bottom_frame, text="🚀 開始轉換為 Markdown")
    convert_btn.pack(side=tk.RIGHT)

    def run_conversion() -> None:
        if not queued_files:
            messagebox.showwarning("提示", "請先加入或拖曳至少一個檔案！")
            return

        convert_btn.config(state=tk.DISABLED)
        preview_box.delete("1.0", tk.END)
        opts = ConversionOptions(
            extract_images=var_extract_img.get(),
            enable_ocr=var_enable_ocr.get(),
            ocr_on_embedded_images=var_ocr_embedded.get(),
            ocr_lang=var_ocr_lang.get().strip() or "chi_tra+eng",
        )
        converter = DocumentToMarkdownConverter(options=opts)
        out_dir_str = var_output_dir.get().strip()
        out_dir = Path(out_dir_str) if out_dir_str else None

        def worker() -> None:
            total = len(queued_files)
            last_md = ""
            success_count = 0
            for idx, fpath in enumerate(list(queued_files), start=1):
                try:
                    res = converter.convert_file_detailed(fpath, out_dir)
                    last_md = res.markdown_text
                    success_count += 1
                    msg = f"[完成 {idx}/{total}] {fpath.name} -> {res.output_md_path} (匯出圖片: {len(res.exported_images)} 張)\n"
                except Exception as exc:
                    msg = f"[失敗 {idx}/{total}] {fpath.name} 轉換錯誤: {exc}\n"

                root.after(0, lambda m=msg, p=(idx / total) * 100: (
                    preview_box.insert(tk.END, m),
                    preview_box.see(tk.END),
                    progress_var.set(p),
                ))

            def finish() -> None:
                if last_md:
                    preview_box.insert(tk.END, "\n" + "=" * 60 + "\n【最新轉換檔案 Markdown 預覽】\n" + "=" * 60 + "\n\n" + last_md)
                convert_btn.config(state=tk.NORMAL)
                messagebox.showinfo("轉換完成", f"已完成 {success_count} / {total} 個檔案轉換！")

            root.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    convert_btn.config(command=run_conversion)

    if initial_paths:
        add_paths(initial_paths)

    root.mainloop()


# ============================================================================
# 網頁拖曳上傳工具 (Web Upload UI & API Server - 相容 Python 3.8 ~ 3.13+)
# ============================================================================

WEB_UI_HTML = """<!DOCTYPE html>
<html lang="zh-Hant">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>多格式文件轉 Markdown 網頁工具 (含 OCR 與圖片擷取)</title>
  <style>
    :root {
      --bg: #f8fafc;
      --card: #ffffff;
      --primary: #2563eb;
      --primary-hover: #1d4ed8;
      --text: #0f172a;
      --muted: #64748b;
      --border: #e2e8f0;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Microsoft JhengHei", "PingFang TC", sans-serif;
      background: var(--bg);
      color: var(--text);
      line-height: 1.6;
    }
    header {
      background: #0f172a;
      color: #fff;
      padding: 20px 32px;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    header h1 { margin: 0; font-size: 20px; font-weight: 700; }
    header span { font-size: 13px; color: #94a3b8; }
    .container {
      max-width: 1200px;
      margin: 24px auto;
      padding: 0 20px;
      display: grid;
      grid-template-columns: 400px 1fr;
      gap: 24px;
    }
    @media (max-width: 900px) {
      .container { grid-template-columns: 1fr; }
    }
    .card {
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 20px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }
    .drop-zone {
      border: 2px dashed #94a3b8;
      border-radius: 10px;
      padding: 32px 16px;
      text-align: center;
      cursor: pointer;
      background: #f1f5f9;
      transition: all 0.2s ease;
    }
    .drop-zone.dragover {
      border-color: var(--primary);
      background: #eff6ff;
    }
    .drop-zone h3 { margin: 0 0 8px; font-size: 16px; color: var(--primary); }
    .drop-zone p { margin: 0; font-size: 13px; color: var(--muted); }
    .options { margin-top: 18px; display: flex; flex-direction: column; gap: 10px; font-size: 14px; }
    .options label { display: flex; align-items: center; gap: 8px; cursor: pointer; }
    .options input[type="text"] {
      padding: 6px 10px;
      border: 1px solid var(--border);
      border-radius: 6px;
      width: 140px;
    }
    .btn {
      margin-top: 16px;
      width: 100%;
      padding: 11px 16px;
      background: var(--primary);
      color: #fff;
      font-weight: 600;
      font-size: 15px;
      border: none;
      border-radius: 8px;
      cursor: pointer;
      transition: background 0.15s;
    }
    .btn:hover { background: var(--primary-hover); }
    .btn:disabled { background: #94a3b8; cursor: not-allowed; }
    .btn-secondary {
      background: #0f172a;
      margin-top: 0;
      width: auto;
      padding: 7px 14px;
      font-size: 13px;
    }
    .file-list {
      margin-top: 14px;
      font-size: 13px;
      color: var(--muted);
      max-height: 150px;
      overflow-y: auto;
    }
    .toolbar {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 12px;
      gap: 8px;
      flex-wrap: wrap;
    }
    textarea {
      width: 100%;
      height: 460px;
      font-family: "JetBrains Mono", Consolas, monospace;
      font-size: 13px;
      padding: 14px;
      border: 1px solid var(--border);
      border-radius: 8px;
      resize: vertical;
      background: #0f172a;
      color: #f8fafc;
    }
    .gallery {
      margin-top: 16px;
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(140px, 1fr));
      gap: 12px;
    }
    .gallery-item {
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 8px;
      text-align: center;
      font-size: 12px;
      background: #fff;
    }
    .gallery-item img {
      max-width: 100%;
      max-height: 100px;
      object-fit: contain;
      border-radius: 4px;
    }
  </style>
</head>
<body>
  <header>
    <h1>📄 多格式文件轉 Markdown 工具 (PDF / Excel / Word / PPT / TXT / 影像 OCR)</h1>
    <span>支援掃描檔 OCR 辨識 ‧ Word/PPT 內嵌圖片匯出 ‧ ZIP 打包下載</span>
  </header>

  <div class="container">
    <div class="card">
      <div id="dropZone" class="drop-zone">
        <h3>📥 拖曳檔案至此，或點擊選擇檔案</h3>
        <p>支援 .pdf, .xlsx, .xls, .csv, .docx, .doc, .pptx, .ppt, .txt, .png, .jpg</p>
        <input type="file" id="fileInput" multiple style="display:none" />
      </div>

      <div id="fileList" class="file-list">尚未選擇檔案</div>

      <div class="options">
        <label>
          <input type="checkbox" id="extractImages" checked />
          匯出 Word / PPT / PDF 內嵌圖片並插入 Markdown 連結
        </label>
        <label>
          <input type="checkbox" id="enableOcr" checked />
          啟用掃描版 PDF 與圖片檔 OCR 文字辨識
        </label>
        <label>
          <input type="checkbox" id="ocrEmbedded" />
          對 Word / PPT 匯出的內嵌圖片也執行 OCR
        </label>
        <label>
          OCR 語言設定：
          <input type="text" id="ocrLang" value="chi_tra+eng" />
        </label>
      </div>

      <button id="convertBtn" class="btn">🚀 開始轉換為 Markdown</button>
    </div>

    <div class="card">
      <div class="toolbar">
        <strong>Markdown 轉換結果預覽</strong>
        <div style="display:flex; gap:8px;">
          <button id="copyBtn" class="btn btn-secondary">📋 複製 Markdown</button>
          <button id="downloadMdBtn" class="btn btn-secondary" disabled>⬇️ 下載 .md</button>
          <button id="downloadZipBtn" class="btn btn-secondary" disabled>📦 下載完整 .zip (含圖片)</button>
        </div>
      </div>
      <textarea id="mdOutput" placeholder="轉換後的 Markdown 內容將顯示於此..."></textarea>
      <div id="imageGallery" class="gallery"></div>
    </div>
  </div>

  <script>
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const fileList = document.getElementById('fileList');
    const convertBtn = document.getElementById('convertBtn');
    const mdOutput = document.getElementById('mdOutput');
    const copyBtn = document.getElementById('copyBtn');
    const downloadMdBtn = document.getElementById('downloadMdBtn');
    const downloadZipBtn = document.getElementById('downloadZipBtn');
    const imageGallery = document.getElementById('imageGallery');

    let selectedFiles = [];
    let lastZipBase64 = null;
    let lastMdFilename = 'converted.md';

    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('dragover'); });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
    dropZone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropZone.classList.remove('dragover');
      selectedFiles = Array.from(e.dataTransfer.files);
      renderFileList();
    });
    fileInput.addEventListener('change', (e) => {
      selectedFiles = Array.from(e.target.files);
      renderFileList();
    });

    function renderFileList() {
      if (!selectedFiles.length) {
        fileList.textContent = '尚未選擇檔案';
        return;
      }
      fileList.innerHTML = selectedFiles.map(f => `<div>• ${f.name} (${(f.size/1024).toFixed(1)} KB)</div>`).join('');
    }

    convertBtn.addEventListener('click', async () => {
      if (!selectedFiles.length) {
        alert('請先拖曳或選擇至少一個檔案！');
        return;
      }
      convertBtn.disabled = true;
      convertBtn.textContent = '⏳ 轉換與辨識處理中...';
      mdOutput.value = '';
      imageGallery.innerHTML = '';

      const formData = new FormData();
      selectedFiles.forEach(f => formData.append('files', f));
      formData.append('extract_images', document.getElementById('extractImages').checked ? '1' : '0');
      formData.append('enable_ocr', document.getElementById('enableOcr').checked ? '1' : '0');
      formData.append('ocr_embedded', document.getElementById('ocrEmbedded').checked ? '1' : '0');
      formData.append('ocr_lang', document.getElementById('ocrLang').value || 'chi_tra+eng');

      try {
        const resp = await fetch('/api/convert', { method: 'POST', body: formData });
        const data = await resp.json();
        if (!resp.ok || data.error) {
          throw new Error(data.error || '轉換失敗');
        }
        mdOutput.value = data.markdown || '';
        lastZipBase64 = data.zip_base64 || null;
        lastMdFilename = data.md_filename || 'converted.md';
        downloadMdBtn.disabled = false;
        downloadZipBtn.disabled = !lastZipBase64;

        if (data.images && data.images.length) {
          imageGallery.innerHTML = data.images.map(img => `
            <div class="gallery-item">
              <img src="data:${img.mime};base64,${img.base64}" alt="${img.name}" />
              <div style="margin-top:4px; word-break:break-all;">${img.rel_path}</div>
            </div>
          `).join('');
        }
      } catch (err) {
        mdOutput.value = '發生錯誤：' + err.message;
      } finally {
        convertBtn.disabled = false;
        convertBtn.textContent = '🚀 開始轉換為 Markdown';
      }
    });

    copyBtn.addEventListener('click', () => {
      navigator.clipboard.writeText(mdOutput.value);
      copyBtn.textContent = '✅ 已複製！';
      setTimeout(() => copyBtn.textContent = '📋 複製 Markdown', 1500);
    });

    downloadMdBtn.addEventListener('click', () => {
      const blob = new Blob([mdOutput.value], { type: 'text/markdown;charset=utf-8' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = lastMdFilename;
      a.click();
    });

    downloadZipBtn.addEventListener('click', () => {
      if (!lastZipBase64) return;
      const bin = atob(lastZipBase64);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      const blob = new Blob([bytes], { type: 'application/zip' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = lastMdFilename.replace(/\\.md$/i, '') + '_bundle.zip';
      a.click();
    });
  </script>
</body>
</html>
"""


def parse_multipart_form_data(content_type: str, body: bytes) -> Tuple[Dict[str, str], List[Tuple[str, bytes]]]:
    """
    使用 Python 標準庫 `email.parser.BytesParser` 解析 multipart/form-data，
    無須依賴在 Python 3.13 移除的 `cgi` 模組。
    """
    raw_mime = b"Content-Type: " + content_type.encode("utf-8") + b"\r\nMIME-Version: 1.0\r\n\r\n" + body
    msg = BytesParser(policy=email_default_policy).parsebytes(raw_mime)

    fields: Dict[str, str] = {}
    files: List[Tuple[str, bytes]] = []

    if msg.is_multipart():
        for part in msg.iter_parts():
            name = part.get_param("name", header="content-disposition")
            filename = part.get_filename()
            payload = part.get_payload(decode=True) or b""
            if filename:
                files.append((filename, payload))
            elif name:
                fields[str(name)] = payload.decode("utf-8", errors="replace")
    return fields, files


class WebConverterHandler(BaseHTTPRequestHandler):
    """處理網頁上傳轉檔請求的 HTTP Handler。"""

    def log_message(self, format: str, *args) -> None:
        logger.info("[Web] " + format, *args)

    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path in {"/", "/index.html"}:
            body = WEB_UI_HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        if parsed.path == "/api/health":
            body = json.dumps({"status": "ok"}).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        self.send_error(404, "Not Found")

    def do_POST(self) -> None:
        if self.path != "/api/convert":
            self.send_error(404, "Not Found")
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            content_type = self.headers.get("Content-Type", "")
            body = self.rfile.read(content_length)

            fields, uploaded_files = parse_multipart_form_data(content_type, body)
            if not uploaded_files:
                raise ValueError("未收到任何上傳檔案。")

            opts = ConversionOptions(
                extract_images=fields.get("extract_images", "1") == "1",
                enable_ocr=fields.get("enable_ocr", "1") == "1",
                ocr_on_embedded_images=fields.get("ocr_embedded", "0") == "1",
                ocr_lang=fields.get("ocr_lang", "chi_tra+eng").strip() or "chi_tra+eng",
            )
            converter = DocumentToMarkdownConverter(options=opts)

            work_dir = Path(tempfile.mkdtemp(prefix="web_doc2md_"))
            try:
                combined_md_parts: List[str] = []
                image_previews: List[Dict[str, str]] = []
                zip_buf = io.BytesIO()
                first_stem = Path(uploaded_files[0][0]).stem or "converted"

                with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                    for orig_filename, file_bytes in uploaded_files:
                        safe_name = Path(orig_filename).name
                        in_file = work_dir / safe_name
                        in_file.write_bytes(file_bytes)

                        res = converter.convert_file_detailed(in_file, work_dir / f"{in_file.stem}.md")
                        combined_md_parts.append(res.markdown_text)
                        zf.write(res.output_md_path, arcname=res.output_md_path.name)

                        for img_path in res.exported_images:
                            rel_arc = f"{res.output_md_path.stem}_assets/{img_path.name}"
                            zf.write(img_path, arcname=rel_arc)
                            if len(image_previews) < 24:
                                ext = img_path.suffix.lower().lstrip(".")
                                mime = f"image/{'jpeg' if ext == 'jpg' else ext}"
                                image_previews.append({
                                    "name": img_path.name,
                                    "rel_path": rel_arc,
                                    "mime": mime,
                                    "base64": base64.b64encode(img_path.read_bytes()).decode("ascii"),
                                })

                payload = {
                    "markdown": "\n\n---\n\n".join(combined_md_parts),
                    "md_filename": f"{first_stem}.md",
                    "zip_base64": base64.b64encode(zip_buf.getvalue()).decode("ascii"),
                    "images": image_previews,
                }
                resp_bytes = json.dumps(payload, ensure_ascii=False).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(resp_bytes)))
                self.end_headers()
                self.wfile.write(resp_bytes)
            finally:
                shutil.rmtree(work_dir, ignore_errors=True)

        except Exception as exc:
            logger.exception("網頁轉檔發生錯誤")
            err_bytes = json.dumps({"error": str(exc)}, ensure_ascii=False).encode("utf-8")
            self.send_response(400)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(err_bytes)))
            self.end_headers()
            self.wfile.write(err_bytes)


def launch_web_server(host: str = "127.0.0.1", port: int = 8000, open_browser: bool = True) -> None:
    """啟動本地端網頁上傳轉檔伺服器。"""
    server = ThreadingHTTPServer((host, port), WebConverterHandler)
    url = f"http://{host}:{port}"
    logger.info("網頁上傳轉檔工具已啟動：%s (按 Ctrl+C 結束)", url)
    if open_browser:
        try:
            webbrowser.open(url)
        except Exception:
            pass
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        logger.info("正在關閉網頁伺服器...")
    finally:
        server.server_close()


# ============================================================================
# Windows 單一執行檔 (.exe) 自動化打包器 (PyInstaller Builder)
# ============================================================================

WINDOWS_BUILD_BAT_TEMPLATE = """@echo off
chcp 65001 >nul
echo ====================================================================
echo   多格式文件轉 Markdown 工具 - Windows 打包腳本 (One-Folder / One-Dir)
echo ====================================================================
echo.

echo [1/3] 正在檢查並安裝必要 Python 套件與 PyInstaller...
pip install --upgrade pip
pip install pyinstaller pdfplumber pypdf pypdfium2 pillow openpyxl pandas python-docx python-pptx tkinterdnd2 pytesseract rapidocr_onnxruntime markitdown

echo.
echo [2/3] 正在透過 PyInstaller 封裝資料夾執行檔 (dist\\DocumentToMarkdownPro\\)...
python "%~dp0convert_to_markdown.py" --build-exe --onedir

echo.
echo [3/3] 打包完成！請至 dist\\DocumentToMarkdownPro\\DocumentToMarkdownPro.exe 啟動程式。
pause
"""


def build_standalone_exe(
    exe_name: str = "DocumentToMarkdownPro",
    console: bool = False,
    bundle_tesseract_dir: Optional[str] = None,
    onedir: bool = True,
) -> int:
    """
    呼叫 PyInstaller 將本程式與所有相依套件（含 tkinterdnd2 拖放 DLL、pypdfium2、
    RapidOCR 離線中英文字庫模型等）封裝為免安裝 Python 的執行檔。
    
    支援：
      - onedir=True  : 打包為資料夾目錄模式 (--onedir / --onefolder，啟動速度極快，免每次解壓縮)
      - onedir=False : 打包為單一檔案模式 (--onefile，單一 .exe)
    """
    script_path = Path(__file__).resolve()
    try:
        import PyInstaller.__main__ as pyi_main  # type: ignore
    except ImportError:
        logger.error(
            "尚未安裝 PyInstaller！請先執行：\n"
            "  pip install pyinstaller pdfplumber pypdf pypdfium2 pillow openpyxl pandas python-docx python-pptx tkinterdnd2 rapidocr_onnxruntime markitdown"
        )
        return 1

    work_dir = Path(tempfile.gettempdir()) / "pyinstaller_doc2md_build"
    dist_dir = script_path.parent / "dist"
    bundle_mode = "--onedir" if onedir else "--onefile"
    pyi_args: List[str] = [
        str(script_path),
        "--noconfirm",
        "--clean",
        bundle_mode,
        f"--name={exe_name}",
        f"--distpath={dist_dir}",
        f"--workpath={work_dir}",
        f"--specpath={work_dir}",
        "--console" if console else "--windowed",
    ]

    # 自動偵測環境已安裝的套件並加入 --collect-all / --hidden-import
    collect_all_pkgs = [
        "tkinterdnd2",
        "pypdfium2",
        "pypdfium2_raw",
        "rapidocr_onnxruntime",
        "pdfplumber",
        "pdfminer",
        "markitdown",
    ]
    for pkg in collect_all_pkgs:
        try:
            __import__(pkg)
            pyi_args.append(f"--collect-all={pkg}")
        except ImportError:
            pass

    hidden_imports = [
        "openpyxl",
        "docx",
        "pptx",
        "pypdf",
        "PIL",
        "pytesseract",
        "pandas",
        "markitdown",
    ]
    for mod in hidden_imports:
        try:
            __import__(mod)
            pyi_args.append(f"--hidden-import={mod}")
        except ImportError:
            pass

    # 若指定隨附可攜版 Tesseract-OCR 資料夾，一併封裝進執行檔
    if bundle_tesseract_dir:
        t_path = Path(bundle_tesseract_dir).expanduser().resolve()
        if t_path.exists():
            sep = ";" if platform.system() == "Windows" else ":"
            pyi_args.append(f"--add-data={t_path}{sep}Tesseract-OCR")
            logger.info("已將 Tesseract-OCR 資料夾納入封裝：%s", t_path)

    mode_name = "資料夾目錄模式 (One-Folder / onedir)" if onedir else "單一檔案模式 (One-File)"
    logger.info("開始執行 PyInstaller 封裝 [%s]：\n  pyinstaller %s", mode_name, " ".join(pyi_args))
    pyi_main.run(pyi_args)

    if onedir:
        dist_folder = dist_dir / exe_name
        dist_exe = dist_folder / (f"{exe_name}.exe" if platform.system() == "Windows" else exe_name)
        logger.info("打包完成！輸出資料夾目錄位置：%s", dist_folder)
        logger.info("主執行檔位置：%s", dist_exe)
    else:
        dist_exe = dist_dir / (f"{exe_name}.exe" if platform.system() == "Windows" else exe_name)
        logger.info("打包完成！輸出單一執行檔位置：%s", dist_exe)
    return 0


def generate_windows_build_bat(target_dir: Optional[Path] = None) -> Path:
    """於當前目錄產生 Windows 一鍵打包批次檔 `build_windows_exe.bat`。"""
    out_dir = target_dir or get_app_base_dir()
    bat_path = out_dir / "build_windows_exe.bat"
    bat_path.write_text(WINDOWS_BUILD_BAT_TEMPLATE, encoding="utf-8")
    logger.info("已產生 Windows 一鍵打包腳本：%s", bat_path)
    return bat_path


# ============================================================================
# 命令列介面 (CLI / GUI / Web / EXE 打包入口)
# ============================================================================

def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="將 PDF、Excel、Word、PowerPoint、純文字與掃描圖片轉為 Markdown（含 OCR、圖片匯出、GUI、Web 與 Windows .exe 打包）。"
    )
    parser.add_argument(
        "inputs",
        nargs="*",
        help="輸入檔案或資料夾路徑清單 (支援拖放多個檔案至 .exe 圖示；若完全不帶參數則預設啟動桌面 GUI)",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="輸出 .md 檔案或輸出資料夾路徑 (預設輸出於原檔案同目錄)",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="當輸入為資料夾時，遞迴搜尋所有子資料夾內的檔案",
    )
    parser.add_argument(
        "--no-images",
        action="store_true",
        help="停用內嵌圖片匯出 (預設會自動匯出 Word/PPT/PDF 圖片至 <檔名>_assets/ 並加上 Markdown 連結)",
    )
    parser.add_argument(
        "--no-ocr",
        action="store_true",
        help="停用掃描版 PDF 與圖片檔的 OCR 辨識",
    )
    parser.add_argument(
        "--ocr-images",
        action="store_true",
        help="對 Word / PPT / PDF 匯出的內嵌圖片也一併執行 OCR 文字辨識",
    )
    parser.add_argument(
        "--ocr-lang",
        default="chi_tra+eng",
        help="設定 Tesseract OCR 辨識語言 (預設: chi_tra+eng)",
    )
    parser.add_argument(
        "--engine",
        choices=["custom", "auto", "markitdown"],
        default="custom",
        help="轉換引擎：custom (預設，支援圖片匯出與 OCR)、markitdown、auto",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="啟動具備檔案拖曳功能的桌面視窗介面 (Desktop GUI)",
    )
    parser.add_argument(
        "--web",
        action="store_true",
        help="啟動網頁拖曳上傳轉檔工具 (Web UI)",
    )
    parser.add_argument(
        "--host",
        default="127.0.0.1",
        help="網頁伺服器綁定位址 (預設: 127.0.0.1)",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="網頁伺服器連接埠 (預設: 8000)",
    )
    parser.add_argument(
        "--build-exe",
        action="store_true",
        help="使用 PyInstaller 將本程式打包為免安裝 Python 的執行檔 (.exe)",
    )
    parser.add_argument(
        "--onedir",
        "--onefolder",
        dest="onedir",
        action="store_true",
        default=False,
        help="以資料夾目錄格式打包 (one-folder / onedir，啟動速度極快，免每次解壓縮)",
    )
    parser.add_argument(
        "--onefile",
        dest="onefile",
        action="store_true",
        default=False,
        help="以單一獨立執行檔模式打包 (--onefile)",
    )
    parser.add_argument(
        "--console",
        action="store_true",
        help="打包為命令列主控台視窗版本 (.exe) (預設為純視窗介面 --windowed)",
    )
    parser.add_argument(
        "--exe-name",
        default="DocumentToMarkdownPro",
        help="打包後的執行檔/資料夾名稱 (預設: DocumentToMarkdownPro)",
    )
    parser.add_argument(
        "--bundle-tesseract",
        default=None,
        help="打包時一併封裝進執行檔的可攜版 Tesseract-OCR 資料夾路徑",
    )
    parser.add_argument(
        "--generate-bat",
        action="store_true",
        help="在同目錄下產生 Windows 一鍵打包批次檔 (build_windows_exe.bat)",
    )
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    configure_windows_runtime()
    parser = build_arg_parser()
    args = parser.parse_args(argv)

    if args.generate_bat:
        generate_windows_build_bat()
        return 0

    if args.build_exe or args.onedir or args.onefile:
        # 當指定 --onedir / --onefolder 或未指定 --onefile 時，預設使用極速資料夾模式 (onedir)
        is_onedir = not args.onefile if (args.onedir or not args.onefile) else False
        return build_standalone_exe(
            exe_name=args.exe_name,
            console=args.console,
            bundle_tesseract_dir=args.bundle_tesseract,
            onedir=is_onedir,
        )

    if args.web:
        launch_web_server(host=args.host, port=args.port)
        return 0

    # 當雙擊 .exe 執行 (未帶任何參數) 或明確指定 --gui 時，啟動桌面拖曳視窗介面
    if args.gui or not args.inputs:
        initial = [Path(p).expanduser() for p in args.inputs] if args.inputs else None
        launch_desktop_gui(initial_paths=initial)
        return 0

    options = ConversionOptions(
        extract_images=not args.no_images,
        enable_ocr=not args.no_ocr,
        ocr_on_embedded_images=args.ocr_images,
        ocr_lang=args.ocr_lang,
        engine=args.engine,
    )
    converter = DocumentToMarkdownConverter(options=options)

    try:
        for raw_input in args.inputs:
            input_path = Path(raw_input).expanduser()
            if input_path.is_dir():
                results = converter.convert_directory(
                    input_dir=input_path,
                    output_dir=args.output,
                    recursive=args.recursive,
                )
                logger.info("資料夾 `%s` 批次轉換完成，共成功轉換 %d 個檔案。", input_path.name, len(results))
            else:
                converter.convert_file_detailed(input_path=input_path, output_path=args.output)
        return 0
    except Exception as exc:
        logger.error("轉換失敗：%s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())