# src/converters/word_to_pdf.py
from __future__ import annotations

import os
from pathlib import Path
from typing import Iterable, Optional, Tuple

# Hỗ trợ đuôi Word
WORD_EXTS = {".docx", ".doc"}

def is_word_file(path: str | os.PathLike) -> bool:
    return Path(path).suffix.lower() in WORD_EXTS

# -------------------- tiện ích --------------------
def mm_to_pt(mm: float) -> float:
    # 1 inch = 25.4 mm, 1 pt = 1/72 inch
    return mm * 72.0 / 25.4

def _ensure_parent_dir(p: Path) -> None:
    p.parent.mkdir(parents=True, exist_ok=True)

# -------------------- Engine: docx2pdf --------------------
def _word_to_pdf_docx2pdf(src: str, dst: str) -> None:
    """
    Dùng thư viện docx2pdf (trên Windows dùng Word ngầm).
    pip install docx2pdf
    """
    from docx2pdf import convert  # ModuleNotFoundError nếu chưa cài
    convert(src, dst)

# -------------------- Engine: COM (Word) --------------------
def _word_to_pdf_com(
    src: str,
    dst: str,
    page_size: Optional[str] = None,           # "A4" | "Letter" | None
    orientation: Optional[str] = None,         # "Portrait" | "Landscape" | None
    margins_mm: Optional[Tuple[float, float, float, float]] = None,  # (left, right, top, bottom)
    page_range: Optional[Tuple[int, int]] = None,   # (from_page, to_page), 1-based inclusive
    optimize_for: str = "Print",               # "Print" | "Screen"
    open_after_export: bool = False,
    pdf_a: bool = False                        # ISO19005-1 (PDF/A)
) -> None:
    """
    Dùng Microsoft Word qua COM (pywin32). Cần Windows + MS Word + pip install pywin32
    """
    import win32com.client as win32  # ModuleNotFoundError nếu chưa cài

    # Constants Word
    wdExportFormatPDF = 17
    wdExportOptimizeForPrint = 0
    wdExportOptimizeForOnScreen = 1
    wdExportAllDocument = 0
    wdExportFromTo = 3

    wdOrientPortrait = 0
    wdOrientLandscape = 1
    wdPaperA4 = 7
    wdPaperLetter = 2

    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    doc = None
    try:
        doc = word.Documents.Open(os.path.abspath(src))

        # Page setup (tuỳ chọn)
        if page_size or orientation or margins_mm:
            ps = doc.PageSetup
            if page_size:
                page_size_u = page_size.strip().lower()
                if page_size_u == "a4":
                    ps.PaperSize = wdPaperA4
                elif page_size_u == "letter":
                    ps.PaperSize = wdPaperLetter
            if orientation:
                ori_u = orientation.strip().lower()
                ps.Orientation = wdOrientLandscape if ori_u == "landscape" else wdOrientPortrait
            if margins_mm:
                left, right, top, bottom = margins_mm
                ps.LeftMargin = mm_to_pt(left)
                ps.RightMargin = mm_to_pt(right)
                ps.TopMargin = mm_to_pt(top)
                ps.BottomMargin = mm_to_pt(bottom)

        # Export options
        if page_range and page_range[0] >= 1 and page_range[1] >= page_range[0]:
            export_range = wdExportFromTo
            from_p, to_p = int(page_range[0]), int(page_range[1])
        else:
            export_range = wdExportAllDocument
            from_p, to_p = 1, 1  # ignored

        optimize = wdExportOptimizeForOnScreen if optimize_for.lower() == "screen" else wdExportOptimizeForPrint

        doc.ExportAsFixedFormat(
            OutputFileName=os.path.abspath(dst),
            ExportFormat=wdExportFormatPDF,
            OpenAfterExport=open_after_export,
            OptimizeFor=optimize,
            Range=export_range,
            From=from_p,
            To=to_p,
            Item=0,  # wdExportDocumentContent
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=1,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=bool(pdf_a),
        )
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()

# -------------------- API chính: word_to_pdf --------------------
def word_to_pdf(
    src_path: str,
    dst_path: Optional[str] = None,
    engine: str = "auto",                       # "auto" | "docx2pdf" | "com"
    *,
    page_size: Optional[str] = None,            # áp dụng cho COM
    orientation: Optional[str] = None,          # áp dụng cho COM
    margins_mm: Optional[Tuple[float, float, float, float]] = None,  # áp dụng cho COM
    page_range: Optional[Tuple[int, int]] = None,                    # áp dụng cho COM
    optimize_for: str = "Print",                # COM: "Print" | "Screen"
    open_after_export: bool = False,            # COM
    pdf_a: bool = False                         # COM
) -> str:
    """
    Chuyển 1 file Word (.doc/.docx) -> PDF. Trả về đường dẫn PDF.

    - Giữ tương thích: có thể truyền dst_path như tham số thứ 2 (positional), hoặc keyword.
    - engine="auto": thử docx2pdf trước, nếu thiếu thì dùng COM.
    - Khi cần khống chế layout in ấn (A4, lề, xoay ngang…), dùng engine="com" kèm các tuỳ chọn.
    """
    if not is_word_file(src_path):
        raise ValueError(f"Không phải file Word hợp lệ: {src_path}")

    src = Path(src_path).resolve()
    if dst_path is None:
        dst = src.with_suffix(".pdf")
    else:
        dst = Path(dst_path).resolve()

    _ensure_parent_dir(dst)

    def try_docx2pdf() -> bool:
        try:
            _word_to_pdf_docx2pdf(str(src), str(dst))
            return True
        except ModuleNotFoundError:
            return False

    if engine == "docx2pdf":
        if not try_docx2pdf():
            raise ModuleNotFoundError("Chưa cài docx2pdf (pip install docx2pdf)")
    elif engine == "com":
        _word_to_pdf_com(
            str(src), str(dst),
            page_size=page_size,
            orientation=orientation,
            margins_mm=margins_mm,
            page_range=page_range,
            optimize_for=optimize_for,
            open_after_export=open_after_export,
            pdf_a=pdf_a,
        )
    else:
        # auto
        if not try_docx2pdf():
            _word_to_pdf_com(
                str(src), str(dst),
                page_size=page_size,
                orientation=orientation,
                margins_mm=margins_mm,
                page_range=page_range,
                optimize_for=optimize_for,
                open_after_export=open_after_export,
                pdf_a=pdf_a,
            )

    return str(dst)
