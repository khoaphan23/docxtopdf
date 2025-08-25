from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional, Tuple

logger = logging.getLogger(__name__)

class DocumentConverter:
    """Word (doc/docx) → PDF.
    1) docx2pdf (Windows/macOS)
    2) win32com (Windows + Word)
    """

    SUPPORTED_EXTS = (".doc", ".docx")

    def __init__(self) -> None:
        self.method = None
        self.available = False
        self._detect()

    def _detect(self) -> None:
        try:
            import docx2pdf  # noqa: F401
            self.method = "docx2pdf"
            self.available = True
            return
        except Exception:
            pass
        try:
            import win32com.client  # noqa: F401
            import pythoncom  # noqa: F401
            self.method = "win32com"
            self.available = True
            return
        except Exception as e:
            logger.warning(f"Word converter: no method available: {e}")
            self.method = None
            self.available = False

    def is_available(self) -> tuple[bool, str]:
        if self.available and self.method:
            return True, f"Sẵn sàng ({self.method})"
        return False, "Chưa có phương thức chuyển Word phù hợp (cần docx2pdf hoặc Word+pywin32)."

    def convert_word_to_pdf(self, input_path: Path, output_path: Path) -> tuple[bool, str]:
        try:
            input_path = Path(input_path)
            output_path = Path(output_path)
            if input_path.suffix.lower() not in self.SUPPORTED_EXTS:
                return False, "File không phải Word (.doc/.docx)."

            output_path.parent.mkdir(parents=True, exist_ok=True)

            if self.method == "docx2pdf":
                from docx2pdf import convert
                convert(str(input_path), str(output_path))
                return True, f"Đã xuất PDF: {output_path.name}"

            if self.method == "win32com":
                import pythoncom
                import win32com.client
                pythoncom.CoInitialize()
                word = None
                doc = None
                try:
                    word = win32com.client.DispatchEx("Word.Application")
                    word.Visible = False
                    doc = word.Documents.Open(str(input_path))
                    # 17 = wdFormatPDF
                    doc.SaveAs(str(output_path), FileFormat=17)
                finally:
                    try:
                        if doc is not None:
                            doc.Close(False)
                    except Exception:
                        pass
                    try:
                        if word is not None:
                            word.Quit()
                    except Exception:
                        pass
                    pythoncom.CoUninitialize()
                return True, f"Đã xuất PDF: {output_path.name}"

            return False, "Không có phương thức chuyển đổi khả dụng."
        except Exception as e:
            err = f"Lỗi Word → PDF: {e}"
            logger.exception(err)
            return False, err

    def convert_and_save_to_downloads(self, input_path: Path) -> tuple[bool, str, Optional[Path]]:
        try:
            downloads = Path.home() / "Downloads"
            downloads.mkdir(parents=True, exist_ok=True)
            name_part = Path(input_path).stem
            out = downloads / f"{name_part}.pdf"
            i = 1
            while out.exists():
                out = downloads / f"{name_part}_{i}.pdf"
                i += 1
            ok, msg = self.convert_word_to_pdf(Path(input_path), out)
            if ok:
                return True, f"Đã lưu PDF vào Downloads: {out.name}", out
            return False, msg, None
        except Exception as e:
            err = f"Lỗi lưu Word → PDF: {e}"
            logger.exception(err)
            return False, err, None
