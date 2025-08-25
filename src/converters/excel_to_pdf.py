# -*- coding: utf-8 -*-
from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional, Tuple

logger = logging.getLogger(__name__)

class ExcelConverter:
    """Excel (xls/xlsx) → PDF via COM (Windows + Microsoft Excel + pywin32)."""
    SUPPORTED_EXTS = (".xls", ".xlsx")

    def __init__(self) -> None:
        self.method, self.available = self._detect_best_method()

    def _detect_best_method(self) -> Tuple[Optional[str], bool]:
        try:
            import win32com.client  # noqa
            import pythoncom  # noqa
            return "win32com", True
        except Exception as e:
            logger.warning(f"ExcelConverter: win32com not available: {e}")
            return None, False

    def is_available(self) -> Tuple[bool, str]:
        if self.available and self.method == "win32com":
            return True, "Sẵn sàng (win32com + Microsoft Excel)"
        return False, "Thiếu pywin32 hoặc Microsoft Excel chưa cài/không khả dụng (Windows)."

    def _validate_input(self, input_path: Path) -> Tuple[bool, str]:
        if not input_path:
            return False, "Đường dẫn tệp rỗng."
        p = Path(input_path)
        if not p.exists():
            return False, f"Không tìm thấy tệp: {input_path}"
        if p.is_dir():
            return False, "Đường dẫn là thư mục, không phải tệp Excel."
        if p.suffix.lower() not in self.SUPPORTED_EXTS:
            return False, f"Tệp không phải Excel hợp lệ {self.SUPPORTED_EXTS}."
        return True, "OK"

    def convert_excel_to_pdf(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        try:
            ok, msg = self._validate_input(input_path)
            if not ok:
                return False, msg
            if not self.available or self.method != "win32com":
                return False, "Không có phương thức chuyển Excel phù hợp (win32com không khả dụng)."

            out = Path(output_path)
            out.parent.mkdir(parents=True, exist_ok=True)

            import pythoncom
            import win32com.client

            pythoncom.CoInitialize()
            excel = None
            wb = None
            try:
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                wb = excel.Workbooks.Open(str(Path(input_path)))
                # 0=xlTypePDF, 0=xlQualityStandard
                wb.ExportAsFixedFormat(0, str(out), 0, True, False, 1, 9999, False)
            finally:
                try:
                    if wb is not None:
                        wb.Close(SaveChanges=False)
                except Exception:
                    pass
                try:
                    if excel is not None:
                        excel.Quit()
                except Exception:
                    pass
                pythoncom.CoUninitialize()

            return True, "Chuyển Excel → PDF thành công"
        except Exception as e:
            err = f"Lỗi chuyển Excel → PDF: {e}"
            logger.exception(err)
            return False, err

    def convert_and_save_to_downloads(self, input_path: Path) -> Tuple[bool, str, Optional[Path]]:
        try:
            downloads = Path.home() / "Downloads"
            downloads.mkdir(parents=True, exist_ok=True)
            name = Path(input_path).stem
            out = downloads / f"{name}.pdf"
            i = 1
            while out.exists():
                out = downloads / f"{name}_{i}.pdf"
                i += 1
            ok, msg = self.convert_excel_to_pdf(Path(input_path), out)
            if ok:
                return True, f"Đã lưu PDF vào Downloads: {out.name}", out
            return False, msg, None
        except Exception as e:
            err = f"Lỗi lưu PDF Excel vào Downloads: {e}"
            logger.exception(err)
            return False, err, None
