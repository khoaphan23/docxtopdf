# -*- coding: utf-8 -*-
"""
Excel -> PDF (Windows + Excel), chống vỡ layout + tránh lỗi 'Document not saved' / WinError 32.

- PrintArea = UsedRange
- AutoFit hàng/cột
- FitToPagesWide = 1, FitToPagesTall = False (co vừa 1 trang ngang)
- Landscape mặc định (đổi sang Portrait nếu muốn)
- Lề gọn, canh giữa ngang
- Nếu file đích đang bị khóa: tự động lưu tên khác (thêm timestamp)
"""

import os
import shutil
import tempfile
from datetime import datetime

SUPPORTED_EXTS = {".xlsx", ".xls", ".xlsm", ".xlsb", ".xltx", ".xltm"}

def is_excel_file(path: str) -> bool:
    if not os.path.isfile(path):
        return False
    base = os.path.basename(path)
    if base.startswith("~$"):
        return False
    ext = os.path.splitext(base)[1].lower()
    return ext in SUPPORTED_EXTS

def _ensure_windows():
    if os.name != "nt":
        raise RuntimeError("Excel to PDF chỉ chạy trên Windows có cài Microsoft Excel.")

def _points(inches: float) -> float:
    return float(inches) * 72.0  # Excel COM dùng points

def _unique_path_like(path: str) -> str:
    base, ext = os.path.splitext(path)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{base} - {stamp}{ext}"

def _export_selected(excel, wb, out_path):
    """Export ActiveSheet(s) ra out_path. Nếu bị khoá -> đổi tên khác tự động."""
    # Xuất ra TEMP trước rồi move về đích để tránh lỗi quyền/độ dài đường dẫn.
    tmp = os.path.join(tempfile.gettempdir(), os.path.basename(out_path))
    try:
        if os.path.exists(tmp):
            os.remove(tmp)
    except Exception:
        pass

    def _do_export(target):
        excel.ActiveSheet.ExportAsFixedFormat(
            Type=0,                      # xlTypePDF
            Filename=target,
            Quality=0,                   # xlQualityStandard
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False
        )

    # Export ra TEMP
    _do_export(tmp)

    # Đảm bảo thư mục đích tồn tại
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)

    # Thử thay thế file đích (nếu tồn tại)
    try:
        if os.path.exists(out_path):
            os.remove(out_path)  # nếu file đang mở sẽ ném WinError 32
        shutil.move(tmp, out_path)
        return out_path
    except Exception:
        # File đích đang mở / khoá -> lưu với tên khác (thêm timestamp)
        alt = _unique_path_like(out_path)
        try:
            shutil.move(tmp, alt)
            return alt
        except Exception:
            # Nếu vẫn lỗi, vứt vào thư mục người dùng (Downloads) cho chắc
            home = os.path.expanduser("~")
            fallback_dir = os.path.join(home, "Downloads")
            os.makedirs(fallback_dir, exist_ok=True)
            alt2 = os.path.join(fallback_dir, os.path.basename(alt))
            shutil.move(tmp, alt2)
            return alt2

def excel_to_pdf(input_excel_path: str, output_pdf_path: str = None, sheet=None) -> str:
    """
    Convert Excel workbook/sheet to PDF (căn trang đẹp, an toàn khi file đích đang mở).

    Args:
        input_excel_path: đường dẫn file Excel.
        output_pdf_path: đường dẫn PDF (nếu None, dùng cùng thư mục/tên với .pdf).
        sheet: None (tất cả), int (1-based), hoặc str (tên sheet).
    """
    _ensure_windows()

    if not is_excel_file(input_excel_path):
        raise ValueError(f"Đường dẫn Excel không hợp lệ hoặc không hỗ trợ: {input_excel_path!r}")

    input_abs = os.path.abspath(input_excel_path)
    output_abs = (os.path.splitext(input_abs)[0] + ".pdf") if not output_pdf_path else os.path.abspath(output_pdf_path)
    output_abs = os.path.normpath(output_abs)

    try:
        import pythoncom
        from win32com.client import DispatchEx, constants, gencache
    except Exception as e:
        raise RuntimeError("Thiếu pywin32. Hãy cài: pip install pywin32") from e

    excel = None
    wb = None
    try:
        pythoncom.CoInitialize()
        try:
            gencache.EnsureDispatch("Excel.Application")
        except Exception:
            pass

        excel = DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        # Open ReadOnly để tránh đụng độ save
        wb = excel.Workbooks.Open(input_abs, UpdateLinks=0, ReadOnly=True)

        def setup_sheet(ws):
            try:
                used = ws.UsedRange
                # AutoFit giúp tránh chữ bị che
                try:
                    used.Columns.AutoFit()
                    used.Rows.AutoFit()
                except Exception:
                    pass

                ps = ws.PageSetup
                try: ps.Zoom = False
                except Exception: pass
                try:
                    ps.FitToPagesWide = 1
                    ps.FitToPagesTall = False
                except Exception:
                    pass
                try: ps.Orientation = constants.xlLandscape
                except Exception: pass
                try:
                    ps.LeftMargin   = _points(0.25)
                    ps.RightMargin  = _points(0.25)
                    ps.TopMargin    = _points(0.5)
                    ps.BottomMargin = _points(0.5)
                    ps.HeaderMargin = _points(0.3)
                    ps.FooterMargin = _points(0.3)
                except Exception:
                    pass
                try:
                    ps.CenterHorizontally = True
                    ps.CenterVertically = False
                except Exception:
                    pass
                try:
                    ps.PrintArea = used.Address
                except Exception:
                    pass
                try:
                    ws.DisplayPageBreaks = False
                except Exception:
                    pass
            except Exception:
                pass

        # Thiết lập trang
        if sheet is not None:
            ws = wb.Sheets(sheet if isinstance(sheet, int) else str(sheet))
            setup_sheet(ws)
            ws.Select()
            out = _export_selected(excel, wb, output_abs)
        else:
            for ws in wb.Worksheets:
                setup_sheet(ws)
            wb.Worksheets.Select()
            out = _export_selected(excel, wb, output_abs)

        return out
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        finally:
            if excel is not None:
                excel.EnableEvents = True
                excel.ScreenUpdating = True
                excel.DisplayAlerts = True
                excel.Quit()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
