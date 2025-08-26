# -*- coding: utf-8 -*-
"""
Excel -> PDF (Windows + Excel), chống vỡ layout + tránh 'Document not saved' / WinError 32
+ sửa lỗi CẮT DẤU tiếng Việt bằng cách tăng chiều cao hàng có kiểm soát.

- AutoFit cột/hàng, sau đó cộng thêm đệm:
    * ROW_PADDING_PT: đệm cơ bản cho mọi hàng
    * ROW_HEIGHT_SCALE: nhân thêm % chiều cao (để chắc chắn)
    * EXTRA_WRAP_PADDING_PT: đệm cộng thêm nếu hàng có wrap/ xuống dòng
    * TOP_ROWS_EXTRA_PAD_PT: đệm bổ sung cho vài hàng đầu (thường là tiêu đề)
- VerticalAlignment = Center để hạn chế cắt trên/dưới
- FitToPagesWide=1, FitToPagesTall=False; Landscape; lề gọn; canh giữa ngang
- Xuất ra %TEMP% rồi move về đích; nếu file đích đang khóa, tự tạo tên mới (thêm timestamp)
"""

import os
import shutil
import tempfile
from datetime import datetime

# === THAM SỐ ĐIỀU CHỈNH (tăng nếu còn cắt) ===
ROW_PADDING_PT = 6.0            # đệm cơ bản (pt)
ROW_HEIGHT_SCALE = 0.06          # cộng thêm 6% chiều cao sau AutoFit
EXTRA_WRAP_PADDING_PT = 4.0      # đệm thêm nếu hàng có WrapText/ xuống dòng
TOP_ROWS_TO_PAD = 3              # số hàng đầu coi như header
TOP_ROWS_EXTRA_PAD_PT = 4.0      # đệm thêm cho các hàng đầu

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

    _do_export(tmp)

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    try:
        if os.path.exists(out_path):
            os.remove(out_path)  # nếu đang mở sẽ ném WinError 32
        shutil.move(tmp, out_path)
        return out_path
    except Exception:
        alt = _unique_path_like(out_path)
        try:
            shutil.move(tmp, alt)
            return alt
        except Exception:
            home = os.path.expanduser("~")
            fallback_dir = os.path.join(home, "Downloads")
            os.makedirs(fallback_dir, exist_ok=True)
            alt2 = os.path.join(fallback_dir, os.path.basename(alt))
            shutil.move(tmp, alt2)
            return alt2

def excel_to_pdf(input_excel_path: str, output_pdf_path: str = None, sheet=None) -> str:
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

        wb = excel.Workbooks.Open(input_abs, UpdateLinks=0, ReadOnly=True)

        def setup_sheet(ws):
            try:
                used = ws.UsedRange

                # (1) AutoFit để có chiều cao/ rộng chuẩn
                try:
                    used.Columns.AutoFit()
                    used.Rows.AutoFit()
                except Exception:
                    pass

                # (2) Đệm chiều cao hàng:
                try:
                    first_row = used.Row
                    last_row = first_row + used.Rows.Count - 1
                    first_col = used.Column
                    last_col = first_col + used.Columns.Count - 1

                    for r in range(first_row, last_row + 1):
                        row_obj = ws.Rows(r)
                        # phát hiện hàng có wrap/ xuống dòng
                        row_has_wrap = False
                        try:
                            rng_row = ws.Range(ws.Cells(r, first_col), ws.Cells(r, last_col))
                            for cell in rng_row:
                                try:
                                    v = cell.Value
                                    if bool(cell.WrapText) or (isinstance(v, str) and ("\n" in v or "\r" in v)):
                                        row_has_wrap = True
                                        break
                                except Exception:
                                    pass
                        except Exception:
                            pass

                        try:
                            h = float(row_obj.RowHeight)
                            # nhân theo tỉ lệ rồi cộng đệm cơ bản
                            new_h = max(h * (1.0 + ROW_HEIGHT_SCALE), h + ROW_PADDING_PT)
                            # đệm thêm nếu có wrap hoặc nằm trong các hàng tiêu đề đầu
                            if row_has_wrap:
                                new_h += EXTRA_WRAP_PADDING_PT
                            if (r - first_row) < TOP_ROWS_TO_PAD:
                                new_h += TOP_ROWS_EXTRA_PAD_PT
                            row_obj.RowHeight = new_h
                        except Exception:
                            pass
                except Exception:
                    pass

                # (3) Căn giữa dọc để giảm rủi ro cắt trên/dưới
                try:
                    used.VerticalAlignment = constants.xlVAlignCenter
                except Exception:
                    pass

                # (4) Thiết lập trang in
                ps = ws.PageSetup
                try: ps.Zoom = False
                except Exception: pass
                try:
                    ps.FitToPagesWide = 1
                    ps.FitToPagesTall = False
                except Exception:
                    pass
                try: ps.Orientation = constants.xlLandscape  # đổi sang xlPortrait nếu bạn muốn
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

        # Thiết lập & export
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
