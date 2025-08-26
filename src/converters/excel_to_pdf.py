# -*- coding: utf-8 -*-

import os

SUPPORTED_EXTS = {".xlsx", ".xls", ".xlsm", ".xlsb", ".xltx", ".xltm"}

def is_excel_file(path: str) -> bool:
    if not os.path.isfile(path):
        return False
    base = os.path.basename(path)
    if base.startswith("~$"):
        return False
    ext = os.path.splitext(base)[1].lower()
    return ext in SUPPORTED_EXTS

def excel_to_pdf(input_excel_path: str, output_pdf_path: str = None, sheet=None) -> str:
    """
    Chuyển 1 file Excel -> PDF.
    - input_excel_path: đường dẫn tới file Excel.
    - output_pdf_path: đường dẫn output PDF (nếu None sẽ dùng cùng tên với file Excel).
    - sheet: có thể là tên sheet hoặc index (1-based). Nếu None sẽ export workbook (tất cả sheet printable).
    Trả về: đường dẫn PDF đã tạo.
    """
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as ie:
        raise RuntimeError("Thiếu thư viện pywin32. Cài bằng: pip install pywin32") from ie

    if not os.path.isfile(input_excel_path):
        raise FileNotFoundError(f"Không thấy file: {input_excel_path}")
    if not is_excel_file(input_excel_path):
        raise ValueError("File không phải định dạng Excel hợp lệ")

    input_abs = os.path.abspath(input_excel_path)
    if output_pdf_path is None:
        output_pdf_path = os.path.splitext(input_abs)[0] + ".pdf"
    output_abs = os.path.abspath(output_pdf_path)
    os.makedirs(os.path.dirname(output_abs), exist_ok=True)

    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(input_abs)

        if sheet is None:
            wb.ExportAsFixedFormat(Type=0, Filename=output_abs)  # 0 = xlTypePDF
        else:
            try:
                ws = wb.Sheets(sheet)  # tên sheet
            except Exception:
                idx = int(sheet)
                ws = wb.Sheets(idx)
            ws.ExportAsFixedFormat(Type=0, Filename=output_abs)

        return output_abs
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
