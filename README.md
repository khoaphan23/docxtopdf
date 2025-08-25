
# Word to PDF & Excel to PDF Converter

## Giới thiệu

Ứng dụng này cho phép người dùng chuyển đổi **file Word** (`.doc`, `.docx`) và **file Excel** (`.xls`, `.xlsx`) sang **định dạng PDF** dễ dàng. Dự án này sử dụng thư viện Python `docx2pdf` cho Word và `pywin32` cho Excel (nếu cần).

### Tính năng chính:
- **Chuyển Word → PDF**: Sử dụng `docx2pdf` (Windows/macOS) hoặc `win32com` (Windows + Microsoft Word).
- **Chuyển Excel → PDF**: Sử dụng `win32com` (Windows + Microsoft Excel).
- **Giao diện người dùng Tkinter**: Giao diện đơn giản để chọn file và chuyển đổi sang PDF.
- **Lưu tự động vào thư mục Downloads** sau khi chuyển đổi thành công.

## Cài đặt

Để chạy ứng dụng, bạn cần cài đặt các thư viện cần thiết.

### 1. Cài đặt yêu cầu

Clone hoặc tải về dự án này về máy tính của bạn:

```bash
git clone https://github.com/khoaphan23/docxtopdf.git
cd docxtopdf
```

Sau đó, cài đặt các phụ thuộc:

```bash
pip install -r requirements.txt
```

Các thư viện cần thiết:
- `docx2pdf`: Chuyển đổi file Word sang PDF (Windows/macOS).
- `pywin32`: Dùng để chuyển đổi Excel sang PDF trên Windows với Microsoft Excel.

### 2. Yêu cầu hệ thống

- **Windows/MacOS**: Đảm bảo đã cài đặt **Microsoft Word** (cho Word to PDF) và **Microsoft Excel** (cho Excel to PDF).
- **Linux**: Hiện tại, không hỗ trợ Excel → PDF trên Linux do phụ thuộc vào `pywin32`.

### 3. Cài đặt và chạy

1. Để chạy ứng dụng Word → PDF:
```bash
python main_word.py
```

2. Để chạy ứng dụng Excel → PDF:
```bash
python main_excel.py
```

Sau khi chạy, giao diện sẽ hiện ra cho phép bạn chọn file và chuyển đổi sang PDF. Sau khi chuyển xong, file PDF sẽ được lưu trong thư mục **Downloads**.

## Cấu trúc thư mục

```
docxtopdf/
├─ .gitignore                # Các file và thư mục không được theo dõi bởi Git
├─ README.md                 # File này
├─ requirements.txt          # Liệt kê các phụ thuộc của dự án
├─ main_word.py              # Main file cho Word → PDF Converter
├─ main_excel.py             # Main file cho Excel → PDF Converter
└─ src/
   ├─ __init__.py            # Cấu hình chung của dự án
   ├─ logging/logger_setup.py # Thiết lập logging
   ├─ interface/tkinter_ui.py # Giao diện người dùng (UI)
   ├─ io/file_handler.py      # Xử lý các thao tác file (chọn file, mở thư mục Downloads)
   └─ converters/
      ├─ word_to_pdf.py      # Chuyển đổi Word → PDF
      └─ excel_to_pdf.py     # Chuyển đổi Excel → PDF
```

## Cách sử dụng

1. **Chạy ứng dụng**:
   - Mở ứng dụng **Word → PDF** hoặc **Excel → PDF** bằng cách chạy `main_word.py` hoặc `main_excel.py`.

2. **Chọn file**:
   - Bấm nút "Chọn file Word" hoặc "Chọn file Excel" để chọn file cần chuyển đổi.

3. **Chuyển đổi**:
   - Sau khi chọn file, bấm nút "Chuyển sang PDF" để bắt đầu quá trình chuyển đổi.

4. **Mở thư mục Downloads**:
   - Sau khi chuyển đổi thành công, bấm nút "Mở thư mục Downloads" để mở thư mục chứa file PDF đã chuyển đổi.

## Ghi chú

- **Microsoft Word** và **Microsoft Excel** cần phải được cài đặt để chuyển đổi từ file Word hoặc Excel sang PDF.
- Thư mục **Downloads** sẽ được sử dụng để lưu file PDF đã chuyển đổi.

---

**Chúc bạn sử dụng ứng dụng thành công!**
