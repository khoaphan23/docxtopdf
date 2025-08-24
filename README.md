# Ứng Dụng Chuyển Đổi Word sang PDF

Ứng dụng desktop đơn giản và hiệu quả để chuyển đổi tài liệu Microsoft Word (.doc và .docx) sang định dạng PDF với giao diện đồ họa thân thiện.

![Python](https://img.shields.io/badge/python-v3.8+-blue.svg)
![Platform](https://img.shields.io/badge/platform-windows-lightgrey.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

## 📋 Tính Năng

- **Hỗ Trợ Nhiều Định Dạng**: Chuyển đổi cả file `.doc` và `.docx` sang PDF
- **Engine Chuyển Đổi Thông Minh**: Tự động chọn phương pháp chuyển đổi tốt nhất
  - File `.doc` → Sử dụng Microsoft Word (win32com) để tối ưu khả năng tương thích
  - File `.docx` → Sử dụng docx2pdf để chuyển đổi nhanh hơn
- **Giao Diện Thân Thiện**: Interface đơn giản với khả năng kéo thả được xây dựng bằng tkinter
- **Xử Lý Hàng Loạt**: Chuyển đổi nhiều file một cách hiệu quả
- **Ghi Log Chi Tiết**: Báo cáo lỗi toàn diện và logs chuyển đổi
- **Quản Lý Output Tự Động**: Tổ chức PDF đầu ra với vị trí có thể tùy chỉnh
- **Khôi Phục Lỗi**: Các phương pháp chuyển đổi dự phòng đảm bảo tỷ lệ thành công cao

## 🚀 Bắt Đầu Nhanh

### Phương án 1: Chạy từ Source Code
```bash
# Clone hoặc tải project về
cd docxtopdf_thangvk

# Cài đặt các thư viện cần thiết
pip install -r requirements.txt

# Chạy ứng dụng
python main.py
```

### Phương án 2: Sử dụng File Thực Thi Đã Build
1. Chạy build script: `python build.py`
2. Tìm file `WordToPDF_Converter.exe` trong thư mục `dist`
3. Double-click để chạy (không cần cài đặt Python)

## 🛠️ Cài Đặt

### Yêu Cầu Hệ Thống
- **Python 3.8+** (để chạy từ source code)
- **Microsoft Word** (khuyến khích cho việc chuyển đổi file .doc)
- **Windows OS** (hỗ trợ chính)

### Thư Viện Phụ Thuộc
```bash
pip install pywin32 docx2pdf python-docx tkinter
```

### Thiết Lập Phát Triển
```bash
# Clone repository
git clone <repository-url>
cd docxtopdf_thangvk

# Cài đặt dependencies cho development
pip install -r requirements.txt

# Chạy ở chế độ phát triển
python main.py
```

## 📁 Cấu Trúc Dự Án

```
docxtopdf_thangvk/
├── main.py                 # Điểm khởi đầu chính với khởi tạo GUI
├── main_force.py           # Điểm khởi đầu thay thế với forced module reloading
├── build.py               # Build script để tạo file thực thi
├── requirements.txt       # Các thư viện Python cần thiết
├── README.md             # Tài liệu dự án
├── config.ini            # Cài đặt cấu hình
├── src/                  # Thư mục source code chính
│   ├── converters/       # Modules engine chuyển đổi
│   │   └── doc_converter.py    # Converter chính với subprocess approach
│   ├── interface/        # Components GUI
│   │   └── gui.py           # Ứng dụng GUI chính
│   ├── io/              # Utilities xử lý file
│   │   └── file_handler.py    # Validation và quản lý file
│   └── logging/         # Hệ thống logging
│       └── logger_setup.py    # Cấu hình logging tập trung
├── logs/                # Logs ứng dụng
│   └── docxtopdf_*.log     # Files log hàng ngày
├── PDF_Output/          # Thư mục output mặc định cho PDF đã chuyển đổi
└── dist/               # File thực thi đã build (sau khi chạy build.py)
    └── WordToPDF_Converter.exe
```

## 🔧 Cách Sử Dụng

### Ứng Dụng GUI
1. **Khởi động**: Chạy `python main.py` hoặc double-click file thực thi
2. **Chọn File**: Click "Choose File" để duyệt và chọn tài liệu Word
3. **Chuyển đổi**: Click "Convert to PDF" để bắt đầu quá trình chuyển đổi
4. **Theo dõi**: Xem tiến trình và kiểm tra logs để có thông tin chi tiết
5. **Truy cập Kết quả**: Tìm PDF đã chuyển đổi trong thư mục `PDF_Output`

### Sử Dụng Command Line
```python
from src.converters.doc_converter import DocumentConverter

converter = DocumentConverter()
result = converter.convert_to_pdf('input.docx', 'output.pdf')
print(f"Đã chuyển đổi: {result}")
```

## ⚙️ Cấu Hình

### Phương Pháp Chuyển Đổi
Ứng dụng tự động chọn phương pháp chuyển đổi tối ưu:

1. **Đối với file .doc**:
   - Chính: Microsoft Word COM (win32com)
   - Dự phòng: docx2pdf (khả năng tương thích hạn chế)

2. **Đối với file .docx**:
   - Chính: docx2pdf (nhanh hơn, không cần Word)
   - Dự phòng: Microsoft Word COM (tương thích cao hơn)

### Cấu Hình Logging
- **Vị trí**: `logs/docxtopdf_YYYYMMDD.log`
- **Cấp độ**: INFO (có thể cấu hình trong logger_setup.py)
- **Tính năng**: Rotation, detailed error traces, conversion metrics

### Cài Đặt Output
- **Thư mục mặc định**: `PDF_Output/`
- **Quy tắc đặt tên**: `{tên_file_gốc}.pdf`
- **Xử lý đường dẫn**: Tự động tạo thư mục và resolve path

## 🐛 Khắc Phục Sự Cố

### Các Vấn Đề Thường Gặp

#### Lỗi "All methods failed"
- **Nguyên nhân**: Cả hai phương pháp chuyển đổi đều thất bại
- **Giải pháp**: Kiểm tra logs chi tiết để xem thông báo lỗi cụ thể
- **Yêu cầu**: Đảm bảo Microsoft Word đã được cài đặt cho file .doc

#### Lỗi Import
- **Nguyên nhân**: Thiếu Python dependencies
- **Giải pháp**: Chạy `pip install -r requirements.txt`

#### Lỗi Quyền Truy Cập
- **Nguyên nhân**: File đang được sử dụng hoặc không đủ quyền
- **Giải pháp**: Đóng tài liệu Word và chạy với quyền administrator nếu cần

#### Vấn Đề Module Caching
- **Nguyên nhân**: Xung đột Python module cache
- **Giải pháp**: Sử dụng `main_force.py` thay vì `main.py`

### Chế Độ Debug
Để debug chi tiết, kiểm tra log files trong thư mục `logs/`:
```bash
tail -f logs/docxtopdf_*.log
```

## 🏗️ Build File Thực Thi

Tạo file thực thi standalone không cần cài đặt Python:

```bash
# Chạy build script
python build.py

# Output sẽ có trong dist/WordToPDF_Converter.exe
```

### Tính Năng Build
- **File thực thi đơn**: Tất cả được gom vào một file .exe
- **Không có console window**: Trải nghiệm GUI sạch sẽ
- **Bao gồm tất cả dependencies**: win32com, docx2pdf, tkinter
- **Tự động tạo spec file**: Cấu hình PyInstaller được tối ưu

## 🔍 Chi Tiết Kỹ Thuật

### Kiến Trúc Chuyển Đổi
- **Subprocess Isolation**: Mỗi chuyển đổi chạy trong process Python riêng biệt
- **Threading Support**: GUI không blocking với background conversion
- **Error Recovery**: Nhiều chiến lược fallback cho chuyển đổi mạnh mẽ
- **Memory Management**: Xử lý hiệu quả tài liệu lớn

### Tính Năng Bảo Mật
- **Path Validation**: Input sanitization và path traversal protection
- **File Type Verification**: Validation extension và content
- **Process Isolation**: Subprocess execution ngăn xung đột hệ thống

### Tối Ưu Hiệu Suất
- **Smart Method Selection**: File-type aware conversion routing
- **Parallel Processing**: Multi-threaded GUI và conversion engine
- **Resource Management**: Automatic cleanup và memory optimization
- **Caching Strategy**: Intelligent module loading và reloading

## 🤝 Đóng Góp

1. **Fork** repository
2. **Tạo** feature branch (`git checkout -b feature/TinhNangTuyetVoi`)
3. **Commit** thay đổi (`git commit -m 'Thêm tính năng tuyệt vời'`)
4. **Push** lên branch (`git push origin feature/TinhNangTuyetVoi`)
5. **Tạo** Pull Request

### Hướng Dẫn Phát Triển
- Tuân theo coding standards PEP 8
- Thêm comprehensive logging cho tính năng mới
- Bao gồm error handling cho tất cả external dependencies
- Test với cả file .doc và .docx
- Cập nhật documentation cho API changes

## 📝 Giấy Phép

Dự án này được cấp phép theo MIT License - xem file [LICENSE](LICENSE) để biết chi tiết.

## 🙏 Lời Cảm Ơn

- **Microsoft Word COM API** cho việc xử lý file .doc
- **Thư viện docx2pdf** cho chuyển đổi .docx hiệu quả
- **PyInstaller** cho packaging executable
- **tkinter** cho GUI framework

## 📞 Hỗ Trợ

Nếu bạn gặp vấn đề hoặc có câu hỏi:

1. **Kiểm tra Logs**: Xem `logs/docxtopdf_*.log` để có thông tin lỗi chi tiết
2. **GitHub Issues**: Báo cáo bugs và yêu cầu tính năng
3. **Documentation**: Tham khảo README này và inline code comments

## 🔄 Lịch Sử Phiên Bản

- **v1.0.0**: Release đầu tiên với chức năng chuyển đổi cơ bản
- **v1.1.0**: Thêm smart method selection và cải thiện error handling
- **v1.2.0**: Triển khai subprocess architecture để ổn định hơn
- **v1.3.0**: Thêm comprehensive logging và cải thiện GUI

---

**Được tạo với ❤️ để chuyển đổi tài liệu hiệu quả**
[Người dùng] 
     ↓
[interface.py] (Nhận yêu cầu → Gọi converter)
     ↓
[io_manager.py] (Đọc file đầu vào)
     ↓
[converter.py] (Xử lý chuyển đổi PNG → JPG)
     ↓
[io_manager.py] (Lưu file đầu ra)
     ↓
[logging_setup.py] (Ghi trạng thái: thành công/lỗi)
