"""
Tkinter UI Module - Clean Class Implementation
"""
import tkinter as tk
from tkinter import messagebox, ttk
import logging
from pathlib import Path
from typing import Callable, Optional, Dict, Any

logger = logging.getLogger(__name__)


class ConverterUI:
    """
    Giao diện chính của ứng dụng Word to PDF Converter
    Quản lý tất cả các thành phần UI và tương tác người dùng
    """
    
    def __init__(self, 
                 root: tk.Tk,
                 callbacks: Dict[str, Callable],
                 config: Optional[Dict[str, Any]] = None):
        """
        Khởi tạo giao diện ứng dụng
        
        Args:
            root: Cửa sổ Tkinter root
            callbacks: Dict chứa các callback functions
            config: Cấu hình UI (optional)
        """
        self.root = root
        self.callbacks = callbacks
        self.config = config or {}
        
        # UI components
        self.main_frame = None
        self.select_button = None
        self.convert_button = None
        self.download_button = None
        self.progress_bar = None
        self.status_label = None
        
        # Style
        self.style = None
        
        # Setup UI
        self._setup_window()
        self._setup_styles()
        self._create_ui_components()
        self._configure_events()
        
        logger.info("ConverterUI initialized successfully")
    
    def _setup_window(self):
        """Cấu hình cửa sổ chính"""
        title = self.config.get('window_title', 'Word to PDF Converter - Chuyển đổi tài liệu')
        self.root.title(title)
        self.root.geometry('800x600')
        self.root.resizable(True, True)
        
        # Đặt cửa sổ ở giữa màn hình
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
        # Icon (nếu có)
        try:
            self.root.iconbitmap()  # Có thể thêm icon file
        except:
            pass
    
    def _setup_styles(self):
        """Cấu hình styles cho UI"""
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Button styles
        self.style.configure('Primary.TButton', 
                           font=('Arial', 12, 'bold'),
                           padding=(15, 8))
        
        self.style.configure('Secondary.TButton',
                           font=('Arial', 10),
                           padding=(10, 6))
        
        # Label styles  
        self.style.configure('Title.TLabel',
                           font=('Arial', 18, 'bold'),
                           foreground='#2c3e50')
        
        self.style.configure('Header.TLabel',
                           font=('Arial', 12, 'bold'),
                           foreground='#34495e')
        
        self.style.configure('Info.TLabel',
                           font=('Arial', 10),
                           foreground='#7f8c8d')
        
        # Progress bar style
        self.style.configure('Custom.Horizontal.TProgressbar',
                           troughcolor='#ecf0f1',
                           background='#3498db')
    
    def _create_ui_components(self):
        """Tạo các thành phần UI"""
        # Main container
        self.main_frame = ttk.Frame(self.root, padding="30")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        
        # Create sections
        self._create_header_section()
        self._create_action_section()
        self._create_progress_section()
        self._create_info_section()
        self._create_footer_section()
    
    def _create_header_section(self):
        """Tạo phần header"""
        header_frame = ttk.Frame(self.main_frame)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 30))
        header_frame.columnconfigure(0, weight=1)
        
        # Title
        title_label = ttk.Label(header_frame, 
                               text="🔄 Word to PDF Converter", 
                               style='Title.TLabel')
        title_label.grid(row=0, column=0)
        
        # Subtitle
        subtitle_label = ttk.Label(header_frame,
                                 text="Chuyển đổi tài liệu Word (.docx, .doc) sang PDF một cách nhanh chóng",
                                 style='Info.TLabel')
        subtitle_label.grid(row=1, column=0, pady=(5, 0))
    
    def _create_action_section(self):
        """Tạo phần các nút chức năng"""
        action_frame = ttk.LabelFrame(self.main_frame, 
                                    text="🎯 Thao tác chính", 
                                    padding="20")
        action_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        action_frame.columnconfigure((0, 1, 2), weight=1)
        
        # Select file button
        self.select_button = ttk.Button(
            action_frame,
            text=self.config.get('select_button_text', '📁 Chọn file Word'),
            command=self.callbacks.get('on_select_file'),
            style='Primary.TButton'
        )
        self.select_button.grid(row=0, column=0, padx=(0, 10), sticky=(tk.W, tk.E))
        
        # Convert button
        self.convert_button = ttk.Button(
            action_frame,
            text=self.config.get('convert_button_text', '🔄 Chuyển sang PDF'),
            command=self.callbacks.get('on_convert'),
            style='Primary.TButton',
            state=tk.DISABLED
        )
        self.convert_button.grid(row=0, column=1, padx=10, sticky=(tk.W, tk.E))
        
        # Download button
        self.download_button = ttk.Button(
            action_frame,
            text=self.config.get('download_button_text', '⬇️ Tải xuống'),
            command=self.callbacks.get('on_download'),
            style='Primary.TButton',
            state=tk.DISABLED
        )
        self.download_button.grid(row=0, column=2, padx=(10, 0), sticky=(tk.W, tk.E))
    
    def _create_progress_section(self):
        """Tạo phần hiển thị tiến trình"""
        progress_frame = ttk.LabelFrame(self.main_frame, 
                                      text="📊 Tiến trình", 
                                      padding="15")
        progress_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        progress_frame.columnconfigure(0, weight=1)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            mode='determinate',
            style='Custom.Horizontal.TProgressbar'
        )
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Status label
        self.status_label = ttk.Label(
            progress_frame,
            text="✅ Sẵn sàng - Hãy chọn file Word để bắt đầu",
            style='Info.TLabel'
        )
        self.status_label.grid(row=1, column=0)
    
    def _create_info_section(self):
        """Tạo phần thông tin hướng dẫn"""
        info_frame = ttk.LabelFrame(self.main_frame, 
                                  text="💡 Hướng dẫn sử dụng", 
                                  padding="15")
        info_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        
        instructions = [
            "1️⃣ Nhấn 'Chọn file Word' để chọn tài liệu cần chuyển đổi",
            "2️⃣ Nhấn 'Chuyển sang PDF' để thực hiện chuyển đổi",
            "3️⃣ Nhấn 'Tải xuống' để lưu PDF vào thư mục Downloads",
            "",
            "📝 Hỗ trợ định dạng: .docx, .doc (tối đa 100MB)",
            "🚀 Sử dụng công nghệ docx2pdf cho chất lượng tốt nhất",
            "📁 File PDF sẽ tự động lưu vào thư mục Downloads"
        ]
        
        for i, instruction in enumerate(instructions):
            if instruction:  # Skip empty lines
                label = ttk.Label(info_frame, text=instruction, style='Info.TLabel')
                label.grid(row=i, column=0, sticky=tk.W, pady=1)
            else:
                ttk.Separator(info_frame, orient='horizontal').grid(row=i, column=0, 
                                                                   sticky=(tk.W, tk.E), 
                                                                   pady=5)
    
    def _create_footer_section(self):
        """Tạo phần footer"""
        footer_frame = ttk.Frame(self.main_frame)
        footer_frame.grid(row=4, column=0, sticky=(tk.W, tk.E))
        footer_frame.columnconfigure(0, weight=1)
        
        # Footer info
        footer_label = ttk.Label(
            footer_frame,
            text="Word to PDF Converter - Phiên bản tối ưu với giao diện đẹp",
            style='Info.TLabel'
        )
        footer_label.grid(row=0, column=0, pady=(10, 0))
    
    def _configure_events(self):
        """Cấu hình các sự kiện"""
        # Window close event
        self.root.protocol("WM_DELETE_WINDOW", self.callbacks.get('on_quit', self.root.quit))
        
        # Keyboard shortcuts
        self.root.bind('<Control-o>', lambda e: self.callbacks.get('on_select_file')())
        self.root.bind('<Control-s>', lambda e: self.callbacks.get('on_convert')())
        self.root.bind('<Control-d>', lambda e: self.callbacks.get('on_download')())
        self.root.bind('<F1>', lambda e: self._show_help())
    
    def update_status(self, message: str, progress: Optional[int] = None):
        """
        Cập nhật status và progress bar
        
        Args:
            message: Thông báo trạng thái
            progress: Giá trị progress (0-100)
        """
        self.status_label.config(text=message)
        
        if progress is not None:
            self.progress_bar['value'] = progress
        
        self.root.update_idletasks()
    
    def enable_convert_button(self, enabled: bool = True):
        """Enable/disable convert button"""
        state = tk.NORMAL if enabled else tk.DISABLED
        self.convert_button.config(state=state)
    
    def enable_download_button(self, enabled: bool = True):
        """Enable/disable download button"""
        state = tk.NORMAL if enabled else tk.DISABLED
        self.download_button.config(state=state)
    
    def show_file_selected_dialog(self, file_path: Path):
        """Hiển thị dialog file đã chọn"""
        try:
            file_info = file_path.stat()
            file_size = file_info.st_size / (1024 * 1024)  # MB
            
            message = (
                f"✅ Đã chọn file thành công!\n\n"
                f"📄 Tên file: {file_path.name}\n"
                f"📁 Đường dẫn: {file_path.parent}\n"
                f"📊 Kích thước: {file_size:.2f} MB\n"
                f"📅 Sửa đổi lần cuối: {file_info.st_mtime}"
            )
            
            self.show_success_message("File đã được chọn", message)
            
        except Exception as e:
            logger.error(f"Error showing file dialog: {e}")
            self.show_success_message("File đã được chọn", f"✅ Đã chọn: {file_path.name}")
    
    def show_conversion_success_dialog(self, output_path: Path):
        """Hiển thị dialog chuyển đổi thành công"""
        try:
            file_size = output_path.stat().st_size / (1024 * 1024)  # MB
            
            message = (
                f"🎉 Chuyển đổi thành công!\n\n"
                f"📄 File PDF: {output_path.name}\n"
                f"📁 Vị trí: {output_path.parent}\n"
                f"📊 Kích thước: {file_size:.2f} MB\n\n"
                f"Bạn có thể nhấn 'Tải xuống' để sao chép vào thư mục Downloads"
            )
            
            self.show_success_message("Chuyển đổi hoàn tất", message)
            
        except Exception as e:
            logger.error(f"Error showing conversion dialog: {e}")
            self.show_success_message("Chuyển đổi hoàn tất", f"✅ Đã tạo PDF: {output_path.name}")
    
    def show_download_success_dialog(self, download_path: Path):
        """Hiển thị dialog tải xuống thành công"""
        message = (
            f"⬇️ Tải xuống thành công!\n\n"
            f"📄 File: {download_path.name}\n"
            f"📁 Thư mục Downloads: {download_path.parent}\n\n"
            f"Bạn có thể mở thư mục Downloads để xem file PDF"
        )
        
        self.show_success_message("Tải xuống hoàn tất", message)
    
    def show_success_message(self, title: str, message: str):
        """Hiển thị thông báo thành công"""
        messagebox.showinfo(f"🎉 {title}", message)
    
    def show_error_message(self, title: str, message: str):
        """Hiển thị thông báo lỗi"""
        messagebox.showerror(f"❌ {title}", message)
    
    def show_info_message(self, title: str, message: str):
        """Hiển thị thông báo thông tin"""
        messagebox.showinfo(f"ℹ️ {title}", message)
    
    def ask_yes_no(self, title: str, message: str) -> bool:
        """Hiển thị dialog xác nhận"""
        return messagebox.askyesno(f"❓ {title}", message)
    
    def _show_help(self):
        """Hiển thị help dialog"""
        help_message = (
            "🔄 Word to PDF Converter - Hướng dẫn sử dụng\n\n"
            "Phím tắt:\n"
            "• Ctrl+O: Chọn file Word\n"
            "• Ctrl+S: Chuyển đổi sang PDF\n"
            "• Ctrl+D: Tải xuống\n"
            "• F1: Hiển thị trợ giúp\n\n"
            "Hỗ trợ:\n"
            "• Định dạng: .docx, .doc\n"
            "• Kích thước tối đa: 100MB\n"
            "• Chất lượng: Cao (docx2pdf)\n\n"
            "Lưu ý: Đảm bảo file Word không bị mật khẩu bảo vệ"
        )
        
        self.show_info_message("Trợ giúp", help_message)
    
    def reset_ui(self):
        """Reset UI về trạng thái ban đầu"""
        self.enable_convert_button(False)
        self.enable_download_button(False)
        self.progress_bar['value'] = 0
        self.update_status("✅ Sẵn sàng - Hãy chọn file Word để bắt đầu")
        
        logger.info("UI reset to initial state")