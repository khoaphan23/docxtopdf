# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from typing import Optional
import threading
import logging

try:
    from ..converters.doc_converter import DocumentConverter
    from ..io.file_handler import FileHandler
    from ..logging.logger_setup import setup_logger, get_logger
    # Import config từ src.__init__
    from .. import (
        APP_NAME, VERSION, WINDOW_WIDTH, WINDOW_HEIGHT, 
        THEME, OUTPUT_PATH, SUPPORTED_EXTENSIONS
    )
except ImportError:
    # Fallback to direct imports when run directly
    import sys
    import importlib.util
    from pathlib import Path
    src_path = Path(__file__).parent.parent
    if str(src_path) not in sys.path:
        sys.path.insert(0, str(src_path))
    
    from converters.doc_converter import DocumentConverter
    
    # Import file_handler directly to avoid conflict with built-in io module
    file_handler_path = src_path / "io" / "file_handler.py"
    spec = importlib.util.spec_from_file_location("file_handler", file_handler_path)
    file_handler_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(file_handler_module)
    FileHandler = file_handler_module.FileHandler
    
    # Import logging setup directly to avoid conflict with built-in logging module
    logger_setup_path = src_path / "logging" / "logger_setup.py"
    spec = importlib.util.spec_from_file_location("logger_setup", logger_setup_path)
    logger_setup_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(logger_setup_module)
    setup_logger = logger_setup_module.setup_logger
    get_logger = logger_setup_module.get_logger


class DocToPdfGUI:
    def __init__(self):
        # Setup logger sử dụng config
        setup_logger()  # Sử dụng mặc định từ config
        self.logger = get_logger()
        self.converter = DocumentConverter()
        self.file_handler = FileHandler()
        
        self.root = tk.Tk()
        self.setup_window()
        self.create_widgets()
        
        self.selected_file: Optional[Path] = None
        self.converted_file: Optional[Path] = None
        
    def setup_window(self):
        # Sử dụng config cho window setup
        try:
            self.root.title(f"{APP_NAME} v{VERSION}")
            self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
            self.root.resizable(False, False)
            
            style = ttk.Style()
            style.theme_use(THEME)
        except NameError:
            # Fallback nếu không import được config
            self.root.title("Word to PDF Converter")
            self.root.geometry("600x500")
            self.root.resizable(False, False)
            
            style = ttk.Style()
            style.theme_use('clam')
        
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Sử dụng APP_NAME từ config
        try:
            title_text = f"🔄 {APP_NAME}"
        except NameError:
            title_text = "🔄 Word to PDF Converter"
            
        title_label = ttk.Label(
            main_frame, 
            text=title_text,
            font=("Arial", 16, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        self.create_step1_frame(main_frame)
        self.create_step2_frame(main_frame)
        self.create_step3_frame(main_frame)
        self.create_buttons_frame(main_frame)
        self.create_status_frame(main_frame)
        
    def create_step1_frame(self, parent):
        step1_frame = ttk.LabelFrame(parent, text="📁 Bước 1: Chọn file Word", padding="10")
        step1_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.file_info_label = ttk.Label(
            step1_frame, 
            text="Chưa chọn file nào",
            foreground="gray"
        )
        self.file_info_label.pack(anchor=tk.W)
        
        self.file_path_label = ttk.Label(
            step1_frame,
            text="",
            foreground="blue",
            font=("Arial", 8)
        )
        self.file_path_label.pack(anchor=tk.W)
        
    def create_step2_frame(self, parent):
        step2_frame = ttk.LabelFrame(parent, text="🔄 Bước 2: Chuyển đổi sang PDF", padding="10")
        step2_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.convert_info_label = ttk.Label(
            step2_frame,
            text="Chưa chuyển đổi",
            foreground="gray"
        )
        self.convert_info_label.pack(anchor=tk.W)
        
        self.convert_path_label = ttk.Label(
            step2_frame,
            text="",
            foreground="blue",
            font=("Arial", 8)
        )
        self.convert_path_label.pack(anchor=tk.W)
        
        self.progress = ttk.Progressbar(
            step2_frame,
            mode='indeterminate'
        )
        self.progress.pack(fill=tk.X, pady=(5, 0))
        
    def create_step3_frame(self, parent):
        step3_frame = ttk.LabelFrame(parent, text="⬇️ Bước 3: Tải xuống", padding="10")
        step3_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.download_info_label = ttk.Label(
            step3_frame,
            text="💾 File PDF sẽ được tải xuống trong thư mục Downloads",
            foreground="gray"
        )
        self.download_info_label.pack(anchor=tk.W)
        
    def create_buttons_frame(self, parent):
        buttons_frame = ttk.Frame(parent)
        buttons_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.select_btn = ttk.Button(
            buttons_frame,
            text="📁 Chọn file DOCX",
            command=self.select_file,
            width=20
        )
        self.select_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.convert_btn = ttk.Button(
            buttons_frame,
            text="🔄 Chuyển sang PDF",
            command=self.convert_file,
            state=tk.DISABLED,
            width=20
        )
        self.convert_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.download_btn = ttk.Button(
            buttons_frame,
            text="⬇️ Tải xuống",
            command=self.download_file,
            state=tk.DISABLED,
            width=20
        )
        self.download_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.exit_btn = ttk.Button(
            buttons_frame,
            text="❌ Thoát",
            command=self.root.quit,
            width=15
        )
        self.exit_btn.pack(side=tk.RIGHT)
        
    def create_status_frame(self, parent):
        status_frame = ttk.Frame(parent)
        status_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.status_label = ttk.Label(
            status_frame,
            text="📄 Sẵn sàng - Hãy chọn file Word để bắt đầu",
            relief=tk.SUNKEN,
            anchor=tk.W,
            padding="5"
        )
        self.status_label.pack(fill=tk.X)
        
    def select_file(self):
        # Sử dụng SUPPORTED_EXTENSIONS từ config
        try:
            extensions = " ".join([f"*{ext}" for ext in SUPPORTED_EXTENSIONS])
            file_types = [
                ("Word files", extensions),
                ("All files", "*.*")
            ]
        except NameError:
            # Fallback
            file_types = [
                ("Word files", "*.docx *.doc"),
                ("All files", "*.*")
            ]
        
        file_path = filedialog.askopenfilename(
            title="Chọn file Word",
            filetypes=file_types
        )
        
        if file_path:
            self.selected_file = Path(file_path)
            
            if self.file_handler.validate_input_file(self.selected_file):
                file_info = self.file_handler.get_file_info(self.selected_file)
                
                self.file_info_label.config(
                    text=f"✅ Đã chọn: {file_info['name']} ({file_info['size_mb']} MB)",
                    foreground="green"
                )
                self.file_path_label.config(text=f"Đường dẫn: {file_info['path']}")
                
                self.convert_btn.config(state=tk.NORMAL)
                self.status_label.config(text="✅ File đã được chọn - Có thể chuyển đổi")
            else:
                self.selected_file = None
                messagebox.showerror("Lỗi", "File không hợp lệ hoặc không được hỗ trợ")
                
    def convert_file(self):
        if not self.selected_file:
            return
            
        self.convert_btn.config(state=tk.DISABLED)
        self.progress.start()
        self.status_label.config(text="🔄 Đang chuyển đổi...")
        
        def conversion_thread():
            try:
                self.logger.info(f"Starting conversion thread for file: {self.selected_file}")
                
                # Check converter methods before convert
                methods = self.converter.check_conversion_methods()
                self.logger.info(f"Available methods: {methods}")
                
                output_dir = self.file_handler.create_output_directory()
                output_file = output_dir / f"{self.selected_file.stem}.pdf"
                
                self.logger.info(f"Output file will be: {output_file}")
                
                self.converted_file = self.converter.convert_to_pdf(
                    self.selected_file, 
                    output_file
                )
                
                self.logger.info("Conversion completed, calling success callback")
                self.root.after(0, self.conversion_success)
                
            except Exception as e:
                self.logger.error(f"Conversion error: {e}")
                self.logger.error(f"Exception type: {type(e)}")
                import traceback
                self.logger.error(f"Traceback: {traceback.format_exc()}")
                error_msg = str(e)
                self.root.after(0, lambda: self.conversion_error(error_msg))
                
        threading.Thread(target=conversion_thread, daemon=True).start()
        
    def conversion_success(self):
        self.progress.stop()
        
        file_info = self.file_handler.get_file_info(self.converted_file)
        self.convert_info_label.config(
            text=f"PDF created: {file_info['name']} ({file_info['size_mb']} MB)",
            foreground="green"
        )
        self.convert_path_label.config(text=f"Location: {file_info['path']}")
        
        self.download_btn.config(state=tk.NORMAL)
        self.convert_btn.config(state=tk.NORMAL)
        self.status_label.config(text="Conversion completed - Ready to download")
        
    def conversion_error(self, error_msg):
        self.progress.stop()
        self.convert_btn.config(state=tk.NORMAL)
        self.status_label.config(text="Conversion failed")
        messagebox.showerror("Conversion Error", f"Cannot convert file:\n{error_msg}")
        
    def download_file(self):
        if not self.converted_file:
            return
            
        try:
            downloaded_file = self.file_handler.copy_to_downloads(self.converted_file)
            
            self.download_info_label.config(
                text=f"✅ Đã tải xuống: {downloaded_file.name}",
                foreground="green"
            )
            
            self.status_label.config(text="✅ Hoàn tất - File đã được tải về Downloads")
            
            messagebox.showinfo(
                "Thành công",
                f"File PDF đã được sao chép vào:\n{downloaded_file}"
            )
            
        except Exception as e:
            self.logger.error(f"Lỗi tải xuống: {e}")
            messagebox.showerror("Lỗi", f"Không thể tải xuống file:\n{e}")
            
    def run(self):
        self.root.mainloop()