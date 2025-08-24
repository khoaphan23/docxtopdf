"""
Word to PDF Converter - Clean Class-based Application
"""
import tkinter as tk
import logging
from pathlib import Path
from typing import Optional

# Import các class từ modules
from src.converters.word_to_pdf import DocumentConverter
from src.io.file_handler import FileHandler
from src.interface.tkinter_ui import ConverterUI
from src.logging.logger_setup import setup_logger
import src as config


class WordToPDFApplication:
    """
    Ứng dụng chuyển đổi Word sang PDF chính
    Sử dụng cấu trúc class clean và tối ưu
    """
    
    def __init__(self):
        """Khởi tạo ứng dụng"""
        # Setup logging
        self.logger = setup_logger(
            name=__name__,
            level=config.LOG_LEVEL,
            log_format=getattr(config, "LOG_FORMAT", "%(asctime)s - %(name)s - %(levelname)s - %(message)s")
        )
        
        # Initialize core components
        self.converter = DocumentConverter()
        self.file_handler = FileHandler()
        
        # UI components
        self.root: Optional[tk.Tk] = None
        self.ui: Optional[ConverterUI] = None
        
        # Application state
        self.selected_file: Optional[Path] = None
        self.converted_file: Optional[Path] = None
        
        self.logger.info("WordToPDFApplication initialized successfully")
    
    def _setup_callbacks(self) -> dict:
        """Tạo dict các callback functions cho UI"""
        return {
            'on_select_file': self._handle_select_file,
            'on_convert': self._handle_convert_file,
            'on_download': self._handle_download_file,
            'on_quit': self._handle_quit
        }
    
    def _setup_ui_config(self) -> dict:
        """Tạo dict cấu hình UI"""
        return {
            'window_title': config.WINDOW_TITLE,
            'select_button_text': config.SELECT_BUTTON_TEXT,
            'convert_button_text': config.CONVERT_BUTTON_TEXT,
            'download_button_text': config.DOWNLOAD_BUTTON_TEXT,
            'quit_button_text': config.QUIT_BUTTON_TEXT
        }
    
    def _handle_select_file(self):
        """Xử lý chọn file Word"""
        try:
            self.ui.update_status("🔍 Đang mở hộp thoại chọn file...", 10)
            
            # Chọn file
            file_path = self.file_handler.select_word_file()
            
            if file_path:
                # Validate file
                is_valid, message = self.file_handler.validate_word_file(file_path)
                
                if not is_valid:
                    self.ui.show_error_message("File không hợp lệ", message)
                    self.ui.update_status("❌ File không hợp lệ", 0)
                    return
                
                # File hợp lệ
                self.selected_file = file_path
                self.ui.show_file_selected_dialog(file_path)
                self.ui.enable_convert_button(True)
                self.ui.update_status("✅ Đã chọn file - Sẵn sàng chuyển đổi", 25)
                
                self.logger.info(f"File selected and validated: {file_path}")
                
            else:
                self.ui.update_status("✅ Sẵn sàng - Hãy chọn file Word để bắt đầu", 0)
                self.logger.info("File selection cancelled by user")
                
        except Exception as e:
            error_msg = f"Lỗi khi chọn file: {e}"
            self.logger.error(error_msg)
            self.ui.show_error_message("Lỗi chọn file", error_msg)
            self.ui.update_status("❌ Lỗi chọn file", 0)
    
    def _handle_convert_file(self):
        """Xử lý chuyển đổi file Word sang PDF"""
        try:
            if not self.selected_file:
                self.ui.show_error_message("Chưa chọn file", "Vui lòng chọn file Word trước!")
                return
            
            # Kiểm tra khả năng chuyển đổi
            if not self.converter.is_available():
                info = self.converter.get_conversion_info()
                error_msg = (f"Không thể chuyển đổi!\n\n"
                           f"Phương thức: {info['method']}\n"
                           f"docx2pdf: {info['docx2pdf_available']}\n"
                           f"win32com: {info['win32com_available']}\n\n"
                           f"Vui lòng cài đặt: pip install docx2pdf")
                self.ui.show_error_message("Thiếu thư viện", error_msg)
                return
            
            # Disable convert button
            self.ui.enable_convert_button(False)
            self.ui.update_status("🔄 Đang chuyển đổi...", 50)
            
            # Generate output path
            output_path = self.file_handler.generate_output_path(
                input_file=self.selected_file,
                output_format=config.OUTPUT_FORMAT
            )
            
            self.logger.info(f"Converting: {self.selected_file} -> {output_path}")
            
            # Perform conversion
            success, message = self.converter.convert_word_to_pdf(
                input_path=self.selected_file,
                output_path=output_path
            )
            
            if success:
                self.converted_file = output_path
                self.ui.show_conversion_success_dialog(output_path)
                self.ui.enable_download_button(True)
                self.ui.update_status("🎉 Chuyển đổi thành công - Có thể tải xuống", 100)
                
                self.logger.info(f"Conversion completed successfully: {output_path}")
                
            else:
                self.ui.show_error_message("Chuyển đổi thất bại", message)
                self.ui.update_status("❌ Chuyển đổi thất bại", 25)
                self.logger.error(f"Conversion failed: {message}")
                
        except Exception as e:
            error_msg = f"Lỗi không mong muốn: {e}"
            self.logger.error(error_msg)
            self.ui.show_error_message("Lỗi chuyển đổi", error_msg)
            self.ui.update_status("❌ Lỗi chuyển đổi", 25)
            
        finally:
            # Re-enable convert button
            self.ui.enable_convert_button(True)
    
    def _handle_download_file(self):
        """Xử lý tải xuống file PDF"""
        try:
            if not self.converted_file or not self.converted_file.exists():
                self.ui.show_error_message(
                    "Chưa có file PDF", 
                    "Vui lòng chuyển đổi file Word trước khi tải xuống!"
                )
                return
            
            self.ui.update_status("⬇️ Đang tải xuống...", 75)
            
            # Copy to Downloads
            success, message, downloads_path = self.file_handler.copy_to_downloads(
                source_file=self.converted_file
            )
            
            if success and downloads_path:
                self.ui.show_download_success_dialog(downloads_path)
                self.ui.update_status("✅ Tải xuống thành công!", 100)
                self.logger.info(f"Download successful: {downloads_path}")
                
            else:
                self.ui.show_error_message("Tải xuống thất bại", message)
                self.ui.update_status("❌ Tải xuống thất bại", 100)
                self.logger.error(f"Download failed: {message}")
                
        except Exception as e:
            error_msg = f"Lỗi khi tải xuống: {e}"
            self.logger.error(error_msg)
            self.ui.show_error_message("Lỗi tải xuống", error_msg)
            self.ui.update_status("❌ Lỗi tải xuống", 100)
    
    def _handle_quit(self):
        """Xử lý thoát ứng dụng"""
        try:
            self.logger.info("Application shutdown initiated")
            
            # Confirm quit if there's work in progress
            if self.converted_file and self.converted_file.exists():
                if not self.ui.ask_yes_no("Xác nhận thoát", 
                                        "Bạn có chắc muốn thoát?\nFile PDF đã được tạo."):
                    return
            
            # Cleanup
            cleanup_count = self.file_handler.cleanup_temp_files() if hasattr(self.file_handler, "cleanup_temp_files") else 0
            if cleanup_count > 0:
                self.logger.info(f"Cleaned up {cleanup_count} temporary files")
            
            # Close application
            if self.root:
                self.root.destroy()
                
            self.logger.info("Application shutdown completed")
            
        except Exception as e:
            self.logger.error(f"Error during shutdown: {e}")
    
    def run(self):
        """Chạy ứng dụng chính"""
        try:
            # Create root window
            self.root = tk.Tk()
            
            # Setup callbacks và config
            callbacks = self._setup_callbacks()
            ui_config = self._setup_ui_config()
            
            # Create UI
            self.ui = ConverterUI(
                root=self.root,
                callbacks=callbacks,
                config=ui_config
            )
            
            # Show initial status
            converter_info = self.converter.get_conversion_info()
            if converter_info['available']:
                status_msg = f"✅ Sẵn sàng ({converter_info['method']}) - Hãy chọn file Word"
            else:
                status_msg = "⚠️ Chưa cài đặt thư viện chuyển đổi"
            
            self.ui.update_status(status_msg, 0)
            
            self.logger.info("Starting main application loop")
            
            # Start main loop
            self.root.mainloop()
            
        except Exception as e:
            error_msg = f"Lỗi nghiêm trọng: {e}"
            self.logger.error(error_msg)
            print(error_msg)
            
        finally:
            self._handle_quit()


def main():
    """
    Entry point của ứng dụng
    Tạo và chạy WordToPDFApplication
    """
    try:
        app = WordToPDFApplication()
        app.run()
        
    except Exception as e:
        print(f"Failed to start application: {e}")
        logging.error(f"Failed to start application: {e}")


if __name__ == "__main__":
    main()