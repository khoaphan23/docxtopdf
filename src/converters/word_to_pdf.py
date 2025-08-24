"""
Word to PDF Converter Module - Clean Class Implementation
"""
import logging
from pathlib import Path
from typing import Optional, Tuple

logger = logging.getLogger(__name__)


class DocumentConverter:
    """
    Xử lý chuyển đổi tài liệu Word sang PDF
    Hỗ trợ nhiều phương thức: docx2pdf, win32com
    """
    
    def __init__(self):
        """Khởi tạo converter với phát hiện phương thức tự động"""
        self.docx2pdf_available = self._check_docx2pdf()
        self.win32com_available = self._check_win32com()
        self.conversion_method = self._detect_best_method()
        
        logger.info(f"DocumentConverter initialized - Method: {self.conversion_method}")
    
    def _check_docx2pdf(self) -> bool:
        """Kiểm tra thư viện docx2pdf"""
        try:
            import docx2pdf
            return True
        except ImportError:
            return False
    
    def _check_win32com(self) -> bool:
        """Kiểm tra thư viện win32com"""
        try:
            import win32com.client
            return True
        except ImportError:
            return False
    
    def _detect_best_method(self) -> str:
        """
        Phát hiện phương thức chuyển đổi tốt nhất
        Ưu tiên: docx2pdf > win32com > none
        """
        if self.docx2pdf_available:
            return "docx2pdf"
        elif self.win32com_available:
            return "win32com"
        else:
            return "none"
    
    def is_available(self) -> bool:
        """Kiểm tra xem có thể chuyển đổi được không"""
        return self.conversion_method != "none"
    
    def get_supported_formats(self) -> list:
        """Lấy danh sách định dạng được hỗ trợ"""
        return [".docx", ".doc"]
    
    def get_conversion_info(self) -> dict:
        """Lấy thông tin về khả năng chuyển đổi"""
        return {
            "method": self.conversion_method,
            "available": self.is_available(),
            "docx2pdf_available": self.docx2pdf_available,
            "win32com_available": self.win32com_available,
            "supported_formats": self.get_supported_formats()
        }
    
    def _convert_with_docx2pdf(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """Chuyển đổi bằng docx2pdf"""
        try:
            from docx2pdf import convert
            convert(str(input_path), str(output_path))
            logger.info("Conversion successful with docx2pdf")
            return True, "Chuyển đổi thành công với docx2pdf"
            
        except Exception as e:
            error_msg = f"Lỗi docx2pdf: {e}"
            logger.error(error_msg)
            return False, error_msg
    
    def _convert_with_win32com(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """Chuyển đổi bằng Microsoft Word COM"""
        try:
            import win32com.client
            
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            doc = word.Documents.Open(str(input_path.absolute()))
            doc.SaveAs(str(output_path.absolute()), FileFormat=17)
            doc.Close()
            word.Quit()
            
            logger.info("Conversion successful with win32com")
            return True, "Chuyển đổi thành công với Microsoft Word"
            
        except Exception as e:
            error_msg = f"Lỗi Microsoft Word: {e}"
            logger.error(error_msg)
            return False, error_msg
    
    def convert_word_to_pdf(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """
        Chuyển đổi file Word sang PDF
        
        Args:
            input_path: Đường dẫn file Word đầu vào
            output_path: Đường dẫn file PDF đầu ra
            
        Returns:
            Tuple[bool, str]: (Thành công, Thông báo)
        """
        try:
            # Kiểm tra file input
            if not input_path.exists():
                return False, f"File không tồn tại: {input_path}"
            
            if input_path.suffix.lower() not in self.get_supported_formats():
                return False, f"Định dạng không hỗ trợ: {input_path.suffix}"
            
            # Tạo thư mục output nếu cần
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            logger.info(f"Converting: {input_path} -> {output_path}")
            
            # Thực hiện chuyển đổi theo method
            if self.conversion_method == "docx2pdf":
                return self._convert_with_docx2pdf(input_path, output_path)
            elif self.conversion_method == "win32com":
                return self._convert_with_win32com(input_path, output_path)
            else:
                error_msg = "Không có phương thức chuyển đổi. Vui lòng cài đặt Microsoft Word hoặc docx2pdf"
                logger.error(error_msg)
                return False, error_msg
                
        except Exception as e:
            error_msg = f"Lỗi không mong muốn: {e}"
            logger.error(error_msg)
            return False, error_msg
    
    def convert_to_downloads(self, input_path: Path, filename: Optional[str] = None) -> Tuple[bool, str, Optional[Path]]:
        """
        Chuyển đổi và lưu trực tiếp vào thư mục Downloads
        
        Args:
            input_path: File Word đầu vào
            filename: Tên file PDF (optional)
            
        Returns:
            Tuple[bool, str, Path]: (Thành công, Thông báo, Đường dẫn PDF)
        """
        try:
            # Tạo đường dẫn output trong Downloads
            downloads_folder = Path.home() / "Downloads"
            downloads_folder.mkdir(exist_ok=True)
            
            if filename:
                output_name = filename if filename.endswith('.pdf') else f"{filename}.pdf"
            else:
                output_name = input_path.stem + ".pdf"
                
            output_path = downloads_folder / output_name
            
            # Xử lý file trùng tên
            counter = 1
            original_path = output_path
            while output_path.exists():
                name_part = original_path.stem
                output_path = downloads_folder / f"{name_part}_{counter}.pdf"
                counter += 1
            
            # Thực hiện chuyển đổi
            success, message = self.convert_word_to_pdf(input_path, output_path)
            
            if success:
                return True, f"Đã lưu PDF vào Downloads: {output_path.name}", output_path
            else:
                return False, message, None
                
        except Exception as e:
            error_msg = f"Lỗi lưu vào Downloads: {e}"
            logger.error(error_msg)
            return False, error_msg, None