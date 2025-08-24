# -*- coding: utf-8 -*-
import os
import sys
import subprocess
from pathlib import Path
import logging

# Import config từ src.__init__
try:
    from .. import (
        CONVERSION_TIMEOUT, CONVERSION_METHODS, 
        DEFAULT_METHOD, BACKUP_METHOD
    )
except ImportError:
    # Fallback values
    CONVERSION_TIMEOUT = 60
    CONVERSION_METHODS = {
        'win32com': 'Microsoft Word COM',
        'docx2pdf': 'docx2pdf Library',
        'libreoffice': 'LibreOffice'
    }
    DEFAULT_METHOD = 'auto'
    BACKUP_METHOD = 'docx2pdf'


class DocumentConverter:
    def __init__(self):
        self.logger = logging.getLogger(__name__)

    def convert_to_pdf(self, input_path, output_path=None):
        from pathlib import Path

        input_path = Path(input_path)

        if not input_path.exists():
            raise FileNotFoundError(f"File not found: {input_path}")

        if output_path is None:
            output_path = input_path.parent / f"{input_path.stem}.pdf"
        else:
            output_path = Path(output_path)

        self.logger.info(f"SIMPLE_CONVERT: {input_path} -> {output_path}")

        # Ensure output directory exists
        output_path.parent.mkdir(parents=True, exist_ok=True)

        # Check file extension to choose best method
        file_ext = input_path.suffix.lower()
        
        if file_ext == '.doc':
            # For .doc files, try win32com first (it handles .doc better)
            self.logger.info("SIMPLE_CONVERT: .doc file detected - trying win32com first...")
            if self._try_win32com_conversion(input_path, output_path):
                return output_path
            # Fallback to docx2pdf (though it will likely fail for .doc)
            if self._try_docx2pdf_conversion(input_path, output_path):
                return output_path
        else:
            # For .docx files, try docx2pdf first (it's faster)
            self.logger.info("SIMPLE_CONVERT: .docx file detected - trying docx2pdf first...")
            if self._try_docx2pdf_conversion(input_path, output_path):
                return output_path
            # Fallback to win32com
            if self._try_win32com_conversion(input_path, output_path):
                return output_path

        # If we get here, everything failed
        raise RuntimeError("SIMPLE_CONVERT: All methods failed")

    def _try_docx2pdf_conversion(self, input_path, output_path):
        """Try docx2pdf conversion, return True if successful"""
        self.logger.info("SIMPLE_CONVERT: Trying subprocess docx2pdf...")
        try:
            # Use subprocess to call docx2pdf directly
            result = subprocess.run(
                [
                    sys.executable,
                    "-c",
                    f"from docx2pdf import convert; convert(r'{input_path}', r'{output_path}')",
                ],
                capture_output=True,
                text=True,
                timeout=CONVERSION_TIMEOUT,
            )

            self.logger.info(f"SIMPLE_CONVERT: docx2pdf returncode={result.returncode}")
            self.logger.info(f"SIMPLE_CONVERT: docx2pdf stdout={result.stdout}")
            self.logger.info(f"SIMPLE_CONVERT: docx2pdf stderr={result.stderr}")
            
            if result.returncode == 0 and output_path.exists():
                self.logger.info("SIMPLE_CONVERT: subprocess docx2pdf SUCCESS")
                return True
            else:
                self.logger.error(
                    f"SIMPLE_CONVERT: subprocess docx2pdf failed - returncode={result.returncode}, stderr={result.stderr}"
                )
                return False
        except Exception as e:
            self.logger.error(f"SIMPLE_CONVERT: subprocess docx2pdf error - {e}")
            return False

    def _try_win32com_conversion(self, input_path, output_path):
        """Try win32com conversion, return True if successful"""
        self.logger.info("SIMPLE_CONVERT: Trying subprocess win32com...")
        try:
            win32_script = f"""
import pythoncom
pythoncom.CoInitialize()
import win32com.client as win32
word = win32.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
doc = word.Documents.Open(r"{input_path.absolute()}")
doc.ExportAsFixedFormat(
    OutputFileName=r"{output_path.absolute()}",
    ExportFormat=17,
    OpenAfterExport=False,
    OptimizeFor=0,
    BitmapMissingFonts=True,
    DocStructureTags=True,
    CreateBookmarks=False,
    Range=0
)
doc.Close(SaveChanges=False)
word.Quit()
pythoncom.CoUninitialize()
print("win32com conversion completed")
"""

            result = subprocess.run(
                [sys.executable, "-c", win32_script],
                capture_output=True,
                text=True,
                timeout=CONVERSION_TIMEOUT,
            )

            self.logger.info(f"SIMPLE_CONVERT: win32com returncode={result.returncode}")
            self.logger.info(f"SIMPLE_CONVERT: win32com stdout={result.stdout}")
            self.logger.info(f"SIMPLE_CONVERT: win32com stderr={result.stderr}")
            
            if result.returncode == 0 and output_path.exists():
                self.logger.info("SIMPLE_CONVERT: subprocess win32com SUCCESS")
                return True
            else:
                self.logger.error(
                    f"SIMPLE_CONVERT: subprocess win32com failed - returncode={result.returncode}, stderr={result.stderr}"
                )
                return False
        except Exception as e:
            self.logger.error(f"SIMPLE_CONVERT: subprocess win32com error - {e}")
            return False

    def check_conversion_methods(self):
        # Return what we know should work, sử dụng config
        return {
            "win32com": True,
            "docx2pdf": True,
            "libreoffice": False,
            "supported": True,
            "methods": CONVERSION_METHODS,
            "default_method": DEFAULT_METHOD,
            "backup_method": BACKUP_METHOD
        }
