# -*- coding: utf-8 -*-
import shutil
from pathlib import Path
from typing import List, Optional, Tuple
import logging

# Lấy config từ src.__init__ (có fallback để an toàn nếu thiếu)
try:
    from .. import (
        SUPPORTED_EXTENSIONS,  # ví dụ: ['.doc', '.docx']
        OUTPUT_PATH,           # thư mục xuất PDF
        DOWNLOADS_PATH         # thư mục Downloads
    )
except Exception:
    SUPPORTED_EXTENSIONS = ['.doc', '.docx']
    OUTPUT_PATH = Path.cwd() / "PDF_Output"
    DOWNLOADS_PATH = Path.home() / "Downloads"


class FileHandler:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.output_folder = Path(OUTPUT_PATH)
        self.downloads_folder = Path(DOWNLOADS_PATH)
        self.temp_dir = self.output_folder / "_temp"

    # ================== Helpers / Queries ==================

    def get_supported_extensions(self) -> List[str]:
        return list(SUPPORTED_EXTENSIONS)

    def is_supported_file(self, file_path: Path) -> bool:
        try:
            return file_path.suffix.lower() in self.get_supported_extensions()
        except Exception:
            return False

    def get_file_info(self, file_path: Path) -> dict:
        p = Path(file_path)
        if not p.exists():
            return {"exists": False, "name": p.name, "path": str(p)}
        stat = p.stat()
        return {
            "exists": True,
            "name": p.name,
            "path": str(p.resolve()),
            "size": stat.st_size,
            "size_mb": round(stat.st_size / (1024 * 1024), 2),
            "extension": p.suffix.lower(),
        }

    # ================== Validation ==================

    def validate_input_file(self, file_path: Path) -> bool:
        """Giữ để tương thích cũ (bool)."""
        p = Path(file_path)
        if not p.exists():
            self.logger.error(f"File not found: {p}")
            return False
        if not self.is_supported_file(p):
            self.logger.error(f"Unsupported file type: {p.suffix}")
            return False
        if p.stat().st_size == 0:
            self.logger.error(f"Empty file: {p}")
            return False
        return True

    def validate_word_file(self, file_path) -> Tuple[bool, str]:
        """API main.py: trả về (is_valid, message)."""
        try:
            p = Path(file_path) if not isinstance(file_path, Path) else file_path
            if not p:
                return False, "Bạn chưa chọn file."
            if not p.exists():
                return False, "File không tồn tại."
            if p.suffix.lower() not in self.get_supported_extensions():
                return False, f"Định dạng không hỗ trợ: {p.suffix}"
            if p.stat().st_size == 0:
                return False, "File trống."
            return True, "OK"
        except Exception as e:
            return False, f"Lỗi khi kiểm tra file: {e}"

    # ================== IO Directories / Paths ==================

    def create_output_directory(self, output_dir: Optional[Path] = None) -> Path:
        if output_dir is None:
            output_dir = self.output_folder
        output_dir.mkdir(parents=True, exist_ok=True)
        self.logger.info(f"Output directory: {output_dir}")
        return output_dir

    def get_unique_filename(self, directory: Path, filename: str) -> Path:
        """Trả về path không trùng tên trong directory."""
        file_path = directory / filename
        counter = 1

        if '.' in filename:
            stem = filename.rsplit('.', 1)[0]
            suffix = '.' + filename.rsplit('.', 1)[1]
        else:
            stem = filename
            suffix = ''

        while file_path.exists():
            file_path = directory / f"{stem}_{counter}{suffix}"
            counter += 1

        return file_path

    def generate_output_path(
        self,
        input_file,
        output_dir: Optional[Path] = None,
        output_format: Optional[str] = None,
        **_
    ) -> Path:
        """API main.py: nhận được output_dir, output_format."""
        p = Path(input_file) if not isinstance(input_file, Path) else input_file
        out_dir = self.create_output_directory(output_dir)

        try:
            from .. import OUTPUT_FORMAT as _DEFAULT_OUT_FMT
        except Exception:
            _DEFAULT_OUT_FMT = "pdf"

        ext = (output_format or _DEFAULT_OUT_FMT or "pdf").strip().lower()
        ext = ext.lstrip(".")
        filename = f"{p.stem}.{ext}"

        return self.get_unique_filename(out_dir, filename)

    # ================== File selection (UI) ==================

    def select_word_file(self):
        """API main.py: mở hộp thoại chọn file Word → Path | None"""
        try:
            from tkinter import filedialog
            path = filedialog.askopenfilename(
                title="Chọn file Word",
                filetypes=[("Word documents", "*.docx *.doc"), ("All files", "*.*")],
            )
            return Path(path) if path else None
        except Exception as e:
            self.logger.error(f"Lỗi khi mở hộp thoại chọn file: {e}")
            return None

    # ================== Copy / Downloads ==================

    def copy_to_downloads(self, source_file: Path) -> Tuple[bool, str, Optional[Path]]:
        """API main.py: trả về (success, message, destination_path)"""
        try:
            source = Path(source_file)
            if not source.exists():
                raise FileNotFoundError(f"File not found: {source}")

            self.downloads_folder.mkdir(parents=True, exist_ok=True)
            destination = self.downloads_folder / source.name

            counter = 1
            while destination.exists():
                destination = self.downloads_folder / f"{source.stem}_{counter}{source.suffix}"
                counter += 1

            shutil.copy2(source, destination)
            self.logger.info(f"Copied file to Downloads: {destination}")
            return True, "OK", destination
        except Exception as e:
            self.logger.error(f"Lỗi copy file: {e}")
            return False, str(e), None

    # ================== Cleanup ==================

    def clean_temp_files(self, temp_dir: Path):
        """Hàm cũ: xóa thư mục tạm cụ thể."""
        try:
            if temp_dir.exists() and temp_dir.is_dir():
                shutil.rmtree(temp_dir)
                self.logger.info(f"Đã xóa thư mục tạm: {temp_dir}")
        except Exception as e:
            self.logger.warning(f"Không thể xóa thư mục tạm {temp_dir}: {e}")

    def cleanup_temp_files(self) -> int:
        """API main.py: không tham số, trả về số item đã xóa."""
        removed = 0

        try:
            if self.output_folder.exists():
                for p in self.output_folder.glob("*.tmp"):
                    try:
                        p.unlink()
                        removed += 1
                    except Exception:
                        pass
        except Exception:
            pass

        try:
            if self.temp_dir.exists() and self.temp_dir.is_dir():
                try:
                    for _ in self.temp_dir.rglob("*"):
                        removed += 1
                except Exception:
                    pass
                shutil.rmtree(self.temp_dir, ignore_errors=True)
        except Exception:
            pass

        return removed
