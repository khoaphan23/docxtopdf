# main_image_to_pdf.py
from __future__ import annotations

import os
import shutil
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
from typing import Optional, Tuple

# Core
from src.logging.logger_setup import setup_logger
from src.interface.tkinter_ui import ConverterUI

# FileHandler (giống Word/Excel). Nếu thiếu thì fallback dùng filedialog
try:
    from src.io.file_handler import FileHandler
except Exception:
    FileHandler = None  # fallback

# Converter ảnh -> PDF
from src.converters.image_to_pdf import image_to_pdf, is_image_file

# Giống style 2 cái kia: dùng pattern *.ext để hiển thị trên UI
SUPPORTED_PATTERNS: Tuple[str, ...] = ("*.png", "*.jpg", "*.jpeg", "*.bmp", "*.tif", "*.tiff", "*.webp")


class ImageToPDFApp:
    def __init__(self) -> None:
        self.logger = setup_logger("image_to_pdf")
        self.root: Optional[tk.Tk] = None
        self.ui: Optional[ConverterUI] = None

        self.fh = FileHandler(supported_extensions=SUPPORTED_PATTERNS) if FileHandler else None

        self.selected_file: Optional[Path] = None
        self.temp_pdf_path: Optional[Path] = None

        # Thư mục tạm đồng bộ như Word/Excel
        self.temp_dir = Path(__file__).resolve().parent / "outputpdf"
        try:
            self.temp_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            self.logger.warning("Không thể tạo ./outputpdf, dùng thư mục tạm hệ thống: %s", e)
            self.temp_dir = Path(os.getenv("TEMP", Path.home())) / "outputpdf_tmp"
            self.temp_dir.mkdir(parents=True, exist_ok=True)

    # ===== Bước 1: Chọn file =====
    def _on_select(self) -> None:
        try:
            if self.fh and hasattr(self.fh, "select_image_file"):
                path_str = self.fh.select_image_file(parent=self.root)
            else:
                path_str = filedialog.askopenfilename(
                    title="Chọn tệp Ảnh…",
                    filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp"),
                               ("All files", "*.*")],
                    initialdir=os.path.expanduser("~"),
                )
        except Exception as e:
            self.logger.warning("Chọn file lỗi, fallback: %s", e)
            path_str = filedialog.askopenfilename(
                title="Chọn tệp Ảnh…",
                filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp"),
                           ("All files", "*.*")],
                initialdir=os.path.expanduser("~"),
            )

        if not path_str:
            return

        path = Path(path_str)
        if not is_image_file(str(path)):
            if self.ui:
                self.ui.alert_warning("Sai định dạng", "Hãy chọn ảnh (PNG/JPG/JPEG/BMP/TIF/TIFF/WEBP).")
                self.ui.update_status("⚠️ Tệp không hợp lệ. Hãy chọn lại.", 0)
            return

        self.selected_file = path
        if self.ui:
            self.ui.update_status(f"✅ Đã chọn: {path.name}", 10)
            self.ui.set_buttons_enabled(select=True, convert=True, open_downloads=True, quit_btn=True)

    # ===== Bước 2: Chuyển & LƯU TẠM vào ./outputpdf =====
    def _on_convert(self) -> None:
        if not self.selected_file:
            if self.ui:
                self.ui.alert_warning("Thiếu tệp", "Hãy chọn ảnh trước khi chuyển.")
                self.ui.update_status("⚠️ Chưa có tệp. Hãy chọn ảnh.", 0)
            return

        src = self.selected_file
        if not is_image_file(str(src)):
            if self.ui:
                self.ui.alert_warning("Sai định dạng", "Tệp đã chọn không phải ảnh hợp lệ.")
                self.ui.update_status("⚠️ Tệp không hợp lệ. Hãy chọn lại.", 0)
            return

        try:
            if self.ui:
                self.ui.set_buttons_enabled(select=False, convert=False, open_downloads=False, quit_btn=False)
                self.ui.update_status("⏳ Đang chuyển sang PDF…", 30)

            out_name = src.with_suffix(".pdf").name
            temp_out = self.temp_dir / out_name

            # LƯU TẠM vào ./outputpdf
            pdf_path = image_to_pdf(str(src), str(temp_out))
            self.temp_pdf_path = Path(pdf_path)

            # --- THÔNG BÁO RÕ RÀNG NHƯ YÊU CẦU ---
            if self.ui:
                self.ui.update_status(
                    f"✅ Chuyển thành công! ĐÃ LƯU vào: {temp_out.name} (thư mục ./outputpdf). "
                    "Bấm 'Tải về…' để chọn nơi lưu cuối.",
                    90
                )
                # popup thông báo
                self.ui.alert_info(
                    "Chuyển thành công",
                    f"PDF đã  vào:\n{temp_out}\n\n"
                    "Đây CHƯA phải nơi lưu cuối. Hãy bấm 'Tải về…' để chọn thư mục đích."
                )
                self.ui.set_buttons_enabled(select=True, convert=True, open_downloads=True, quit_btn=True)

        except Exception as e:
            self.logger.exception("Lỗi khi chuyển ảnh sang PDF: %s", e)
            if self.ui:
                self.ui.alert_error("Lỗi", f"Không thể chuyển sang PDF: {e}")
                self.ui.update_status("❌ Lỗi khi chuyển. Hãy thử lại.", 0)
                self.ui.set_buttons_enabled(select=True, convert=True, open_downloads=True, quit_btn=True)

    # ===== Bước 3: Tải về (chọn nơi lưu cuối) =====
    def _on_open_downloads(self) -> None:
        if not self.temp_pdf_path or not self.temp_pdf_path.exists():
            if self.ui:
                self.ui.alert_warning("Chưa có PDF", "Hãy bấm 'Chuyển sang PDF' trước.")
                self.ui.update_status("⚠️ Chưa có PDF tạm. Hãy chuyển trước.", 0)
            return

        try:
            final_path_str = filedialog.asksaveasfilename(
                parent=self.root,
                title="Chọn nơi lưu PDF…",
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                initialfile=self.temp_pdf_path.name,
                initialdir=str(Path.home() / "Downloads"),
            )
            if not final_path_str:
                return

            final_path = Path(final_path_str)
            final_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(str(self.temp_pdf_path), str(final_path))

            if self.ui:
                self.ui.update_status(f"✅ Đã lưu về: {final_path}", 100)
                self.ui.alert_info("Hoàn tất", f"Đã lưu PDF: {final_path}")
        except Exception as e:
            self.logger.exception("Lỗi khi lưu về: %s", e)
            if self.ui:
                self.ui.alert_error("Lỗi", f"Không thể lưu về: {e}")
                self.ui.update_status("❌ Lỗi khi lưu về. Hãy thử lại.", 0)

    def run(self) -> None:
        self.root = tk.Tk()
        self.ui = ConverterUI(
            self.root,
            on_select=self._on_select,
            on_convert=self._on_convert,
            on_open_downloads=self._on_open_downloads,
            title_text="🖼️ Image → PDF",
            select_button_text="Chọn ảnh",
            convert_button_text="Chuyển sang PDF",
            downloads_hint_text="📥 Bước 3: Tải về – chọn nơi lưu cuối",
            supported_extensions=SUPPORTED_PATTERNS,
            window_title="Image to PDF Converter",
            window_size="680x420",
        )

        if self.ui:
            self.ui.update_status("✅ Sẵn sàng - Chọn tệp Ảnh để bắt đầu", 0)
            self.ui.set_buttons_enabled(select=True, convert=False, open_downloads=True, quit_btn=True)

        self.root.mainloop()


if __name__ == "__main__":
    ImageToPDFApp().run()
