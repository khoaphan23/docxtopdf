"""
Word → PDF (clean main compatible with new ConverterUI)
"""
from __future__ import annotations

import tkinter as tk
from pathlib import Path
from typing import Optional, Tuple

# Core modules
from src.logging.logger_setup import setup_logger
from src.interface.tkinter_ui import ConverterUI
from src.converters.word_to_pdf import DocumentConverter
from src.io.file_handler import FileHandler

SUPPORTED_EXTENSIONS_WORD: Tuple[str, ...] = (".doc", ".docx")


class WordApp:
    """
    Ứng dụng Word → PDF gọn – chuẩn, tương thích ConverterUI mới (không dùng `callbacks=`).
    Luồng: Chọn file → Convert (lưu thẳng vào Downloads) → Mở thư mục Downloads.
    """

    def __init__(self) -> None:
        # Logger
        self.logger = setup_logger("WordApp", level="INFO")

        # Core
        self.converter = DocumentConverter()
        self.fh = FileHandler()

        # State
        self.root: Optional[tk.Tk] = None
        self.ui: Optional[ConverterUI] = None
        self.selected_file: Optional[Path] = None

    # ----------------------- UI Callbacks -----------------------
    def _on_select(self) -> None:
        """Chọn file Word (ưu tiên hàm trong FileHandler; fallback nếu không có)."""
        try:
            # prefer project-specific selector if present
            if hasattr(self.fh, "select_word_file"):
                path = self.fh.select_word_file(parent=self.root)
            else:
                # Fallback dialog (rất ít khi cần nếu FileHandler đã có)
                from tkinter import filedialog
                path = filedialog.askopenfilename(
                    title="Chọn file Word",
                    filetypes=[("Word files", "*.doc *.docx"), ("All files", "*.*")],
                    parent=self.root,
                )

            if not path:
                # user cancelled
                if self.ui:
                    self.ui.update_status("✅ Sẵn sàng - Hãy chọn file Word để bắt đầu", 0)
                return

            p = Path(path)

            # validate nếu FileHandler có sẵn validate_word_file
            if hasattr(self.fh, "validate_word_file"):
                ok, msg = self.fh.validate_word_file(p)  # type: ignore[attr-defined]
                if not ok:
                    if self.ui:
                        self.ui.alert_error("File không hợp lệ", msg)
                        self.ui.update_status("❌ File không hợp lệ", 0)
                    return

            self.selected_file = p
            if self.ui:
                self.ui.update_status(f"✅ Đã chọn: {p.name}", 15)

            self.logger.info(f"Selected file: {p}")

        except Exception as e:
            self.logger.exception("Lỗi khi chọn file")
            if self.ui:
                self.ui.alert_error("Lỗi chọn file", f"{e}")
                self.ui.update_status("❌ Lỗi chọn file", 0)

    def _on_convert(self) -> None:
        """Chuyển file Word → PDF: lưu thẳng vào Downloads (tránh bước 'tải xuống')."""
        if not self.selected_file:
            if self.ui:
                self.ui.alert_warning("Chưa chọn file", "Vui lòng chọn file Word trước khi chuyển!")
            return

        if self.ui:
            self.ui.set_buttons_enabled(select=False, convert=False, open_downloads=False)
            self.ui.update_status("🔄 Đang chuyển…", 35)

        try:
            ok_avail, info = self.converter.is_available()
            if not ok_avail:
                if self.ui:
                    self.ui.alert_error("Thiếu thư viện", info)
                    self.ui.update_status("⚠️ Chưa sẵn sàng", 0)
                return

            # Lưu trực tiếp vào Downloads (đã có tránh trùng tên)
            ok, msg, out_path = self.converter.convert_and_save_to_downloads(self.selected_file)

            if ok:
                if self.ui:
                    self.ui.alert_info("Thành công", msg)
                    self.ui.update_status("🎉 Chuyển đổi thành công", 100)
            else:
                if self.ui:
                    self.ui.alert_error("Chuyển đổi thất bại", msg)
                    self.ui.update_status("❌ Chuyển đổi thất bại", 25)

            self.logger.info(f"Convert result: ok={ok}, msg={msg}, out={out_path}")

        except Exception as e:
            self.logger.exception("Lỗi không mong muốn khi chuyển đổi")
            if self.ui:
                self.ui.alert_error("Lỗi chuyển đổi", f"{e}")
                self.ui.update_status("❌ Lỗi chuyển đổi", 25)

        finally:
            if self.ui:
                self.ui.set_buttons_enabled(select=True, convert=True, open_downloads=True)

    def _on_open_downloads(self) -> None:
        """Mở thư mục Downloads."""
        try:
            self.fh.open_downloads_folder()
        except Exception as e:
            self.logger.exception("Lỗi mở thư mục Downloads")
            if self.ui:
                self.ui.alert_error("Lỗi", f"Không mở được Downloads: {e}")

    # ----------------------- Run -----------------------
    def run(self) -> None:
        self.root = tk.Tk()

        # Tạo UI theo đúng signature mới (3 callback positional)
        self.ui = ConverterUI(
            self.root,
            self._on_select,
            self._on_convert,
            self._on_open_downloads,
            title_text="🔄 Word to PDF Converter",
            select_button_text="Chọn file Word",
            convert_button_text="Chuyển sang PDF",
            downloads_hint_text="💾 File PDF sẽ được lưu vào thư mục Downloads",
            supported_extensions=SUPPORTED_EXTENSIONS_WORD,
        )

        # Trạng thái ban đầu
        ok, info = self.converter.is_available()
        if self.ui:
            self.ui.update_status(f"{'✅' if ok else '⚠️'} {info} - Hãy chọn file Word", 0)

        self.root.mainloop()


def main() -> None:
    WordApp().run()


if __name__ == "__main__":
    main()
