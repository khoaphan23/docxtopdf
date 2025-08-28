
from __future__ import annotations

import os
import shutil
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
from typing import Optional, Tuple
from src.converters.word_to_pdf import word_to_pdf, is_word_file


# Core modules
from src.logging.logger_setup import setup_logger
from src.interface.tkinter_ui import ConverterUI
from src.io.file_handler import FileHandler

# Converters — adjust imports to your actual module names if different
try:
    from src.converters.word_to_pdf import word_to_pdf, is_word_file  # preferred if exists
except Exception:
    # Fallback to docx_to_pdf naming if your project uses that
    from src.converters.docx_to_pdf import docx_to_pdf as word_to_pdf  # type: ignore
    try:
        from src.converters.docx_to_pdf import is_docx_file as is_word_file  # type: ignore
    except Exception:
        # last resort — simple extension check
        def is_word_file(p: str) -> bool:
            return str(p).lower().endswith((".doc", ".docx"))

SUPPORTED_EXTENSIONS_WORD: Tuple[str, ...] = (".doc", ".docx")


class WordApp:
    def __init__(self) -> None:
        # Logger
        self.logger = setup_logger("WordApp", level="INFO")

        # Core helpers
        self.fh = FileHandler(supported_extensions=SUPPORTED_EXTENSIONS_WORD)

        # State
        self.root: Optional[tk.Tk] = None
        self.ui: Optional[ConverterUI] = None
        self.selected_file: Optional[Path] = None
        self.temp_pdf_path: Optional[Path] = None

        # temp directory inside project (./outputpdf)
        self.temp_dir = Path(__file__).parent / "outputpdf"
        try:
            self.temp_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            # fallback: system temp if cannot create project temp
            self.logger.warning("Không thể tạo thư mục ./outputpdf, dùng thư mục tạm của hệ thống. Lý do: %s", e)
            self.temp_dir = Path(os.getenv("TEMP", Path.home())) / "outputpdf_tmp"
            self.temp_dir.mkdir(parents=True, exist_ok=True)

    # ----------------------- UI Callbacks -----------------------
    def _on_select(self) -> None:
        try:
            path_str = self.fh.select_file(parent=self.root, title="Chọn tệp Word…")
        except Exception as e:
            self.logger.warning("Select via FileHandler lỗi: %s -> dùng fallback tk filedialog", e)
            from tkinter import filedialog as fd
            import os
            patterns = ("*.docx", "*.doc")
            path_str = fd.askopenfilename(
                parent=self.root,
                title="Chọn tệp Word…",
                filetypes=[("Word documents", patterns), ("All files", "*.*")],
                initialdir=os.path.expanduser("~"),
            )

        path = Path(path_str)
        if not is_word_file(str(path)):
            if self.ui:
                self.ui.alert_warning("Sai định dạng", "Hãy chọn tệp Word hợp lệ (doc, docx).")
                self.ui.update_status("⚠️ Tệp không hợp lệ. Hãy chọn lại.", 0)
            return

        self.selected_file = path
        self.temp_pdf_path = None  # reset temp if reselect
        if self.ui:
            self.ui.update_status(f"✅ Đã chọn: {path.name}. Nhấn 'Chuyển sang PDF' để tạo bản tạm.", 10)
            self.ui.set_buttons_enabled(convert=True)

    def _make_unique(self, dest: Path) -> Path:
        """Nếu file tồn tại, thêm hậu tố _1, _2,... để tránh ghi đè."""
        if not dest.exists():
            return dest
        stem, suffix = dest.stem, dest.suffix
        i = 1
        while True:
            p = dest.with_name(f"{stem}_{i}{suffix}")
            if not p.exists():
                return p
            i += 1

    def _on_convert(self) -> None:
        """Thực hiện chuyển đổi Word → PDF, lưu TẠM vào ./outputpdf/."""
        if not self.selected_file:
            if self.ui:
                self.ui.alert_warning("Thiếu tệp", "Hãy chọn tệp Word trước khi chuyển đổi.")
                self.ui.update_status("⚠️ Chưa có tệp. Vui lòng chọn tệp Word.", 0)
            return

        src = self.selected_file
        if not is_word_file(str(src)):
            if self.ui:
                self.ui.alert_warning("Sai định dạng", "Tệp đã chọn không phải Word hợp lệ.")
                self.ui.update_status("⚠️ Tệp không hợp lệ. Hãy chọn lại.", 0)
            return

        try:
            if self.ui:
                # khóa các nút trong lúc chạy
                self.ui.set_buttons_enabled(select=False, convert=False, open_downloads=False, quit_btn=False)
                self.ui.update_status("🔄 Đang chuyển đổi… (vui lòng đợi)", 25)

            tmp_out = self._make_unique(self.temp_dir / (src.stem + ".pdf"))

            # Thực thi converter → xuất TẠM
            pdf_path_str = word_to_pdf(str(src), str(tmp_out))
            pdf_path = Path(pdf_path_str) if pdf_path_str else tmp_out
            self.temp_pdf_path = pdf_path

            if self.ui:
                self.ui.update_status(
                    f"✅ Đã tạo bản TẠM: {pdf_path.name} (trong thư mục outputpdf). "
                    "Bây giờ nhấn 'Tải về…' để chọn nơi lưu bản chính.",
                    100
                )
                try:
                    # Nếu ConverterUI có API đổi nhãn nút thứ 3, ta đổi thành 'Tải về…'
                    self.ui.set_open_downloads_text("Tải về…")
                except Exception:
                    pass
                self.ui.alert_info("Thành công", f"Đã tạo PDF tạm: {pdf_path}")
                # mở lại các nút + bật nút 'Tải về…'
                self.ui.set_buttons_enabled(select=True, convert=True, open_downloads=True, quit_btn=True)
        except Exception as e:
            self.logger.exception("Lỗi khi chuyển đổi Word → PDF: %s", e)
            if self.ui:
                self.ui.alert_error("Lỗi", f"Không thể chuyển đổi: {e}")
                self.ui.update_status("❌ Lỗi khi chuyển đổi. Hãy thử lại hoặc kiểm tra file Word.", 0)
                self.ui.set_buttons_enabled(select=True, convert=True, open_downloads=True, quit_btn=True)

    def _on_save_as(self) -> None:
        """Bước 3: Chọn nơi 'Tải về…' (Save As) từ bản PDF tạm."""
        if not self.temp_pdf_path or not self.temp_pdf_path.exists():
            if self.ui:
                self.ui.alert_warning(
                    "Chưa có bản tạm",
                    "Chưa có PDF tạm để tải về. Hãy nhấn 'Chuyển sang PDF' trước."
                )
                self.ui.update_status("ℹ️ Hãy chuyển sang PDF để tạo bản tạm trước khi tải về.", 0)
            return

        try:
            initialfile = self.temp_pdf_path.name
            final_path_str = filedialog.asksaveasfilename(
                parent=self.root,
                defaultextension=".pdf",
                filetypes=[("PDF", "*.pdf")],
                initialfile=initialfile,
                title="Chọn nơi lưu tệp PDF…",
            )
            if not final_path_str:
                if self.ui:
                    self.ui.update_status("ℹ️ Bạn đã hủy thao tác 'Tải về…'.", 0)
                return

            final_path = Path(final_path_str)
            # đảm bảo tồn tại thư mục đích
            final_path.parent.mkdir(parents=True, exist_ok=True)

            # Sao chép bản tạm → đích
            shutil.copyfile(self.temp_pdf_path, final_path)

            if self.ui:
                self.ui.update_status(f"✅ Đã lưu về: {final_path}", 100)
                self.ui.alert_info("Hoàn tất", f"Đã lưu PDF: {final_path}")
        except Exception as e:
            self.logger.exception("Lỗi khi 'Tải về…': %s", e)
            if self.ui:
                self.ui.alert_error("Lỗi", f"Không thể lưu về: {e}")
                self.ui.update_status("❌ Lỗi khi lưu về. Hãy thử lại.", 0)

    # ----------------------- App life-cycle -----------------------
    def run(self) -> None:
        self.root = tk.Tk()

        # Khởi tạo UI chung để đồng bộ giao diện
        self.ui = ConverterUI(
            root=self.root,
            on_select=self._on_select,
            on_convert=self._on_convert,
            # DÙNG callback nút thứ 3 để 'Tải về…' (Save As)
            on_open_downloads=self._on_save_as,
            title_text="🔄 Word → PDF",
            select_button_text="Chọn tệp Word…",
            convert_button_text="Chuyển sang PDF",
            downloads_hint_text="📥 Bước 3: Nhấn 'Tải về…'",
            supported_extensions=SUPPORTED_EXTENSIONS_WORD,
            window_title="Word → PDF",
            window_size="700x440",
        )

        # cố gắng đổi nhãn nút thứ 3 → 'Tải về…' (nếu UI hỗ trợ)
        try:
            self.ui.set_open_downloads_text("Tải về…")
        except Exception:
            pass

        # Trạng thái ban đầu
        if self.ui:
            self.ui.update_status("✅ Sẵn sàng - Chọn tệp Word để bắt đầu", 0)
            self.ui.set_buttons_enabled(select=True, convert=False, open_downloads=True, quit_btn=True)

        self.root.mainloop()


def main() -> None:
    WordApp().run()


if __name__ == "__main__":
    main()
