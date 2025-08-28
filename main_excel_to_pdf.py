# main_excel_to_pdf.py
from __future__ import annotations

import os
import shutil
import tkinter as tk
import sys, subprocess
from tkinter import filedialog
from pathlib import Path
from typing import Optional, Tuple

# Core modules
from src.logging.logger_setup import setup_logger
from src.interface.tkinter_ui import ConverterUI
from src.io.file_handler import FileHandler
from src.converters.excel_to_pdf import excel_to_pdf, is_excel_file

SUPPORTED_EXTENSIONS_EXCEL: Tuple[str, ...] = (".xls", ".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm")


class ExcelApp:
    def __init__(self) -> None:
        # Logger
        self.logger = setup_logger("ExcelApp", level="INFO")

        # Core helpers
        self.fh = FileHandler(supported_extensions=SUPPORTED_EXTENSIONS_EXCEL)

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

    # ----------------------- Back button -----------------------
    def _on_back(self) -> None:
        """Đóng trang hiện tại và mở lại launcher (nếu có)."""
        try:
            main_path = Path(__file__).resolve().parent / "main.py"
            if main_path.exists():
                subprocess.Popen([sys.executable, str(main_path)], cwd=str(main_path.parent))
        finally:
            if self.root:
                self.root.destroy()

    # ----------------------- UI Callbacks -----------------------
    def _on_select(self) -> None:
        """Chọn file Excel (ưu tiên dùng FileHandler)."""
        try:
            path_str = self.fh.select_excel_file(parent=self.root)
        except Exception as e:
            self.logger.exception("Lỗi khi mở hộp thoại chọn file: %s", e)
            if self.ui:
                self.ui.alert_error("Lỗi", "Không mở được hộp thoại chọn file.")
            return

        if not path_str:
            if self.ui:
                self.ui.update_status("ℹ️ Bạn chưa chọn tệp nào.", 0)
            return

        path = Path(path_str)
        if not is_excel_file(str(path)):
            if self.ui:
                self.ui.alert_warning("Sai định dạng", "Hãy chọn tệp Excel hợp lệ (xls, xlsx, xlsm, xlsb, xltx, xltm).")
                self.ui.update_status("⚠️ Tệp không hợp lệ. Hãy chọn lại.", 0)
            return

        self.selected_file = path
        self.temp_pdf_path = None  # reset temp if reselect
        if self.ui:
            self.ui.update_status(f"✅ Đã chọn: {path.name}. Nhấn 'Chuyển sang PDF'.", 10)
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
        """Thực hiện chuyển đổi Excel → PDF, lưu TẠM vào ./outputpdf/."""
        if not self.selected_file:
            if self.ui:
                self.ui.alert_warning("Thiếu tệp", "Hãy chọn tệp Excel trước khi chuyển đổi.")
                self.ui.update_status("⚠️ Chưa có tệp. Vui lòng chọn tệp Excel.", 0)
            return

        src = self.selected_file
        if not is_excel_file(str(src)):
            if self.ui:
                self.ui.alert_warning("Sai định dạng", "Tệp đã chọn không phải Excel hợp lệ.")
                self.ui.update_status("⚠️ Tệp không hợp lệ. Hãy chọn lại.", 0)
            return

        try:
            if self.ui:
                # khóa các nút trong lúc chạy
                self.ui.set_buttons_enabled(select=False, convert=False, open_downloads=False, quit_btn=False)
                self.ui.update_status("🔄 Đang chuyển đổi… (vui lòng đợi)", 25)

            tmp_out = self._make_unique(self.temp_dir / (src.stem + ".pdf"))

            # Thực thi converter → xuất TẠM
            pdf_path_str = excel_to_pdf(str(src), str(tmp_out))
            pdf_path = Path(pdf_path_str) if pdf_path_str else tmp_out
            self.temp_pdf_path = pdf_path

            if self.ui:
                self.ui.update_status(
                    f"✅ Đã tạo bản TẠM: {pdf_path.name} (trong thư mục outputpdf). "
                    "Bây giờ nhấn 'Tải về…' để chọn nơi lưu bản chính.",
                    100
                )
                # đổi nhãn nút bước 3 nếu UI hỗ trợ
                try:
                    self.ui.set_open_downloads_text("Tải về…")
                except Exception:
                    pass
                self.ui.alert_info("Thành công", f"Đã tạo PDF tạm: {pdf_path}")
                # mở lại các nút + bật nút 'Tải về…'
                self.ui.set_buttons_enabled(select=True, convert=True, open_downloads=True, quit_btn=True)
        except Exception as e:
            self.logger.exception("Lỗi khi chuyển đổi Excel → PDF: %s", e)
            if self.ui:
                self.ui.alert_error("Lỗi", f"Không thể chuyển đổi: {e}")
                self.ui.update_status("❌ Lỗi khi chuyển đổi. Hãy thử lại hoặc kiểm tra file Excel.", 0)
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

        # Top bar + nút quay lại
        topbar = tk.Frame(self.root)
        topbar.pack(fill="x", padx=10, pady=(8, 0))
        tk.Button(topbar, text="← Quay lại", command=self._on_back).pack(side="left")

        # Khởi tạo UI chung để đồng bộ giao diện với Word
        self.ui = ConverterUI(
            root=self.root,
            on_select=self._on_select,
            on_convert=self._on_convert,
            # DÙNG callback nút thứ 3 để 'Tải về…' (Save As) thay vì 'Mở Downloads'
            on_open_downloads=self._on_save_as,
            title_text="🔄 Excel → PDF",
            select_button_text="Chọn tệp Excel…",
            convert_button_text="Chuyển sang PDF",
            downloads_hint_text="📥 Bước 3: Nhấn 'Tải về…'",
            supported_extensions=SUPPORTED_EXTENSIONS_EXCEL,
            window_title="Excel → PDF",
            window_size="1000x580",
        )

        # Sau khi tạo UI, cố gắng đổi nhãn nút thứ 3 → 'Tải về…' (nếu UI hỗ trợ)
        try:
            self.ui.set_open_downloads_text("Tải về…")
        except Exception:
            pass

        # Trạng thái ban đầu
        if self.ui:
            self.ui.update_status("✅ Sẵn sàng - Chọn tệp Excel để bắt đầu", 0)
            # chỉ bật nút Convert sau khi có file, nút 'Tải về…' cho phép bấm nhưng sẽ báo nếu chưa có bản tạm
            self.ui.set_buttons_enabled(select=True, convert=False, open_downloads=True, quit_btn=True)

        self.root.mainloop()


def main() -> None:
    ExcelApp().run()


if __name__ == "__main__":
    main()
