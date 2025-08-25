from __future__ import annotations

import os
import sys
import subprocess
from pathlib import Path
from tkinter import filedialog, messagebox

class FileHandler:
    def __init__(self, supported_extensions: tuple[str, ...] = ()) -> None:
        self.supported_extensions = tuple(supported_extensions) if supported_extensions else tuple()

    def select_word_file(self, parent=None) -> str | None:
        return filedialog.askopenfilename(
            title="Chọn file Word",
            filetypes=[("Word files", "*.doc *.docx"), ("All files", "*.*")],
            parent=parent,
        )

    def select_excel_file(self, parent=None) -> str | None:
        return filedialog.askopenfilename(
            title="Chọn file Excel",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")],
            parent=parent,
        )

    def open_downloads_folder(self) -> None:
        downloads = Path.home() / "Downloads"
        downloads.mkdir(parents=True, exist_ok=True)
        if os.name == "nt":
            os.startfile(downloads)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.call(["open", str(downloads)])
        else:
            subprocess.call(["xdg-open", str(downloads)])

    def show_message(self, parent, ok: bool, msg: str) -> None:
        if ok:
            messagebox.showinfo("Thành công", msg, parent=parent)
        else:
            messagebox.showerror("Lỗi", msg, parent=parent)
