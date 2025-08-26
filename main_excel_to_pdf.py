# -*- coding: utf-8 -*-
import os
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox

try:
    import ttkbootstrap as tb
    from ttkbootstrap.constants import *
    THEME_OK = True
except Exception:
    # Fallback nếu chưa cài ttkbootstrap -> dùng ttk chuẩn
    import tkinter.ttk as tb
    THEME_OK = False

from src.converters.excel_to_pdf import excel_to_pdf, is_excel_file

def get_downloads_dir() -> str:
    home = os.path.expanduser("~")
    dl = os.path.join(home, "Downloads")
    if not os.path.isdir(dl):
        os.makedirs(dl, exist_ok=True)
    return dl

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to PDF Converter")
        self.root.geometry("740x520")
        try:
            self.root.iconbitmap(default='')
        except Exception:
            pass

        container = tb.Frame(self.root, padding=10)
        container.pack(fill="both", expand=True)

        title = tb.Label(container, text="📊 Excel to PDF Converter", font=("Segoe UI", 16, "bold"))
        title.pack(anchor="w", pady=(0, 8))

        # ===== BƯỚC 1 =====
        self.card1 = self._card(container)
        self._card_header(self.card1, "📁 Bước 1: Chọn tệp")
        self.lbl_support = tb.Label(self.card1, text="Hỗ trợ: .xls, .xlsx, .xlsm, .xlsb", foreground="#555")
        self.lbl_support.pack(anchor="w", padx=10, pady=(0, 6))

        self.path_var = tk.StringVar(value="Chưa chọn file...")
        self.lbl_path = tb.Label(self.card1, textvariable=self.path_var, wraplength=650)
        self.lbl_path.pack(anchor="w", padx=10, pady=(0, 8))

        self.btn_choose = tb.Button(self.card1, text="Chọn file Excel", command=self.choose_file, width=22)
        self.btn_choose.pack(anchor="w", padx=10, pady=(0, 10))

        # ===== BƯỚC 2 =====
        self.card2 = self._card(container)
        self._card_header(self.card2, "⚙️ Bước 2: Chuyển đổi")

        self.btn_convert = tb.Button(self.card2, text="Chuyển sang PDF", command=self.convert, width=22, state="disabled")
        self.btn_convert.pack(anchor="w", padx=10, pady=(4, 6))

        self.progress = tb.Progressbar(self.card2, mode="indeterminate")
        self.progress.pack(fill="x", padx=10, pady=(0, 8))

        self.status_var = tk.StringVar(value=" ")
        self.lbl_status = tb.Label(self.card2, textvariable=self.status_var, foreground="#555", wraplength=650)
        self.lbl_status.pack(anchor="w", padx=10, pady=(0, 6))

        # ===== BƯỚC 3 =====
        self.card3 = self._card(container)
        self._card_header(self.card3, "⬇️ Bước 3: Tải xuống")

        self.lbl_note = tb.Label(self.card3, text="File PDF sẽ được lưu vào thư mục Downloads", foreground="#555")
        self.lbl_note.pack(anchor="w", padx=10, pady=(0, 8))

        bwrap = tb.Frame(self.card3)
        bwrap.pack(anchor="w", padx=10, pady=(0, 10))

        self.btn_open_pdf = tb.Button(bwrap, text="Mở file PDF", command=self.open_pdf, width=20, state="disabled")
        self.btn_open_pdf.grid(row=0, column=0, padx=(0, 10))

        self.btn_open_downloads = tb.Button(bwrap, text="Mở thư mục Downloads", command=self.open_downloads, width=22)
        self.btn_open_downloads.grid(row=0, column=1)

        # internal state
        self.last_output_path = None

    # helpers
    def _card(self, parent):
        f = tb.Labelframe(parent)
        f.pack(fill="x", expand=False, pady=6)
        return f

    def _card_header(self, frame, text):
        lbl = tb.Label(frame, text=text, font=("Segoe UI", 11, "bold"))
        lbl.pack(anchor="w", padx=10, pady=6)

    # actions
    def choose_file(self):
        filetypes = [
            ("Excel files", "*.xlsx *.xls *.xlsm *.xlsb *.xltx *.xltm"),
            ("All files", "*.*"),
        ]
        path = filedialog.askopenfilename(title="Chọn file Excel", filetypes=filetypes)
        if not path:
            return
        if not is_excel_file(path):
            messagebox.showerror("Lỗi", "Vui lòng chọn đúng file Excel (.xls/.xlsx/...)")
            return

        self.path_var.set(path)
        self.status_var.set(" ")
        self.btn_convert["state"] = "normal"
        self.last_output_path = None
        self.btn_open_pdf["state"] = "disabled"

    def open_downloads(self):
        os.startfile(get_downloads_dir())

    def open_pdf(self):
        if self.last_output_path and os.path.isfile(self.last_output_path):
            os.startfile(self.last_output_path)
        else:
            messagebox.showwarning("Chưa có file", "Hãy chuyển đổi trước, rồi mới mở file PDF.")

    def convert(self):
        src = self.path_var.get()
        if not os.path.isfile(src):
            messagebox.showwarning("Chưa chọn file", "Vui lòng chọn file Excel trước.")
            return

        downloads = get_downloads_dir()
        base = os.path.splitext(os.path.basename(src))[0] + ".pdf"
        dst = os.path.join(downloads, base)

        # chạy nền
        self.btn_convert["state"] = "disabled"
        self.progress.start(10)
        self.status_var.set("Đang chuyển đổi...")
        self.last_output_path = None
        self.btn_open_pdf["state"] = "disabled"

        def _work():
            try:
                out = excel_to_pdf(src, dst)
                self.root.after(0, lambda: self.on_success(out))
            except Exception as e:
                tb = traceback.format_exc()
                self.root.after(0, lambda: self.on_error(str(e), tb))
            finally:
                self.root.after(0, self._done)

        threading.Thread(target=_work, daemon=True).start()

    def _done(self):
        self.progress.stop()
        self.btn_convert["state"] = "normal"

    def on_success(self, path):
        self.status_var.set(f"✔ Hoàn thành: {path}")
        self.last_output_path = path
        self.btn_open_pdf["state"] = "normal"
        try:
            self.root.bell()
        except Exception:
            pass
        messagebox.showinfo("Thành công", f"Đã xuất PDF:\n{path}\n\nBạn có thể bấm 'Mở file PDF' hoặc 'Mở thư mục Downloads'.")

    def on_error(self, msg, detail=""):
        self.status_var.set(f"✖ Lỗi: {msg}")
        messagebox.showerror("Lỗi", f"{msg}\n\nChi tiết:\n{detail}")

def main():
    if THEME_OK:
        root = tb.Window(themename="sandstone")
    else:
        root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
