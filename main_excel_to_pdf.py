# -*- coding: utf-8 -*-
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
try:
    import ttkbootstrap as tb
    THEME = "sandstone"
except Exception:
    tb = None
from src.converters.excel_to_pdf import excel_to_pdf, is_excel_file

def pick_excel():
    return filedialog.askopenfilename(
        title="Chọn file Excel",
        filetypes=[("Excel files", "*.xlsx *.xls *.xlsm *.xlsb *.xltx *.xltm"), ("All files", "*.*")]
    )

def ensure_downloads():
    home = os.path.expanduser("~")
    for p in (os.path.join(home, "Downloads"), os.path.join(home, "Download"), home):
        if os.path.isdir(p): return p
    return home

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel → PDF (3 bước)")
        self.file_var = tk.StringVar()
        self.save_dir = tk.StringVar(value=ensure_downloads())
        self.out_var = tk.StringVar()
        self.status = tk.StringVar(value="Chưa chọn file.")
        self._build()

    def _build(self):
        frm = tk.Frame(self.root, padx=12, pady=12)
        frm.pack(fill=tk.BOTH, expand=True)

        # B1: chọn file
        tk.Label(frm, text="BƯỚC 1: CHỌN FILE EXCEL", font=("Segoe UI", 10, "bold")).pack(anchor=tk.W)
        r1 = tk.Frame(frm); r1.pack(fill=tk.X, pady=6)
        tk.Entry(r1, textvariable=self.file_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,8))
        tk.Button(r1, text="Chọn…", command=self.on_pick).pack(side=tk.LEFT)

        # B1.5: nơi lưu
        tk.Label(frm, text="THƯ MỤC LƯU PDF", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, pady=(8,0))
        r1b = tk.Frame(frm); r1b.pack(fill=tk.X, pady=6)
        tk.Entry(r1b, textvariable=self.save_dir).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,8))
        tk.Button(r1b, text="Chọn nơi lưu…", command=self.on_pick_dir).pack(side=tk.LEFT)

        # B2: chuyển đổi
        tk.Label(frm, text="BƯỚC 2: CHUYỂN ĐỔI", font=("Segoe UI", 10, "bold")).pack(anchor=tk.W, pady=(10,0))
        r2 = tk.Frame(frm); r2.pack(fill=tk.X, pady=6)
        tk.Button(r2, text="Chuyển sang PDF", command=self.on_convert).pack(side=tk.LEFT)

        # B3: mở thư mục
        tk.Label(frm, text="BƯỚC 3: TẢI XUỐNG (mở thư mục chứa PDF)", font=("Segoe UI", 10, "bold")).pack(anchor=tk.W, pady=(10,0))
        r3 = tk.Frame(frm); r3.pack(fill=tk.X, pady=6)
        tk.Entry(r3, textvariable=self.out_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,8))
        tk.Button(r3, text="Mở thư mục", command=self.on_open_dir).pack(side=tk.LEFT)

        tk.Label(frm, textvariable=self.status, fg="#006400").pack(anchor=tk.W, pady=(10,0))

    def on_pick(self):
        p = pick_excel()
        if p:
            self.file_var.set(p)
            self.status.set("Đã chọn file. Nhấn 'Chuyển sang PDF'.")

    def on_pick_dir(self):
        d = filedialog.askdirectory(title="Chọn thư mục lưu PDF", initialdir=self.save_dir.get() or ensure_downloads())
        if d:
            self.save_dir.set(d)

    def on_convert(self):
        src = self.file_var.get().strip()
        if not src:
            messagebox.showwarning("Thiếu file", "Hãy chọn file Excel trước.")
            return
        if not is_excel_file(src):
            messagebox.showerror("Sai định dạng", "File không phải Excel hợp lệ.")
            return
        try:
            out_dir = self.save_dir.get().strip() or ensure_downloads()
            os.makedirs(out_dir, exist_ok=True)
            out_path = os.path.join(out_dir, os.path.splitext(os.path.basename(src))[0] + ".pdf")

            self.status.set("Đang chuyển đổi…")
            self.root.update()
            pdf_path = excel_to_pdf(src, out_path)

            self.out_var.set(pdf_path)
            self.status.set("Xong! File PDF đã lưu.")
            # Thông báo rõ nếu tên đã tự đổi do file bị khóa
            if os.path.normpath(pdf_path) != os.path.normpath(out_path):
                messagebox.showinfo("Đã lưu (đổi tên)",
                    f"File đích đang bị mở/khóa nên đã lưu thành:\n{pdf_path}")
            else:
                messagebox.showinfo("Thành công", f"Đã lưu PDF:\n{pdf_path}")
        except Exception as e:
            messagebox.showerror("Lỗi chuyển đổi", str(e))
            self.status.set("Có lỗi xảy ra.")

    def on_open_dir(self):
        p = self.out_var.get().strip()
        if not p:
            messagebox.showinfo("Chưa có file", "Bạn cần chuyển đổi trước.")
            return
        folder = os.path.dirname(p) or os.getcwd()
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)
            elif sys.platform == "darwin":
                os.system(f'open "{folder}"')
            else:
                os.system(f'xdg-open "{folder}"')
        except Exception:
            messagebox.showwarning("Không mở được", folder)

def main():
    root = tb.Window(themename="sandstone") if tb is not None else tk.Tk()
    App(root)
    root.minsize(720, 320)
    root.mainloop()

if __name__ == "__main__":
    main()
