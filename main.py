import sys
import subprocess
import tkinter as tk
from tkinter import messagebox
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
PY = sys.executable  # dùng đúng Python/venv hiện tại

CANDIDATES = {
    "Word → PDF": ["main_word_to_pdf.py", "word_to_pdf_main.py"],
    "Excel → PDF": ["main_excel_to_pdf.py", "excel_to_pdf_main.py"],
    "Ảnh → PDF":   ["main_image_to_pdf.py", "image_to_pdf.py", "img_to_pdf.py"],
}

def find_script(name_list):
    for name in name_list:
        p = PROJECT_ROOT / name
        if p.exists():
            return p
    return None

def launch(name_list):
    script = find_script(name_list)
    if not script:
        messagebox.showerror(
            "Không tìm thấy",
            "Không tìm thấy file:\n" + "\n".join(name_list) +
            "\nĐặt 1 trong các file này cạnh main.py."
        )
        return
    try:
        subprocess.Popen([PY, str(script)], cwd=str(PROJECT_ROOT))
    except Exception as e:
        messagebox.showerror("Lỗi chạy chương trình", f"{script.name}\n\n{e}")

def build_ui():
    root = tk.Tk()
    root.title("DOCXTOPDF Launcher")
    root.geometry("420x240")

    tk.Label(root, text="Chọn chức năng", font=("Segoe UI", 14, "bold"), pady=12).pack()

    style = {"width": 26, "height": 2, "cursor": "hand2"}

    def launch_and_close(name_list):
        script = find_script(name_list)
        if not script:
            messagebox.showerror(
                "Không tìm thấy",
                "Không tìm thấy file:\n" + "\n".join(name_list),
                parent=root
            )
            return
        # Ẩn cửa sổ launcher ngay khi mở app con
        root.withdraw()
        try:
            subprocess.Popen([PY, str(script)], cwd=str(PROJECT_ROOT))
        except Exception as e:
            # Nếu mở thất bại, hiện lại launcher và báo lỗi
            root.deiconify()
            messagebox.showerror("Lỗi chạy chương trình", f"{script.name}\n\n{e}", parent=root)
            return
        # Đóng hẳn launcher sau khi spawn app con (tránh 2 cửa sổ)
        root.after(100, root.destroy)

    tk.Button(root, text="📝 Word → PDF",
              command=lambda: launch_and_close(CANDIDATES["Word → PDF"]), **style).pack(pady=5)
    tk.Button(root, text="📈 Excel → PDF",
              command=lambda: launch_and_close(CANDIDATES["Excel → PDF"]), **style).pack(pady=5)
    tk.Button(root, text="🖼 Ảnh → PDF",
              command=lambda: launch_and_close(CANDIDATES["Ảnh → PDF"]), **style).pack(pady=5)

    tk.Button(root, text="Đóng", command=root.destroy, width=10).pack(pady=8)
    return root

if __name__ == "__main__":
    ui = build_ui()
    ui.mainloop()
