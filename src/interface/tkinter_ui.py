from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Callable, Optional, Tuple

class ConverterUI:
    """
    Reusable Tkinter UI. Business logic nằm trong callbacks do main truyền vào.
    """

    def __init__(
        self,
        root: tk.Tk,
        on_select: Callable[[], None],
        on_convert: Callable[[], None],
        on_open_downloads: Callable[[], None],
        *,
        title_text: str = "🔄 Word to PDF Converter",
        select_button_text: str = "Chọn file Word",
        convert_button_text: str = "Chuyển sang PDF",
        downloads_hint_text: str = "💾 File PDF sẽ được tải xuống trong thư mục Downloads",
        supported_extensions: Tuple[str, ...] = (".doc", ".docx"),
        window_title: Optional[str] = None,
        window_size: str = "680x420",
    ) -> None:
        self.root = root
        self.on_select = on_select
        self.on_convert = on_convert
        self.on_open_downloads = on_open_downloads

        self._title_text = title_text
        self._select_button_text = select_button_text
        self._convert_button_text = convert_button_text
        self._downloads_hint_text = downloads_hint_text
        self._supported_extensions = tuple(supported_extensions)

        self.status_label: Optional[ttk.Label] = None
        self.progress_bar: Optional[ttk.Progressbar] = None
        self.select_btn: Optional[ttk.Button] = None
        self.convert_btn: Optional[ttk.Button] = None
        self.download_btn: Optional[ttk.Button] = None
        self.quit_btn: Optional[ttk.Button] = None

        self._setup_window(window_title or self._title_text.strip("🔄 ").strip())
        self.root.geometry(window_size)
        self.root.minsize(560, 360)

        self._setup_styles()
        self._build_layout()

    # Public API
    def set_progress(self, value: float) -> None:
        if self.progress_bar is not None:
            self.progress_bar["value"] = max(0, min(100, value))
            self.root.update_idletasks()

    def update_status(self, text: str, progress: Optional[float] = None) -> None:
        if self.status_label is not None:
            self.status_label.config(text=text)
        if progress is not None:
            self.set_progress(progress)

    def set_buttons_enabled(
        self, *, select: Optional[bool] = None, convert: Optional[bool] = None,
        open_downloads: Optional[bool] = None, quit_btn: Optional[bool] = None
    ) -> None:
        if select is not None and self.select_btn is not None:
            self._set_btn_state(self.select_btn, select)
        if convert is not None and self.convert_btn is not None:
            self._set_btn_state(self.convert_btn, convert)
        if open_downloads is not None and self.download_btn is not None:
            self._set_btn_state(self.download_btn, open_downloads)
        if quit_btn is not None and self.quit_btn is not None:
            self._set_btn_state(self.quit_btn, quit_btn)

    def alert_info(self, title: str, message: str) -> None:
        messagebox.showinfo(title, message, parent=self.root)

    def alert_warning(self, title: str, message: str) -> None:
        messagebox.showwarning(title, message, parent=self.root)

    def alert_error(self, title: str, message: str) -> None:
        messagebox.showerror(title, message, parent=self.root)

    # Internal
    def _setup_window(self, window_title: str) -> None:
        self.root.title(window_title)
        try:
            self.root.iconbitmap(default="")
        except Exception:
            pass

    def _setup_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            if "azure" in style.theme_names():
                style.theme_use("azure")
            elif "clam" in style.theme_names():
                style.theme_use("clam")
        except Exception:
            pass

        style.configure("Title.TLabel", font=("Arial", 16, "bold"))
        style.configure("Info.TLabel", foreground="#1b5e20")
        style.configure("Warn.TLabel", foreground="#e65100")
        style.configure("Error.TLabel", foreground="#b71c1c")
        style.configure("Hint.TLabel", foreground="#666666")

        style.configure("TButton", padding=(10, 6))
        style.configure("Accent.TButton", font=("Arial", 10, "bold"))
        style.configure("Custom.Horizontal.TProgressbar", troughcolor="#e0e0e0")

    def _build_layout(self) -> None:
        container = ttk.Frame(self.root, padding=16)
        container.pack(fill=tk.BOTH, expand=True)

        title = ttk.Label(container, text=self._title_text, style="Title.TLabel")
        title.pack(anchor=tk.CENTER, pady=(0, 18))

        step1 = ttk.LabelFrame(container, text="📁 Bước 1: Chọn tệp", padding=12)
        step1.pack(fill=tk.X, pady=(0, 10))

        ext_text = ", ".join(self._supported_extensions) if self._supported_extensions else "*.*"
        ttk.Label(step1, text=f"Hỗ trợ: {ext_text}").pack(anchor=tk.W)

        self.select_btn = ttk.Button(step1, text=self._select_button_text, command=self.on_select)
        self.select_btn.pack(anchor=tk.W, pady=(8, 0))

        step2 = ttk.LabelFrame(container, text="⚙️ Bước 2: Chuyển đổi", padding=12)
        step2.pack(fill=tk.X, pady=(0, 10))

        self.convert_btn = ttk.Button(step2, text=self._convert_button_text, command=self.on_convert, style="Accent.TButton")
        self.convert_btn.pack(anchor=tk.W)

        step3 = ttk.LabelFrame(container, text="⬇️ Bước 3: Tải xuống", padding=12)
        step3.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(step3, text=self._downloads_hint_text, style="Hint.TLabel").pack(anchor=tk.W)
        self.download_btn = ttk.Button(step3, text="Mở thư mục Downloads", command=self.on_open_downloads)
        self.download_btn.pack(anchor=tk.W, pady=(8, 0))

        prog = ttk.Frame(container, padding=(0, 4))
        prog.pack(fill=tk.X, pady=(6, 0))

        self.progress_bar = ttk.Progressbar(prog, mode="determinate", style="Custom.Horizontal.TProgressbar")
        self.progress_bar.pack(fill=tk.X)

        self.status_label = ttk.Label(prog, text="✅ Sẵn sàng - Hãy chọn tệp để bắt đầu", style="Info.TLabel")
        self.status_label.pack(anchor=tk.W, pady=(6, 0))

        footer = ttk.Frame(container)
        footer.pack(fill=tk.X, pady=(12, 0))

        self.quit_btn = ttk.Button(footer, text="Thoát", command=self.root.quit)
        self.quit_btn.pack(side=tk.RIGHT)

    @staticmethod
    def _set_btn_state(btn: ttk.Button, enabled: bool) -> None:
        try:
            btn.config(state=(tk.NORMAL if enabled else tk.DISABLED))
        except Exception:
            pass
