# src/converters/image_to_pdf.py
from __future__ import annotations

from pathlib import Path
from typing import Optional

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}

def is_image_file(p: str | Path) -> bool:
    return Path(p).suffix.lower() in IMAGE_EXTS

def _open_image_fixed(path: Path):
    """Mở ảnh, sửa xoay EXIF, flatten alpha lên nền trắng để in/nhúng PDF không lỗi."""
    from PIL import Image, ImageOps
    im = Image.open(str(path))
    try:
        im = ImageOps.exif_transpose(im)
    except Exception:
        pass

    # Flatten nếu có alpha để in/PDF không có nền đen
    if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
        from PIL import Image as _Image
        bg = _Image.new("RGB", im.size, (255, 255, 255))
        im = im.convert("RGBA")
        bg.paste(im, mask=im.split()[-1])
        im = bg
    else:
        im = im.convert("RGB")
    return im

def image_to_pdf(
    src_path: str | Path,
    dst_path: Optional[str | Path] = None,
    *,
    dpi: int = 300,
) -> str:
    """
    Ảnh -> PDF 'nét' (ưu tiên lossless):
    - Nếu có reportlab: tạo trang PDF đúng theo kích thước ảnh tại dpi chỉ định (không upscale, không mờ).
    - Nếu không: fallback Pillow với quality cao.
    Trả về đường dẫn PDF.
    """
    src = Path(src_path)
    if not src.exists() or not is_image_file(src):
        raise ValueError(f"Tệp ảnh không hợp lệ hoặc không hỗ trợ: {src}")

    dst = Path(dst_path) if dst_path else src.with_suffix(".pdf")
    dst.parent.mkdir(parents=True, exist_ok=True)

    # Thử dùng ReportLab cho chất lượng hiển thị/print tốt nhất
    try:
        # Import tại runtime để không bắt buộc nếu người dùng chưa cài
        from reportlab.pdfgen import canvas
        from reportlab.lib.utils import ImageReader

        im = _open_image_fixed(src)
        w_px, h_px = im.size

        # Quy đổi pixel -> point theo dpi mong muốn (72 pt = 1 inch)
        # Đảm bảo KHÔNG upscale: trang PDF đúng kích thước ảnh ở dpi đã chọn
        page_w = w_px * 72.0 / dpi
        page_h = h_px * 72.0 / dpi

        c = canvas.Canvas(str(dst), pagesize=(page_w, page_h))
        # Vẽ ảnh phủ toàn bộ trang, không thay đổi tỉ lệ (không upscale vì page size = ảnh/dpi)
        c.drawImage(ImageReader(im), 0, 0, width=page_w, height=page_h, preserveAspectRatio=True, anchor='sw', mask='auto')
        c.showPage()
        c.save()
        return str(dst)

    except Exception:
        # Fallback: dùng Pillow -> PDF (giữ chất lượng cao nhất có thể)
        from PIL import Image
        im = _open_image_fixed(src)
        # 'resolution' ảnh hưởng kích thước hiển thị trên trang, giữ chi tiết gốc
        # 'quality' nếu PDF backend sử dụng JPEG (thường sẽ được dùng)
        im.save(str(dst), "PDF", resolution=dpi, quality=95, optimize=True)
        return str(dst)
