from __future__ import annotations

import base64
import io
from pathlib import Path
from typing import Iterable

from PIL import Image, ImageEnhance, ImageFilter, ImageOps

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}
PDF_EXTENSIONS = {".pdf"}


def supported_file_filter() -> str:
    return "Documents (*.pdf *.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp)"


def is_supported(path: str | Path) -> bool:
    return Path(path).suffix.lower() in IMAGE_EXTENSIONS | PDF_EXTENSIONS


def render_page(path: str | Path, page_number: int = 0, zoom: float = 2.0) -> Image.Image:
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix in IMAGE_EXTENSIONS:
        return Image.open(path).convert("RGB")
    if suffix in PDF_EXTENSIONS:
        try:
            import fitz  # type: ignore
        except ImportError as exc:
            raise RuntimeError("PyMuPDF is required to read PDF files.") from exc
        doc = fitz.open(path)
        if page_number < 0 or page_number >= len(doc):
            page_number = 0
        page = doc.load_page(page_number)
        matrix = fitz.Matrix(zoom, zoom)
        pixmap = page.get_pixmap(matrix=matrix, alpha=False)
        image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
        doc.close()
        return image
    raise ValueError(f"Unsupported file type: {path.suffix}")


def page_count(path: str | Path) -> int:
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix in IMAGE_EXTENSIONS:
        return 1
    if suffix in PDF_EXTENSIONS:
        try:
            import fitz  # type: ignore
        except ImportError as exc:
            raise RuntimeError("PyMuPDF is required to read PDF files.") from exc
        doc = fitz.open(path)
        count = len(doc)
        doc.close()
        return max(1, count)
    raise ValueError(f"Unsupported file type: {path.suffix}")


def crop_normalized(image: Image.Image, rect: tuple[float, float, float, float]) -> Image.Image:
    x, y, width, height = rect
    left = max(0, int(x * image.width))
    top = max(0, int(y * image.height))
    right = min(image.width, int((x + width) * image.width))
    bottom = min(image.height, int((y + height) * image.height))
    if right <= left or bottom <= top:
        return Image.new("RGB", (1, 1), "white")
    return image.crop((left, top, right, bottom))


def image_to_base64_png(image: Image.Image) -> str:
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return base64.b64encode(buffer.getvalue()).decode("ascii")


def make_thumbnail(image: Image.Image, max_side: int = 1800) -> Image.Image:
    copy = image.copy()
    copy.thumbnail((max_side, max_side))
    return copy


def prepare_for_vision(image: Image.Image, min_width: int = 1800, max_side: int = 3000) -> Image.Image:
    prepared = image.convert("RGB")
    width, height = prepared.size
    scale = max(1.0, min_width / max(1, width))
    scale = min(scale, max_side / max(width, height))
    if scale > 1.01:
        prepared = prepared.resize((int(width * scale), int(height * scale)), Image.Resampling.LANCZOS)
    prepared = ImageOps.autocontrast(prepared.convert("L")).convert("RGB")
    prepared = ImageEnhance.Contrast(prepared).enhance(1.25)
    return prepared.filter(ImageFilter.SHARPEN)


def existing_supported(paths: Iterable[str]) -> list[str]:
    return [str(Path(path)) for path in paths if Path(path).exists() and is_supported(path)]
