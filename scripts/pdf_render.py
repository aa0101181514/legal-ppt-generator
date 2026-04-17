"""PDF 頁面渲染為 PNG，供 PPT 卷證截圖頁使用。

使用 PyMuPDF (fitz)，無需系統層 poppler。

使用方式：
    from pdf_render import render_page
    png_path = render_page("/path/to/卷17.pdf", page_num=9, out_dir="/tmp/renders/")

    # 或命令列
    python pdf_render.py <pdf_path> <page_num> [out_path]
"""

from __future__ import annotations

import hashlib
import sys
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None


DEFAULT_DPI = 200    # 足夠 PPT 清晰，檔案大小可控
DEFAULT_OUT_DIR = Path("/tmp/legal_ppt_renders")


def render_page(
    pdf_path: str | Path,
    page_num: int,
    out_dir: str | Path | None = None,
    dpi: int = DEFAULT_DPI,
) -> Path:
    """渲染 PDF 單頁成 PNG。

    Args:
        pdf_path: PDF 絕對路徑
        page_num: 1-based 頁碼（跟律師溝通的頁碼習慣一致）
        out_dir: 輸出資料夾，預設 /tmp/legal_ppt_renders
        dpi: 解析度，預設 200

    Returns:
        產出的 PNG 絕對路徑

    輸出檔名：{pdf basename}_p{page}_{hash[:8]}.png
    hash 確保同一 pdf + 同一頁 + 同一 dpi 不會重複渲染。
    """
    if fitz is None:
        raise RuntimeError(
            "未安裝 PyMuPDF。請執行：pip install pymupdf"
        )

    pdf_path = Path(pdf_path).expanduser().resolve()
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF 不存在：{pdf_path}")

    out_dir = Path(out_dir) if out_dir else DEFAULT_OUT_DIR
    out_dir.mkdir(parents=True, exist_ok=True)

    # 穩定檔名（避免同一筆渲染多次）
    key = f"{pdf_path}|{page_num}|{dpi}".encode("utf-8")
    h = hashlib.md5(key).hexdigest()[:8]
    out_path = out_dir / f"{pdf_path.stem}_p{page_num}_{h}.png"

    if out_path.exists():
        return out_path

    doc = fitz.open(pdf_path)
    try:
        total = len(doc)
        if page_num < 1 or page_num > total:
            raise ValueError(f"{pdf_path.name} 共 {total} 頁，要求第 {page_num} 頁")
        page = doc[page_num - 1]
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        pix.save(str(out_path))
    finally:
        doc.close()

    return out_path


def get_page_count(pdf_path: str | Path) -> int:
    if fitz is None:
        raise RuntimeError("未安裝 PyMuPDF")
    doc = fitz.open(Path(pdf_path).expanduser())
    try:
        return len(doc)
    finally:
        doc.close()


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("用法：python pdf_render.py <pdf_path> <page_num> [out_path]")
        sys.exit(1)
    pdf = sys.argv[1]
    page = int(sys.argv[2])
    out_dir = sys.argv[3] if len(sys.argv) > 3 else None
    result = render_page(pdf, page, out_dir)
    print(f"✓ 已渲染：{result}")
