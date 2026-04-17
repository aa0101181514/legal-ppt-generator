"""卷頁偏移偵測工具。

律師閱卷 PDF 常見「PDF 頁碼」與「卷頁戳印」不一致的情況。

三種用法：

1. 給定 (PDF + 已知某頁內容) → 找出該內容所在的 PDF 頁，推算偏移
       python detect_offset.py find 113偵27269卷17.pdf "第3條 醫師自負稅金"

2. 給定 (PDF + 卷頁戳) 並指定 PDF 頁碼 → 計算偏移
       python detect_offset.py calc 113偵27269卷17.pdf --pdf-page 9 --stamp 9

3. 把 PDF 頁碼 ↔ 卷頁戳相互轉換（需已知偏移）
       python detect_offset.py convert --offset -1 --stamp 9
       → 輸出 PDF 頁 = 10

偏移公式：
    pdf_page = stamp_page + offset
    例：某卷 offset = -1 表示「卷頁戳的第 9 頁」對應「PDF 的第 10 頁」
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Optional

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None


def search_text(pdf_path: str | Path, keyword: str, max_pages: int = 0) -> list[int]:
    """在 PDF 中搜尋 keyword，回傳所有匹配的 1-based 頁碼。

    max_pages=0 表示搜全部；>0 限制前 N 頁。
    """
    if fitz is None:
        raise RuntimeError("未安裝 PyMuPDF。pip install pymupdf")

    pdf_path = Path(pdf_path).expanduser()
    if not pdf_path.exists():
        raise FileNotFoundError(pdf_path)

    doc = fitz.open(pdf_path)
    try:
        hits = []
        total = len(doc)
        last_page = total if max_pages == 0 else min(max_pages, total)
        for i in range(last_page):
            page = doc[i]
            if page.search_for(keyword):
                hits.append(i + 1)
        return hits
    finally:
        doc.close()


def calc_offset(pdf_page: int, stamp_page: int) -> int:
    """offset = pdf_page - stamp_page"""
    return pdf_page - stamp_page


def stamp_to_pdf(stamp_page: int, offset: int) -> int:
    return stamp_page + offset


def pdf_to_stamp(pdf_page: int, offset: int) -> int:
    return pdf_page - offset


def cmd_find(args):
    hits = search_text(args.pdf, args.keyword, max_pages=args.max_pages)
    if not hits:
        print(json.dumps({
            "found": False,
            "keyword": args.keyword,
            "message": "PDF 中查無此關鍵字。"
        }, ensure_ascii=False, indent=2))
        sys.exit(2)

    result = {
        "found": True,
        "keyword": args.keyword,
        "pdf_pages": hits,
        "count": len(hits),
    }
    # 若使用者同時提供 --stamp，計算每個匹配頁的 offset 建議
    if args.stamp is not None:
        result["stamp"] = args.stamp
        result["offset_candidates"] = [p - args.stamp for p in hits]
    print(json.dumps(result, ensure_ascii=False, indent=2))


def cmd_calc(args):
    offset = calc_offset(args.pdf_page, args.stamp)
    result = {
        "pdf_page": args.pdf_page,
        "stamp_page": args.stamp,
        "offset": offset,
        "formula": "pdf_page = stamp_page + offset",
        "interpretation": _interpret(offset),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


def cmd_convert(args):
    if args.stamp is not None and args.pdf_page is not None:
        print("請只指定 --stamp 或 --pdf-page 其中一個", file=sys.stderr)
        sys.exit(1)
    if args.stamp is not None:
        pdf = stamp_to_pdf(args.stamp, args.offset)
        print(json.dumps({
            "input_stamp": args.stamp,
            "offset": args.offset,
            "pdf_page": pdf,
        }, ensure_ascii=False, indent=2))
    elif args.pdf_page is not None:
        st = pdf_to_stamp(args.pdf_page, args.offset)
        print(json.dumps({
            "input_pdf_page": args.pdf_page,
            "offset": args.offset,
            "stamp_page": st,
        }, ensure_ascii=False, indent=2))
    else:
        print("請指定 --stamp 或 --pdf-page", file=sys.stderr)
        sys.exit(1)


def _interpret(offset: int) -> str:
    if offset == 0:
        return "PDF 頁與卷頁戳一致，無偏移。"
    elif offset > 0:
        return f"PDF 頁比卷頁戳多 {offset}（卷首可能有封面/目錄）。"
    else:
        return f"PDF 頁比卷頁戳少 {-offset}（PDF 去掉了前 {-offset} 頁封面）。"


def main():
    parser = argparse.ArgumentParser(description="卷頁偏移偵測")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_find = sub.add_parser("find", help="關鍵字找出所在 PDF 頁")
    p_find.add_argument("pdf")
    p_find.add_argument("keyword")
    p_find.add_argument("--stamp", type=int, default=None,
                        help="若指定，同時計算 offset 建議")
    p_find.add_argument("--max-pages", type=int, default=0,
                        help="限制搜尋前 N 頁（0=全部）")
    p_find.set_defaults(func=cmd_find)

    p_calc = sub.add_parser("calc", help="已知 pdf_page 與 stamp 計算 offset")
    p_calc.add_argument("pdf", help="PDF 路徑（僅紀錄用，不驗證）", nargs="?", default="")
    p_calc.add_argument("--pdf-page", type=int, required=True)
    p_calc.add_argument("--stamp", type=int, required=True)
    p_calc.set_defaults(func=cmd_calc)

    p_conv = sub.add_parser("convert", help="以 offset 做頁碼轉換")
    p_conv.add_argument("--offset", type=int, required=True)
    p_conv.add_argument("--stamp", type=int, default=None)
    p_conv.add_argument("--pdf-page", type=int, default=None)
    p_conv.set_defaults(func=cmd_convert)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
