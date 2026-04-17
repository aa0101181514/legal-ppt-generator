"""讀 JSON 大綱，產出法庭風格 .pptx（貼近 v2_20260407 規格）。

JSON 大綱結構參見 examples/sample_outline.json

使用：
    python generate_pptx.py <outline.json> <output.pptx>
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt, Emu

import ppt_theme as T

try:
    import pdf_render
except ImportError:
    pdf_render = None


# ─── 共用 helpers ─────────────────────────────────────────

def _add_run(paragraph, text, *, size=None, color=None, bold=False):
    """在既有 paragraph 上加 run，套用字型樣式。"""
    run = paragraph.add_run()
    run.text = text
    T.apply_run_style(run, size=size, color=color, bold=bold)
    return run


def _textbox(slide, left, top, width, height, *, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, wrap=True, no_margin=True):
    """建 textbox 回傳 (box, tf)。"""
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    if no_margin:
        tf.margin_left = Inches(0.05)
        tf.margin_right = Inches(0.05)
        tf.margin_top = Inches(0.02)
        tf.margin_bottom = Inches(0.02)
    # 預設第一 paragraph 對齊
    tf.paragraphs[0].alignment = align
    return box, tf


def _add_rect(slide, left, top, width, height, fill_color):
    """加一個實心矩形。"""
    from pptx.shapes.autoshape import Shape
    from pptx.enum.shapes import MSO_SHAPE
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    return shape


def _add_left_strip(slide):
    """左側深藍豎條（所有頁共用元素）。"""
    return _add_rect(
        slide, 0, 0,
        T.LEFT_STRIP_WIDTH, T.SLIDE_HEIGHT,
        T.LEFT_STRIP_COLOR,
    )


def _add_title_with_sep(slide, title_text, *, size=None):
    """頁面標題 + 下方細分隔線（論述/卷證/其他內容頁共用）。"""
    size = size or T.FONT_SIZE_TITLE
    _, tf = _textbox(
        slide, T.TITLE_LEFT, T.TITLE_TOP,
        T.TITLE_WIDTH, T.TITLE_HEIGHT,
    )
    p = tf.paragraphs[0]
    _add_run(p, title_text, size=size, color=T.COLOR_PRIMARY, bold=True)
    # 分隔線
    _add_rect(
        slide,
        T.TITLE_LEFT, T.TITLE_SEPARATOR_TOP,
        T.TITLE_WIDTH, T.TITLE_SEPARATOR_HEIGHT,
        T.COLOR_PRIMARY,
    )


def _add_page_number(slide, num):
    """右下角頁碼。"""
    _, tf = _textbox(
        slide, T.PAGE_NUM_LEFT, T.PAGE_NUM_TOP,
        T.PAGE_NUM_WIDTH, T.PAGE_NUM_HEIGHT,
        align=PP_ALIGN.RIGHT,
    )
    _add_run(
        tf.paragraphs[0], str(num),
        size=T.FONT_SIZE_PAGE_NUM, color=T.COLOR_MUTED,
    )


def _fill_bg(slide, color):
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = color


def _blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


# ─── 各 layout 渲染器 ─────────────────────────────────────

def render_cover(prs, meta):
    slide = _blank(prs)
    _fill_bg(slide, T.COLOR_BG)
    _add_left_strip(slide)

    # 主標（2 行）：法院案號
    _, tf = _textbox(
        slide, Inches(1.20), Inches(1.50),
        Inches(11.0), Inches(1.50),
    )
    p = tf.paragraphs[0]
    _add_run(p, meta.get("court", "臺灣OO地方法院"),
             size=T.FONT_SIZE_COVER_TITLE, color=T.COLOR_PRIMARY, bold=True)
    _add_run(p, f" {meta.get('case_no', '')}",
             size=T.FONT_SIZE_COVER_TITLE, color=T.COLOR_PRIMARY, bold=True)

    # 被告
    defendants = meta.get("defendants", [])
    if defendants:
        defs = defendants if isinstance(defendants, list) else [defendants]
        _, tf = _textbox(
            slide, Inches(1.20), Inches(3.00),
            Inches(10.0), Inches(0.6),
        )
        _add_run(
            tf.paragraphs[0],
            f"被告 {('、').join(defs)}",
            size=T.FONT_SIZE_COVER_DEFENDANT, color=T.COLOR_BODY,
        )

    # 副標
    if meta.get("subtitle"):
        _, tf = _textbox(
            slide, Inches(1.20), Inches(3.90),
            Inches(10.0), Inches(0.6),
        )
        _add_run(
            tf.paragraphs[0], meta["subtitle"],
            size=T.FONT_SIZE_COVER_SUBTITLE, color=T.COLOR_BODY,
        )

    # 細分隔線
    _add_rect(
        slide,
        Inches(1.20), Inches(5.00),
        Inches(5.00), Emu(18288),
        T.COLOR_PRIMARY,
    )

    # 日期 + 律師
    _, tf = _textbox(
        slide, Inches(1.20), Inches(5.40),
        Inches(10.0), Inches(1.4),
    )
    info_lines = []
    if meta.get("date"):
        info_lines.append(meta["date"])
    for name in meta.get("lawyers", []) or []:
        info_lines.append(name)
    for i, line in enumerate(info_lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        _add_run(p, line, size=T.FONT_SIZE_COVER_DATE, color=T.COLOR_MUTED)


def render_agenda(prs, meta, agenda_items):
    """答辯架構總覽。"""
    slide = _blank(prs)
    _fill_bg(slide, T.COLOR_BG)
    _add_left_strip(slide)
    _add_title_with_sep(slide, "答辯架構總覽")

    # 每 item 一行
    _, tf = _textbox(
        slide, T.CONTENT_LEFT, Inches(1.40),
        T.CONTENT_WIDTH, Inches(5.5),
    )
    for i, item in enumerate(agenda_items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        _add_run(p, f"{item}", size=T.FONT_SIZE_BODY, color=T.COLOR_BODY)


def render_content(prs, slide_data):
    """標準論述頁：標題 + 核心主張 + 關鍵事實。"""
    slide = _blank(prs)
    _fill_bg(slide, T.COLOR_BG)
    _add_left_strip(slide)
    _add_title_with_sep(slide, slide_data.get("title", ""))

    current_top = Inches(1.25)

    # 依序渲染 sections
    for section in slide_data.get("sections", []):
        label = section.get("label")
        bullets = section.get("bullets", [])
        if not bullets:
            continue

        # 區塊標籤（深紅 bold）
        _, tf_lb = _textbox(
            slide, T.CONTENT_LEFT, current_top,
            T.CONTENT_WIDTH, Inches(0.45),
        )
        _add_run(
            tf_lb.paragraphs[0], label,
            size=T.FONT_SIZE_LABEL, color=T.COLOR_LABEL, bold=True,
        )
        current_top += Inches(0.45)

        # bullets
        body_height = Inches(0.45 * len(bullets) + 0.3)
        _, tf_body = _textbox(
            slide, T.CONTENT_LEFT + Inches(0.20), current_top,
            T.CONTENT_WIDTH - Inches(0.20), body_height,
        )
        for i, b in enumerate(bullets):
            p = tf_body.paragraphs[0] if i == 0 else tf_body.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            _add_run(p, "• ", size=T.FONT_SIZE_BODY, color=T.COLOR_BODY)
            # 若 bullet 是 dict，支援 emphasis
            if isinstance(b, dict):
                text = b.get("text", "")
                if b.get("bold"):
                    _add_run(p, text, size=T.FONT_SIZE_BODY,
                             color=T.COLOR_BODY, bold=True)
                else:
                    _add_run(p, text, size=T.FONT_SIZE_BODY, color=T.COLOR_BODY)
            else:
                _add_run(p, str(b), size=T.FONT_SIZE_BODY, color=T.COLOR_BODY)

        current_top += body_height + T.BLOCK_GAP


def render_exhibit(prs, slide_data, base_dir=None):
    """卷證原件頁：左側卷頁副標 + 右側整頁 PNG。

    slide_data 欄位：
      - title: "卷證原件 ① 113.8.16 — 國稅局復函"
      - subtitle: "113偵27269卷5 p.253 — 財政部高雄國稅局113.8.19復函"
      - pdf_path: "~/Desktop/ABC牙醫聯盟/3_刑事/閱卷資料/113偵27269卷5.pdf"
      - page_num: 253
      - note: （選填）下方說明
    """
    slide = _blank(prs)
    _fill_bg(slide, T.COLOR_BG)
    _add_left_strip(slide)

    _add_title_with_sep(
        slide, slide_data.get("title", ""),
        size=T.FONT_SIZE_EXHIBIT_TITLE,
    )

    # 副標（卷頁標記）
    if slide_data.get("subtitle"):
        _, tf = _textbox(
            slide, T.CONTENT_LEFT, Inches(1.20),
            T.CONTENT_WIDTH, Inches(0.45),
        )
        _add_run(
            tf.paragraphs[0], slide_data["subtitle"],
            size=T.FONT_SIZE_SUBLABEL, color=T.COLOR_PRIMARY, bold=True,
        )

    # 渲染 PDF 頁成 PNG 並插入右側
    pdf_path = slide_data.get("pdf_path")
    page_num = slide_data.get("page_num")
    png_path = slide_data.get("image_path")  # 可直接指定已有的 PNG

    if not png_path and pdf_path and page_num and pdf_render:
        try:
            resolved_pdf = Path(pdf_path).expanduser()
            if not resolved_pdf.is_absolute() and base_dir:
                resolved_pdf = Path(base_dir) / pdf_path
            png_path = pdf_render.render_page(resolved_pdf, page_num)
        except Exception as e:
            # 失敗仍保留 slide，標示警告
            png_path = None
            _, tf_err = _textbox(
                slide, T.CONTENT_LEFT, Inches(2.0),
                T.CONTENT_WIDTH, Inches(0.6),
            )
            _add_run(
                tf_err.paragraphs[0],
                f"⚠ PDF 渲染失敗：{e}",
                size=T.FONT_SIZE_BODY, color=T.COLOR_WARNING,
            )

    if png_path and Path(png_path).exists():
        # 右側圖片：起始 x=6.35"，最大寬 6.45"，最大高 5.8"
        img = slide.shapes.add_picture(
            str(png_path),
            left=Inches(6.35), top=Inches(1.30),
            width=Inches(6.45),
        )
        # 若高度超出 5.8"，縮小
        max_h = Inches(5.8)
        if img.height > max_h:
            ratio = max_h / img.height
            img.height = int(img.height * ratio)
            img.width = int(img.width * ratio)

    # 額外說明
    if slide_data.get("note"):
        _, tf_note = _textbox(
            slide, T.CONTENT_LEFT, Inches(1.80),
            Inches(5.5), Inches(5.0),
        )
        _add_run(
            tf_note.paragraphs[0], slide_data["note"],
            size=T.FONT_SIZE_BODY, color=T.COLOR_BODY,
        )


def render_conclusion(prs, slide_data):
    slide = _blank(prs)
    _fill_bg(slide, T.COLOR_BG)
    _add_left_strip(slide)
    _add_title_with_sep(slide, slide_data.get("title", "結論"))

    _, tf = _textbox(
        slide, T.CONTENT_LEFT, Inches(1.35),
        T.CONTENT_WIDTH, Inches(5.5),
    )
    statements = slide_data.get("statements", [])
    for i, s in enumerate(statements):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        _add_run(p, f"{i+1}. ", size=T.FONT_SIZE_BODY,
                 color=T.COLOR_PRIMARY, bold=True)
        _add_run(p, s, size=T.FONT_SIZE_BODY, color=T.COLOR_BODY)


def render_appendix(prs, citations_flat):
    slide = _blank(prs)
    _fill_bg(slide, T.COLOR_BG)
    _add_left_strip(slide)
    _add_title_with_sep(slide, "附錄 — 本簡報引用判決（均經 dr-lawbot 驗證）")

    _, tf = _textbox(
        slide, T.CONTENT_LEFT, Inches(1.35),
        T.CONTENT_WIDTH, Inches(5.5),
    )
    for i, c in enumerate(citations_flat):
        line = _format_citation_line(c)
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        _add_run(p, f"{i+1}. {line}", size=T.FONT_SIZE_BODY_SMALL, color=T.COLOR_BODY)
        if c.get("search_url"):
            p_url = tf.add_paragraph()
            p_url.alignment = PP_ALIGN.LEFT
            _add_run(p_url, f"    {c['search_url']}",
                     size=T.FONT_SIZE_BODY_SMALL, color=T.COLOR_MUTED)


def _format_citation_line(cite):
    jyear = cite.get("jyear")
    jcase = cite.get("jcase")
    jno = cite.get("jno")
    title = cite.get("title") or ""
    doc_id = cite.get("doc_id") or ""
    prefix = doc_id.split(",")[0] if "," in doc_id else ""
    court_map = {
        "TPSM": "最高法院", "TPSV": "最高法院",
        "TPAA": "最高行政法院",
        "TPHV": "台灣高等法院", "TCHV": "台中高分院",
        "TNHV": "台南高分院", "KSHV": "高雄高分院",
        "TPBA": "台北高等行政法院", "TCBA": "台中高等行政法院",
        "KSBA": "高雄高等行政法院",
    }
    court = court_map.get(prefix, "")
    if jyear and jcase and jno:
        line = f"{court} {jyear} 年度{jcase}字第 {jno} 號判決"
    else:
        line = cite.get("query", "（未指定字號）")
    if title:
        line += f"（{title}）"
    return line


# ─── main ───────────────────────────────────────────────

def generate(outline: dict, output_path: str | Path) -> Path:
    prs = Presentation()
    prs.slide_width = T.SLIDE_WIDTH
    prs.slide_height = T.SLIDE_HEIGHT

    slides = outline.get("slides", [])
    meta = outline.get("meta", {})
    base_dir = outline.get("base_dir")

    all_citations = []

    for idx, s in enumerate(slides, start=1):
        layout = s.get("layout", T.LAYOUT_CONTENT)

        if layout == T.LAYOUT_COVER:
            render_cover(prs, meta)
        elif layout == T.LAYOUT_AGENDA:
            render_agenda(prs, meta, s.get("items", []))
        elif layout == T.LAYOUT_CONTENT:
            render_content(prs, s)
            for c in s.get("citations", []):
                all_citations.append(c)
        elif layout == T.LAYOUT_EXHIBIT:
            render_exhibit(prs, s, base_dir=base_dir)
        elif layout == T.LAYOUT_CONCLUSION:
            render_conclusion(prs, s)
        elif layout == T.LAYOUT_APPENDIX:
            render_appendix(prs, all_citations)
        else:
            # 未知 fallback
            render_content(prs, s)

        # 頁碼（封面通常不顯示，但 v2 封面有寫 1）
        slide = prs.slides[idx - 1]
        _add_page_number(slide, idx)

    out = Path(output_path).expanduser()
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out)
    return out.resolve()


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("用法：python generate_pptx.py <outline.json> <output.pptx>")
        sys.exit(1)
    json_path = Path(sys.argv[1])
    out = Path(sys.argv[2])
    outline = json.loads(json_path.read_text(encoding="utf-8"))
    result = generate(outline, out)
    print(f"✓ 已產出：{result}")
