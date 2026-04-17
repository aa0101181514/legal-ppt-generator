"""法庭風格 PPT 樣式常數。

所有 generate_pptx.py 用到的顏色、字型、版面參數集中在此。
以 v2_20260407 逃漏稅答辯簡報為基準。
"""

from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor


# ─── 版面 ─────────────────────────────────────────────────

SLIDE_WIDTH = Inches(13.333)   # 16:9 標準
SLIDE_HEIGHT = Inches(7.5)

MARGIN_LEFT = Inches(0.6)
MARGIN_RIGHT = Inches(0.6)
MARGIN_TOP = Inches(0.5)
MARGIN_BOTTOM = Inches(0.5)

CONTENT_WIDTH = SLIDE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT
CONTENT_HEIGHT = SLIDE_HEIGHT - MARGIN_TOP - MARGIN_BOTTOM


# ─── 字型 ─────────────────────────────────────────────────

# 法庭常用字型，按優先順序
FONT_ZH = "標楷體"        # 書狀主字型；macOS 通常內建 BiauKai
FONT_ZH_FALLBACK = "PingFang TC"
FONT_EN = "Times New Roman"

FONT_SIZE_TITLE_COVER = Pt(40)     # 封面主標
FONT_SIZE_SUBTITLE_COVER = Pt(20)  # 封面副資訊

FONT_SIZE_SECTION_HEADER = Pt(36)  # 章節分隔頁（壹、貳...）

FONT_SIZE_TITLE = Pt(28)           # 一般內容頁標題
FONT_SIZE_SUBTITLE = Pt(18)        # 副標
FONT_SIZE_BODY = Pt(18)            # 正文 bullet
FONT_SIZE_BODY_SUB = Pt(14)        # 第二層 bullet
FONT_SIZE_QUOTE = Pt(14)           # 判決引文 blockquote
FONT_SIZE_FOOTER = Pt(10)          # 卷頁引用 footer


# ─── 顏色 ─────────────────────────────────────────────────
# v2_20260407 配色：深藍主色 + 深灰正文 + 橙黃強調

COLOR_PRIMARY = RGBColor(0x1F, 0x3A, 0x68)       # 深藍（標題、章節）
COLOR_PRIMARY_LIGHT = RGBColor(0x4A, 0x6A, 0x9A) # 次標題淺藍
COLOR_BODY = RGBColor(0x33, 0x33, 0x33)          # 深灰正文
COLOR_ACCENT = RGBColor(0xC0, 0x80, 0x00)        # 橙黃強調（核心主張）
COLOR_WARNING = RGBColor(0xC0, 0x30, 0x30)       # 紅色（驗證失敗、未確認）
COLOR_QUOTE_BG = RGBColor(0xF5, 0xF2, 0xE8)      # 引文淺米色底
COLOR_QUOTE_TEXT = RGBColor(0x50, 0x3A, 0x1E)    # 引文棕色字
COLOR_BG = RGBColor(0xFF, 0xFF, 0xFF)            # 純白底
COLOR_SECTION_BG = RGBColor(0x1F, 0x3A, 0x68)    # 章節頁深藍底
COLOR_LINE = RGBColor(0xC0, 0xC0, 0xC0)          # 分隔線淺灰


# ─── 段落樣式 ─────────────────────────────────────────────

LINE_SPACING_BODY = 1.3
LINE_SPACING_QUOTE = 1.25
LINE_SPACING_TITLE = 1.1


# ─── Slide Layout 類型 ───────────────────────────────────
# 供 generate_pptx.py 使用的語意標籤

LAYOUT_COVER = "cover"              # 封面
LAYOUT_TOC = "toc"                  # 目錄
LAYOUT_SECTION = "section"          # 章節分隔頁（壹、貳...）
LAYOUT_CONTENT = "content"          # 標準內容頁
LAYOUT_CITATION = "citation"        # 判決引用專頁
LAYOUT_CONCLUSION = "conclusion"    # 結論 / 聲明
LAYOUT_APPENDIX = "appendix"        # 附錄（引用判決清單）


# ─── 引用標籤 ─────────────────────────────────────────────
# 用於顯示驗證狀態

VERIFIED_MARK = "✓ 已驗證"
UNVERIFIED_MARK = "⚠ 未驗證"
FAILED_MARK = "⚠ API 驗證失敗"


def apply_run_style(run, *, size=None, color=None, bold=False, font_zh=None):
    """快速套用字型樣式到 python-pptx Run。"""
    font = run.font
    if size is not None:
        font.size = size
    if color is not None:
        font.color.rgb = color
    font.bold = bold
    font.name = font_zh or FONT_ZH
