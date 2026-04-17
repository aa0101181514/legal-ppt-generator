"""法庭風格 PPT 樣式常數。

規格來源：ABC牙醫聯盟 115金重訴1 逃漏稅答辯簡報 v2_20260407.pptx
色碼、字級、版面尺寸均為實測值。
"""

from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor


# ─── 版面 16:9 ─────────────────────────────────────────────

SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)


# ─── 左側深藍豎條 ──────────────────────────────────────────

LEFT_STRIP_WIDTH = Inches(0.08)
LEFT_STRIP_COLOR = RGBColor(0x1A, 0x3C, 0x6E)


# ─── 標題區 ────────────────────────────────────────────────

TITLE_LEFT = Inches(0.60)
TITLE_TOP = Inches(0.25)
TITLE_WIDTH = Inches(12.00)
TITLE_HEIGHT = Inches(0.80)

TITLE_SEPARATOR_TOP = Inches(1.00)   # 標題下方細分隔線
TITLE_SEPARATOR_HEIGHT = Emu(18288)  # 0.02"


# ─── 內容區 ────────────────────────────────────────────────

CONTENT_LEFT = Inches(0.80)
CONTENT_TOP = Inches(1.25)
CONTENT_WIDTH = Inches(12.00)


# ─── 頁碼 ─────────────────────────────────────────────────

PAGE_NUM_LEFT = Inches(11.80)
PAGE_NUM_TOP = Inches(7.05)
PAGE_NUM_WIDTH = Inches(1.20)
PAGE_NUM_HEIGHT = Inches(0.35)


# ─── 字型 ─────────────────────────────────────────────────
# v2 全部用 Times New Roman，但中文字距擠，改用標楷體
# 英文 / 數字仍用 Times，python-pptx 會依字元選 east_asian font

FONT_ZH = "標楷體"
FONT_EN = "Times New Roman"

# 字級（pt）
FONT_SIZE_COVER_TITLE = Pt(36)       # 封面主標
FONT_SIZE_COVER_DEFENDANT = Pt(28)   # 封面被告列
FONT_SIZE_COVER_SUBTITLE = Pt(26)    # 封面副標
FONT_SIZE_COVER_DATE = Pt(18)        # 封面日期/律師

FONT_SIZE_TITLE = Pt(30)             # 論述頁標題
FONT_SIZE_EXHIBIT_TITLE = Pt(26)     # 卷證原件頁標題

FONT_SIZE_LABEL = Pt(18)             # 「核心主張」「關鍵事實」區塊標籤
FONT_SIZE_SUBLABEL = Pt(16)          # 卷證副標（「113偵27269卷5 p.253 — …」）
FONT_SIZE_BODY = Pt(17)              # 正文 bullet
FONT_SIZE_BODY_SMALL = Pt(14)        # 第二層、附註

FONT_SIZE_PAGE_NUM = Pt(9)


# ─── 顏色 ─────────────────────────────────────────────────

COLOR_PRIMARY = RGBColor(0x1A, 0x3C, 0x6E)       # 深藍 — 標題、豎條
COLOR_LABEL = RGBColor(0x8B, 0x00, 0x00)         # 深紅 — 區塊標籤
COLOR_BODY = RGBColor(0x33, 0x33, 0x33)          # 深灰 — 正文
COLOR_MUTED = RGBColor(0x66, 0x66, 0x66)         # 中灰 — 日期、頁碼
COLOR_BG = RGBColor(0xFF, 0xFF, 0xFF)            # 白底
COLOR_LINE = RGBColor(0xC0, 0xC0, 0xC0)          # 分隔線

# 狀態色
COLOR_WARNING = RGBColor(0xC0, 0x30, 0x30)       # 警告/未驗證


# ─── 標籤字串 ─────────────────────────────────────────────

LABEL_MAIN_POINT = "核心主張"
LABEL_KEY_FACTS = "關鍵事實"
LABEL_REASONING = "論理基礎"
LABEL_CITATION = "相關判決"
LABEL_EVIDENCE = "卷證依據"

VERIFIED_MARK = "✓"
UNVERIFIED_MARK = "⚠"
FAILED_MARK = "✗"


# ─── 區塊間距 ─────────────────────────────────────────────

BLOCK_GAP = Inches(0.25)
LABEL_TO_BODY_GAP = Inches(0.05)


# ─── Layout 類型 ──────────────────────────────────────────

LAYOUT_COVER = "cover"
LAYOUT_AGENDA = "agenda"          # 答辯架構總覽
LAYOUT_CONTENT = "content"        # 論述頁（核心主張 + 關鍵事實）
LAYOUT_EXHIBIT = "exhibit"        # 卷證原件頁（截圖）
LAYOUT_TIMELINE = "timeline"      # 時間軸表格頁
LAYOUT_CONCLUSION = "conclusion"
LAYOUT_APPENDIX = "appendix"


def apply_run_style(run, *, size=None, color=None, bold=False,
                    font_zh=None, font_en=None):
    """套用字型樣式到 python-pptx Run。

    中文字元由東亞字型處理，英文 / 數字用拉丁字型。
    """
    font = run.font
    if size is not None:
        font.size = size
    if color is not None:
        font.color.rgb = color
    font.bold = bold
    # Latin 字型
    font.name = font_en or FONT_EN
    # East Asian 字型（需要設定 XML）
    from pptx.oxml.ns import qn
    rPr = run._r.get_or_add_rPr()
    # 移除舊的 eastAsia
    for ea in rPr.findall(qn("a:ea")):
        rPr.remove(ea)
    from lxml import etree
    ea = etree.SubElement(rPr, qn("a:ea"))
    ea.set("typeface", font_zh or FONT_ZH)
