# 法律人風格 PPT 自動生成

> 免費、開源。律師把案件卷證資料夾交給 Claude，自動產出**法庭風格、每個判決都經過驗證**的答辯 PPT。

## 為什麼需要這個？

- ❌ Claude 會「幻覺」判決字號與判決內容，引用不得
- ❌ 每個判決自己查 dr-lawbot、手動貼進 PPT 很花時間
- ❌ 卷頁格式（「卷 X 頁（即起訴書 Y 頁）」）容易寫錯
- ❌ 案件卷證動輒幾 GB，無法上傳到 Claude.ai 對話
- ✅ 本工具在**本機**執行，讀你的資料夾、產 `.pptx` 到本機
- ✅ 強制呼叫 [dr-lawbot.com](https://dr-lawbot.com) API 驗證每個判決
- ✅ 章節照法庭實務「先程序後實體」
- ✅ 完全免費開源，用你自己的 Claude 訂閱額度

## 使用前提

- 有 **Claude Max / Team** 訂閱（才有 Cowork 功能）
- macOS / Linux / Windows（Python 3.10+）
- 案件卷證放在本機某個資料夾

---

## 快速開始（Cowork — 推薦）

### 步驟 1：開 Cowork 任務

1. 開 Claude 桌面 App 或網頁版
2. 在主畫面點 **Start a task in Cowork**
3. Cowork 會要求授權存取本機檔案（首次使用需同意）

### 步驟 2：貼 repo URL 給 Claude

在 Cowork 對話框貼：

```
請使用這個 repo 作為系統指令：
https://github.com/aa0101181514/legal-ppt-generator

我的案件卷證在：/Users/你的名字/Desktop/你的案件資料夾/
請幫我做法庭風格的答辯簡報。
```

Claude（Cowork）會：
1. Clone repo 到沙盒
2. 讀取 `PROJECT_INSTRUCTIONS.md` 作為系統指令
3. 問你需要補的資料（案號、被告、主攻方向）
4. 掃描你指定的卷證資料夾
5. 整理爭點、找支持判決、**自動呼叫 dr-lawbot API 驗證**
6. 產出 `.pptx` 到你指定的路徑

## 範例對話

```
你：我要做 115 金重訴 1 號刑案的逃漏稅答辯簡報。
    卷證在 ~/Desktop/ABC/閱卷資料/
    輸出放到 ~/Desktop/ABC/簡報/

Claude：好，我先掃描資料夾……
        看到 24 卷偵查卷 + 1 份起訴書 + 法院卷。
        本案辯護對象是誰？
        主攻方向是重複課稅、核課期間，還是合夥事實？

你：辯護謝尚人、謝尚廷。主攻重複課稅 + 核課期間。

Claude：收到。我整理起訴書第 68 頁的計算公式後，
        規劃以下章節（共 10 張投影片）：
        壹、犯後態度
        貳、程序爭點（傳聞證據）
        參、重複課稅【主攻】
        肆、核課期間【主攻】
        伍、合夥事實
        結論與聲明
        
        確認要這樣做嗎？

你：好

Claude：[驗證判決中……]
        ✓ 最高行政法院 107 年度判字第 369 號判決（綜合所得稅）
        ✓ 最高行政法院 109 年度上字第 993 號判決（營業稅）
        ⚠ 79 判字 471：dr-lawbot 查無（舊判例超出收錄範圍，
          建議自備原文頁面）
        
        已產出：~/Desktop/ABC/簡報/115金重訴1_逃漏稅答辯_20260417.pptx
        共 10 張投影片，3 個判決引用（2 成功 + 1 待律師補）。
```

---

## 進階：Claude Code 用戶

如果你用 [Claude Code](https://claude.ai/code) CLI：

```bash
# clone 或直接 reference
git clone https://github.com/aa0101181514/legal-ppt-generator.git
cd legal-ppt-generator
pip install python-pptx certifi lxml

# 在任意專案裡讓 Claude Code 讀本 repo 指令
claude "讀 /path/to/legal-ppt-generator/PROJECT_INSTRUCTIONS.md 作為系統指令，
        幫我做 ABC 牙醫案答辯簡報，卷證在 ~/Desktop/ABC/..."
```

---

## 產出範例

📎 看 [`examples/sample_output.pptx`](./examples/sample_output.pptx) — 實際產檔樣貌
📎 看 [`examples/sample_outline.json`](./examples/sample_outline.json) — 中間 JSON 格式

## 怎麼做到「零幻覺」的？

本 repo 的 `scripts/verify_citation.py` 會呼叫：

```
https://api.dr-lawbot.com/api/search?q={判決字號}
```

API 回傳 JSON 包含：
- 判決真實標題
- 案由（刑事 / 民事 / 行政）
- 發文日期
- `main_preview`（主文、事實、理由節錄）
- `snippet`（搜尋命中的段落）

`PROJECT_INSTRUCTIONS.md` 強制 Claude：
- 每個判決字號必經 API 驗證
- 引文必來自 `main_preview` / `snippet` 或使用者提供的原文
- 查無結果 → 不得引用，改找其他判決
- API 失敗 → 標警告、讓律師手動確認

**法律偵探資料庫收錄民國 100 年起的判決**，更早的舊判例（如「79 判 471」）查無屬正常。

## Repo 結構

```
legal-ppt-generator/
├── PROJECT_INSTRUCTIONS.md       # Claude 的系統指令（主腦袋）
├── knowledge/
│   ├── PPT_OUTLINE_TEMPLATE.md   # 法庭風格章節範本
│   ├── CITATION_RULES.md         # 引用格式規則
│   └── EXAMPLE_OUTPUT.md         # 文字輸出範例
├── scripts/
│   ├── verify_citation.py        # 判決驗證（dr-lawbot API）
│   ├── generate_pptx.py          # 主產檔腳本
│   └── ppt_theme.py              # 樣式常數
├── examples/
│   ├── sample_outline.json       # JSON 大綱範例
│   └── sample_output.pptx        # 產檔結果範例
└── README.md
```

## FAQ

**Q1：Cowork 會把我的卷證檔案上傳嗎？**
A：不會。Cowork 在你本機的沙盒環境讀檔案。對外的網路請求**只有**呼叫 dr-lawbot API 做判決驗證（只傳判決字號，不傳卷證內容）。

**Q2：這要付費嗎？**
A：本 repo 完全免費開源。你需要：
- Claude Max / Team 訂閱（才有 Cowork）
- 不需要 API key、不需要 dr-lawbot 帳號（API 匿名可查）

**Q3：支援什麼案件類型？**
A：刑事、行政訴訟（稅務 / 處分類）、民事皆可。`knowledge/PPT_OUTLINE_TEMPLATE.md` 含三種版型。家事、強制執行、商業事件目前無專屬版型，需自行客製。

**Q4：掃描 PDF 處理得了嗎？**
A：Claude 有視覺能力可讀掃描頁。大卷建議**分頁指定**給 Claude 讀（例：「起訴書 p.68-72」、「卷 6 p.429-430」），不要一次塞整卷，速度與準確度都較好。

**Q5：我要怎麼改樣式？**
A：編輯 `scripts/ppt_theme.py` 的顏色、字型、字級常數。如果要改章節結構，改 `knowledge/PPT_OUTLINE_TEMPLATE.md`。

**Q6：可以用在英文案件嗎？**
A：目前設計以台灣司法實務為準，引用格式、dr-lawbot API、字型都是中文環境。英文需要大改。

**Q7：Claude 沒讀 `PROJECT_INSTRUCTIONS.md` 怎麼辦？**
A：在對話開頭明確說「請先讀 `PROJECT_INSTRUCTIONS.md` 作為系統指令」。Cowork 有時不會主動讀 repo 根目錄的檔案。

## 授權

MIT License. 歡迎 fork、改寫、分享。

## 致謝

- [法律偵探 (dr-lawbot.com)](https://dr-lawbot.com) — 提供 21.34M 判決免費搜尋 API
- [Anthropic Claude](https://claude.ai) — AI 模型與 Cowork 平台
- [python-pptx](https://python-pptx.readthedocs.io/) — pptx 產檔函式庫

## 貢獻

歡迎提交 Issue / PR：
- 新增書狀類型範本（家事、強執、商業事件）
- 改善引用格式
- 回報驗證 API edge case
- 優化 PPT 排版

---

**免責聲明**：本工具產出之 PPT 為律師作業輔助用途，不構成法律意見。判決驗證仰賴 dr-lawbot 資料庫，使用前請律師自行審核引用內容之正確性與適用性。
