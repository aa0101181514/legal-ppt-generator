# 法律書狀 PPT 助手

> 免費、開源的 Claude.ai Project 設定檔，幫律師把卷證資料變成**法庭風格、判決引用零幻覺**的答辯 PPT 大綱。

## 這是什麼？

一份寫給 Claude.ai 用的「指令 + 範本」檔案組。你只要：

1. 把這份指令貼到自己的 Claude.ai Project
2. 上傳案件 PDF（起訴書、卷證、判決）
3. 跟 Claude 說「幫我做法庭風格答辯 PPT」

Claude 就會幫你產出一份**每個判決引用都經過法律偵探資料庫驗證**的 Markdown 大綱，你再複製到 Gamma.app 或 Google Slides 排版成 PPT。

## 為什麼需要這個？

- ❌ Claude 會「幻覺」判決字號與判決內容
- ❌ 手動查每個判決很花時間
- ❌ 格式（「卷 X 頁（即起訴書 Y 頁）」）易寫錯
- ✅ 這份 Project 強制 Claude 呼叫 [dr-lawbot.com](https://dr-lawbot.com) API 驗證每個判決
- ✅ 章節結構照法庭實務「先程序後實體」
- ✅ 完全免費，用你自己的 Claude.ai 訂閱額度

---

## 快速開始（3 步驟，約 5 分鐘）

### 步驟 1：建立 Claude.ai Project

1. 登入 [claude.ai](https://claude.ai)（需要 Pro 或 Team 訂閱才能用 Projects）
2. 左側選單點 **Projects** → **+ Create Project**
3. Project 名稱填：`法律書狀 PPT 助手`

### 步驟 2：設定 Custom Instructions

1. 在 Project 頁面，找到 **Set project instructions**（或 **Custom Instructions**）
2. 打開本 repo 的 [`PROJECT_INSTRUCTIONS.md`](./PROJECT_INSTRUCTIONS.md)
3. **複製全部內容** → 貼到 Claude.ai 的 Custom Instructions 欄位
4. 點 **Save**

### 步驟 3：上傳知識檔案

在 Project 頁面找到 **Project knowledge**（或 **Knowledge base**），依序上傳本 repo 的這三個檔案：

- [`knowledge/PPT_OUTLINE_TEMPLATE.md`](./knowledge/PPT_OUTLINE_TEMPLATE.md)
- [`knowledge/CITATION_RULES.md`](./knowledge/CITATION_RULES.md)
- [`knowledge/EXAMPLE_OUTPUT.md`](./knowledge/EXAMPLE_OUTPUT.md)

**怎麼下載這些檔案？**

- 點上方任一個連結 → 到 GitHub 頁面 → 右上角 **Download raw file** 圖示
- 或整個 repo 下載：點 repo 首頁的綠色 **Code** 按鈕 → **Download ZIP**

---

## 怎麼使用（日常流程）

1. 進入你剛建好的 Project
2. 點 **Start new chat in project** 開新對話
3. 把案件 PDF 拖拉上傳（起訴書、卷證、判決、筆錄都可以）
4. 打字：
   > 幫我做一份法庭風格的答辯簡報大綱。被告是 OOO，起訴罪名 OOO，核心抗辯方向是 OOO。

5. Claude 會：
   - 讀你上傳的檔案
   - 整理爭點
   - **自動呼叫 dr-lawbot API 驗證每一個判決引用**
   - 產出 Markdown 大綱

6. 複製 Claude 給你的 Markdown → 貼到：
   - [**Gamma.app**](https://gamma.app)（推薦，免費版可用，自動排成 PPT）
   - **Google Slides**（手動或用 Markdown 匯入外掛）
   - **Canva**（先轉 PPT 再匯入）

---

## 產出範例

請看 [`knowledge/EXAMPLE_OUTPUT.md`](./knowledge/EXAMPLE_OUTPUT.md)。

---

## 關於「判決零幻覺」

本 Project 強制 Claude 每次引用判決時，呼叫：

```
https://api.dr-lawbot.com/api/search?q={判決字號}
```

API 回傳 JSON 包含判決真實標題、案由、主文節錄。Claude 依此驗證：

- ✅ 判決存在 → 採用
- ❌ 查無結果 → 標「未驗證」或拒絕引用
- ❌ 字號吻合但 `snippet` 不支持你的論點 → 換判決

**法律偵探資料庫收錄民國 100 年起的判決**，更早的舊判例（例如「79 判 471」）查無屬正常，Claude 會主動告知。

## FAQ

**Q1：這個要付費嗎？**
A：本 repo 完全免費開源。但你需要：
- **Claude.ai 訂閱**（Pro / Team / Max，才能用 Projects 功能）
- 不需要 API key、不需要 dr-lawbot.com 帳號（API 匿名可查）

**Q2：為什麼 Claude.ai 不能直接產 .pptx 檔？**
A：Claude.ai 對話視窗無法產出 PowerPoint 檔案格式。我們輸出 Markdown，讓你用 Gamma.app 等工具一鍵轉 PPT，品質與彈性都更好。

**Q3：可以用在民事、行政訴訟嗎？**
A：可以。`PPT_OUTLINE_TEMPLATE.md` 含刑事、行政、民事三種版型。

**Q4：我要怎麼檢查 Claude 有沒有真的驗證判決？**
A：Claude 產出時會附 dr-lawbot 的 API 連結，你可以點進去看 JSON 確認判決存在；或到 [dr-lawbot.com](https://dr-lawbot.com) 手動搜尋。

**Q5：可以客製成我們事務所的風格嗎？**
A：可以。把 repo fork 一份，修改 `PROJECT_INSTRUCTIONS.md` 第六節「回覆風格」或 `PPT_OUTLINE_TEMPLATE.md` 的章節結構即可。

**Q6：卷頁引用 Claude 怎麼知道要寫什麼？**
A：使用者需自行在對話中告知卷頁（或上傳有頁碼的 PDF）。Claude **不會自動推算頁碼**，未提供則標「（頁碼待補）」。

## 授權

MIT License. 歡迎 fork、改寫、分享。

## 致謝

- [法律偵探 (dr-lawbot.com)](https://dr-lawbot.com) — 提供 21.34M 判決免費搜尋 API
- [Anthropic Claude](https://claude.ai) — AI 模型與 Project 平台
- [Gamma.app](https://gamma.app) — Markdown to PPT

## 貢獻

歡迎提交 Issue 或 PR：
- 新增其他類型書狀範本（家事、強制執行、商業事件）
- 改善引用格式規則
- 回報驗證 API 的 edge case

---

**免責聲明**：本工具產出之大綱為律師作業輔助用途，不構成法律意見。判決驗證仰賴 dr-lawbot 資料庫，使用前請律師自行審核引用內容之正確性與適用性。
