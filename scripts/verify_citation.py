"""呼叫 dr-lawbot.com 公開搜尋 API 驗證判決是否存在。

使用方式：
    from verify_citation import verify
    result = verify("108 台上 2027")
    if result["verified"]:
        print(result["title"], result["snippet"])

設計原則：
1. 只打 api.dr-lawbot.com/api/search（公開、匿名可用）
2. 失敗自動 retry 1 次（TLS 偶發不穩）
3. 比對字號吻合才算驗證通過
4. 回傳格式固定，呼叫端好處理
"""

from __future__ import annotations

import json
import re
import ssl
import time
import urllib.parse
import urllib.request
from dataclasses import dataclass, asdict
from typing import Optional

try:
    import certifi
    _SSL_CONTEXT = ssl.create_default_context(cafile=certifi.where())
except ImportError:
    _SSL_CONTEXT = ssl.create_default_context()


API_URL = "https://api.dr-lawbot.com/api/search"
TIMEOUT = 10  # 秒
RETRY_COUNT = 1
RETRY_WAIT = 2  # 秒


@dataclass
class VerifyResult:
    verified: bool
    query: str
    message: str                     # 給人看的狀態描述
    doc_id: Optional[str] = None
    jyear: Optional[str] = None
    jcase: Optional[str] = None
    jno: Optional[str] = None
    jdate: Optional[str] = None
    title: Optional[str] = None
    court_level: Optional[str] = None
    case_type: Optional[str] = None
    doc_type_label: Optional[str] = None
    main_preview: Optional[str] = None
    snippet: Optional[str] = None
    search_url: Optional[str] = None  # 讓 PPT 附可點連結

    def to_dict(self):
        return asdict(self)


def parse_citation(text: str) -> Optional[tuple[str, str, str]]:
    """從使用者輸入解析 (年度, 字, 號)。

    支援格式：
        108 台上 2027
        108年度台上字第2027號
        最高法院 108 年度台上字第 2027 號
        107 判字 369
        109 上字 993
    """
    text = text.strip()
    patterns = [
        r"(\d+)\s*年度\s*([^\d\s]+?)\s*字第\s*(\d+)\s*號",
        r"(\d+)\s+([^\d\s]+?)\s+(\d+)",
        r"(\d+)([^\d\s]+?)(\d+)",
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            year, case_word, no = m.groups()
            case_word = case_word.replace("字", "").strip()
            return year, case_word, no
    return None


def _fetch(query: str) -> dict:
    """原始 API 呼叫，回傳 JSON dict。失敗 raise。"""
    url = f"{API_URL}?q={urllib.parse.quote(query)}"
    req = urllib.request.Request(
        url,
        headers={"Accept": "application/json", "User-Agent": "legal-ppt-generator/0.1"},
    )
    with urllib.request.urlopen(req, timeout=TIMEOUT, context=_SSL_CONTEXT) as resp:
        raw = resp.read().decode("utf-8")
    return json.loads(raw)


def verify(citation: str) -> VerifyResult:
    """驗證單一判決字號。

    citation 可為使用者原始輸入字串，本函式會嘗試解析後查詢並比對。
    """
    search_url = f"{API_URL}?q={urllib.parse.quote(citation)}"
    parsed = parse_citation(citation)

    last_err: Exception | None = None
    for attempt in range(RETRY_COUNT + 1):
        try:
            data = _fetch(citation)
            break
        except Exception as e:
            last_err = e
            if attempt < RETRY_COUNT:
                time.sleep(RETRY_WAIT)
    else:
        return VerifyResult(
            verified=False,
            query=citation,
            message=f"API 呼叫失敗（已重試）：{last_err}",
            search_url=search_url,
        )

    results = data.get("results", [])
    if not results:
        return VerifyResult(
            verified=False,
            query=citation,
            message="判決資料庫查無此字號（dr-lawbot 收錄民國 100 年起判決）。",
            search_url=search_url,
        )

    # 若解析到結構化字號，嚴格比對 jyear/jcase/jno
    if parsed:
        year, case_word, no = parsed
        for r in results:
            if (str(r.get("jyear")) == year
                and str(r.get("jcase")) == case_word
                and str(r.get("jno")) == no):
                return _to_result(r, citation, search_url, "驗證成功，字號完全吻合。")
        # 沒有完全吻合的
        return VerifyResult(
            verified=False,
            query=citation,
            message=f"查無完全吻合 {year} {case_word} {no} 的判決（共 {len(results)} 筆相近結果）。",
            search_url=search_url,
        )

    # 沒解析到結構 → 取第一筆但標示「非精確比對」
    return _to_result(
        results[0],
        citation,
        search_url,
        "查詢成功，但輸入字串未能解析為結構化字號，結果為第一筆相關結果，請人工核對。",
    )


def _to_result(r: dict, query: str, search_url: str, message: str) -> VerifyResult:
    return VerifyResult(
        verified=True,
        query=query,
        message=message,
        doc_id=r.get("doc_id"),
        jyear=str(r.get("jyear")) if r.get("jyear") else None,
        jcase=r.get("jcase"),
        jno=str(r.get("jno")) if r.get("jno") else None,
        jdate=r.get("jdate"),
        title=r.get("title"),
        court_level=r.get("court_level"),
        case_type=r.get("case_type"),
        doc_type_label=r.get("doc_type_label"),
        main_preview=r.get("main_preview"),
        snippet=r.get("snippet"),
        search_url=search_url,
    )


# ─── 法院代號對照 ─────────────────────────────────────────

COURT_NAMES = {
    "TPSM": "最高法院（刑事）",
    "TPSV": "最高法院（民事）",
    "TPAA": "最高行政法院",
    "TPHV": "台灣高等法院",
    "TCHV": "台中高分院",
    "TNHV": "台南高分院",
    "KSHV": "高雄高分院",
    "HLHV": "花蓮高分院",
    "TPBA": "台北高等行政法院",
    "TCBA": "台中高等行政法院",
    "KSBA": "高雄高等行政法院",
}


def court_from_doc_id(doc_id: str) -> str | None:
    if not doc_id:
        return None
    prefix = doc_id.split(",")[0] if "," in doc_id else doc_id[:4]
    return COURT_NAMES.get(prefix)


def format_full_citation(r: VerifyResult) -> str:
    """把 VerifyResult 格式化成正式書狀引用字串。"""
    court = court_from_doc_id(r.doc_id or "") or "（法院未識別）"
    doc_type = r.doc_type_label or "判決"
    case_type_map = {"criminal": "刑事", "civil": "民事", "administrative": ""}
    case_type = case_type_map.get(r.case_type or "", "")
    return f"{court} {r.jyear} 年度{r.jcase}字第 {r.jno} 號{case_type}{doc_type}"


# ─── CLI ─────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("用法：python verify_citation.py '108 台上 2027'")
        sys.exit(1)
    q = " ".join(sys.argv[1:])
    result = verify(q)
    print(json.dumps(result.to_dict(), ensure_ascii=False, indent=2))
    if result.verified:
        print("\n正式引用：", format_full_citation(result))
