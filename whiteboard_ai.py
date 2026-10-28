"""
白板照片 AI 辨識模組
- 讀取施工白板（手寫或印好的表格都可以）
- 自動對應 Google Sheets 資料庫中的「分類」與「說明」項目
- 依資料庫格式填好 說明 / 設計 / 實測

secrets.toml 設定範例（擇一）：

[ai]
provider = "gemini"            # 或 "claude"
api_key  = "你的 API KEY"
model    = "gemini-flash-latest"   # 可省略；claude 預設 claude-haiku-5-5
"""
import base64
import io
import json
import re
import time
import datetime
from concurrent.futures import ThreadPoolExecutor

import requests
from PIL import Image, ImageOps

DEFAULT_MODELS = {
    "gemini": "gemini-flash-latest",
    "claude": "claude-haiku-5-5",
}


# ------------------------------------------------------------------
# 設定
# ------------------------------------------------------------------
def get_ai_config(secrets):
    """從 st.secrets 取得 AI 設定，沒設定就回傳 None"""
    try:
        cfg = secrets["ai"]
        provider = str(cfg.get("provider", "gemini")).lower().strip()
        api_key = str(cfg["api_key"]).strip()
        model = str(cfg.get("model", DEFAULT_MODELS.get(provider, ""))).strip()
        if not api_key or provider not in DEFAULT_MODELS:
            return None
        return {"provider": provider, "api_key": api_key, "model": model}
    except Exception:
        return None


# ------------------------------------------------------------------
# 圖片前處理：辨識手寫需要較高解析度，比 Word 用的 800px 大
# ------------------------------------------------------------------
def prepare_image_b64(file_obj, max_side=1280):
    raw = file_obj.getvalue() if hasattr(file_obj, "getvalue") else file_obj.read()
    img = Image.open(io.BytesIO(raw))
    try:
        img = ImageOps.exif_transpose(img)
    except Exception:
        pass
    if img.mode != "RGB":
        img = img.convert("RGB")
    w, h = img.size
    scale = max_side / float(max(w, h))
    if scale < 1:
        img = img.resize((int(w * scale), int(h * scale)), Image.Resampling.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85)
    return base64.b64encode(buf.getvalue()).decode("ascii")


# ------------------------------------------------------------------
# Prompt
# ------------------------------------------------------------------
def build_db_listing(checks_db):
    """把資料庫壓成給 AI 看的編號清單：C3 = 分類，C3-5 = 該分類第 5 個說明"""
    lines = []
    for ci, (cat, items) in enumerate(checks_db.items(), start=1):
        lines.append(f"[C{ci}] {cat}")
        for ii, it in enumerate(items, start=1):
            lines.append(
                f"  [C{ci}-{ii}] 說明: {it['desc']} | 設計格式: {it['design'] or '(空)'} | 實測格式: {it['result'] or '(空)'}"
            )
    return "\n".join(lines)


PROMPT_TEMPLATE = """你是台灣營造工地的品管助理。照片是施工自主檢查／抽查用的白板或夾板表格，
內容可能是手寫，也可能是印好的表格再用筆填空。請讀取白板並對應到下面的檢查項目資料庫。

【判讀規則】
1. 只讀「這次寫上去」的字跡。白板上擦不乾淨的淡色殘影、上次留下的模糊字一律忽略。
2. 印好的表格：打勾（✓、V、塗黑）的選項才算數，沒勾的列忽略；印好的空白欄位以手寫內容為準。
3. 日期多為民國年（例如 115.10.14 = 2026-10-14），請轉成西元 YYYY-MM-DD。
4. 抽查位置照實抄寫並整理成乾淨文字（例如「實驗棟 line 7-12/A-C」）。
5. 數值保留白板上的單位與正負號（GL-2.225m、2.5cm），不要自己換算或四捨五入。
6. 看不清楚的字不要猜，對應欄位留空字串，並在 notes 說明。

【對應資料庫】
- category_code：從下列 [C數字] 選最符合「抽查項目」的一個；都不符合就給 ""。
- item_code：在該分類底下選最符合這張照片檢查內容的 [C數字-數字]；都不符合就給 ""。
- desc / design / result：請產生最後要寫進報告的文字。
  * 有對應到 item 時，desc 用資料庫的「說明」原文；design、result 依資料庫的「格式」，
    把白板上的數值填進去（格式是空的就直接寫白板上的值）。
  * 沒對應到 item 時，用白板內容自己寫簡短的 desc / design / result。
  * result 若白板有多個實測值（例如高程與超挖量），用「、」串在一起。

【資料庫】
{db_listing}

{extra_hint}
只回傳 JSON，不要任何其他文字，格式：
{{
  "has_board": true,
  "board": {{
    "project": "", "location": "", "date": "", "inspection_item": "", "seq": "",
    "design_values": "", "measured_values": "", "signers": ""
  }},
  "category_code": "",
  "item_code": "",
  "desc": "",
  "design": "",
  "result": "",
  "confidence": 0.0,
  "notes": ""
}}
"""


def build_prompt(checks_db, current_category=None):
    hint = ""
    if current_category:
        hint = f"（使用者目前選的分類是「{current_category}」，若白板內容相符請優先使用此分類。）\n"
    return PROMPT_TEMPLATE.format(db_listing=build_db_listing(checks_db), extra_hint=hint)


# ------------------------------------------------------------------
# 呼叫模型
# ------------------------------------------------------------------
# Gemini 的「思考」會讓回應慢好幾秒，讀白板用不到，盡量關掉。
# 新舊模型參數不同，依序嘗試，成功的那組記下來之後直接用。
_GEMINI_THINKING_OPTIONS = [
    {"thinkingConfig": {"thinkingLevel": "minimal"}},   # Gemini 3 系列
    {"thinkingConfig": {"thinkingBudget": 0}},          # Gemini 2.5 系列
    {},                                                  # 都不支援就不帶
]
_gemini_thinking_ok = {}


def _post_with_retry(url, retries=3, **kw):
    """遇到 429（免費額度每分鐘次數上限）或 5xx 時自動等一下重試"""
    for attempt in range(retries + 1):
        r = requests.post(url, **kw)
        if r.status_code in (429, 500, 502, 503) and attempt < retries:
            wait = float(r.headers.get("retry-after", 0) or 0) or (2 ** attempt * 2)
            time.sleep(min(wait, 20))
            continue
        return r
    return r


def _call_gemini(cfg, prompt, img_b64):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{cfg['model']}:generateContent"
    base_cfg = {"temperature": 0, "responseMimeType": "application/json"}
    options = ([_gemini_thinking_ok[cfg["model"]]] if cfg["model"] in _gemini_thinking_ok
               else _GEMINI_THINKING_OPTIONS)
    last = None
    for opt in options:
        body = {
            "contents": [{
                "parts": [
                    {"inline_data": {"mime_type": "image/jpeg", "data": img_b64}},
                    {"text": prompt},
                ]
            }],
            "generationConfig": {**base_cfg, **opt},
        }
        r = _post_with_retry(url, params={"key": cfg["api_key"]}, json=body, timeout=90)
        last = r
        if r.status_code == 400 and opt:
            continue  # 這個模型不吃這種 thinking 參數，換下一種
        r.raise_for_status()
        _gemini_thinking_ok[cfg["model"]] = opt
        data = r.json()
        return "".join(p.get("text", "") for p in data["candidates"][0]["content"]["parts"]
                       if not p.get("thought"))
    last.raise_for_status()


def _call_claude(cfg, prompt, img_b64):
    r = _post_with_retry(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": cfg["api_key"],
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        },
        json={
            "model": cfg["model"],
            "max_tokens": 1500,
            "temperature": 0,
            "messages": [{
                "role": "user",
                "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": img_b64}},
                    {"type": "text", "text": prompt},
                ],
            }],
        },
        timeout=90,
    )
    r.raise_for_status()
    data = r.json()
    return "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")


def _parse_json(text):
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text)
    m = re.search(r"\{.*\}", text, re.S)
    if not m:
        raise ValueError(f"AI 回傳不是 JSON：{text[:200]}")
    return json.loads(m.group(0))


def _resolve_codes(result, checks_db):
    """把 C3 / C3-5 轉回真正的分類名稱與項目索引"""
    cats = list(checks_db.keys())
    result["category"] = None
    result["item_index"] = None  # 0-based

    m = re.match(r"C(\d+)", str(result.get("category_code", "")).strip(), re.I)
    if m and 1 <= int(m.group(1)) <= len(cats):
        result["category"] = cats[int(m.group(1)) - 1]

    m = re.match(r"C(\d+)-(\d+)", str(result.get("item_code", "")).strip(), re.I)
    if m:
        ci, ii = int(m.group(1)), int(m.group(2))
        if 1 <= ci <= len(cats):
            items = checks_db[cats[ci - 1]]
            if 1 <= ii <= len(items):
                if result["category"] is None:
                    result["category"] = cats[ci - 1]
                if result["category"] == cats[ci - 1]:
                    result["item_index"] = ii - 1
    return result


def recognize_one(cfg, file_obj, checks_db, current_category=None):
    try:
        img_b64 = prepare_image_b64(file_obj)
        prompt = build_prompt(checks_db, current_category)
        if cfg["provider"] == "claude":
            text = _call_claude(cfg, prompt, img_b64)
        else:
            text = _call_gemini(cfg, prompt, img_b64)
        res = _parse_json(text)
        res = _resolve_codes(res, checks_db)
        res["ok"] = True
        return res
    except requests.HTTPError as e:
        detail = ""
        try:
            detail = e.response.text[:300]
        except Exception:
            pass
        return {"ok": False, "error": f"API 錯誤 {e.response.status_code if e.response is not None else ''}：{detail}"}
    except Exception as e:
        return {"ok": False, "error": str(e)}


def recognize_many(cfg, file_objs, checks_db, current_category=None, max_workers=8):
    """多張照片平行辨識，回傳順序與輸入相同"""
    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futures = [ex.submit(recognize_one, cfg, f, checks_db, current_category) for f in file_objs]
        return [f.result() for f in futures]


def parse_board_date(s):
    """'2026-10-14' / '115.10.14' / '115/10/14' → datetime.date，失敗回傳 None"""
    if not s:
        return None
    nums = [int(x) for x in re.findall(r"\d+", str(s))]
    if len(nums) < 3:
        return None
    y, mth, d = nums[:3]
    if y < 1911:
        y += 1911
    try:
        return datetime.date(y, mth, d)
    except ValueError:
        return None
