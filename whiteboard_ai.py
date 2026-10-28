"""
白板照片 AI 辨識模組
- 一組照片一次送出（一次 API 呼叫），AI 自己挑出哪幾張是白板
- 只辨識白板照片，其他檢查照片不動
- 白板內容對應 Google Sheets 資料庫的「分類」與「說明」，依格式填好 說明 / 設計 / 實測

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

BATCH_SIZE = 8          # 一次 API 呼叫最多放幾張
BATCH_IMG_SIDE = 1024   # 整組送出時每張的解析度
SINGLE_IMG_SIDE = 1280  # 單張重新辨識時用較高解析度


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
# 圖片前處理
# ------------------------------------------------------------------
def prepare_image_b64(file_obj, max_side=SINGLE_IMG_SIDE):
    raw = file_obj.getvalue() if hasattr(file_obj, "getvalue") else file_obj.read()
    img = Image.open(io.BytesIO(raw))
    try:
        img = ImageOps.exif_transpose(img)
    except Exception:
        pass
    if img.mode != "RGB":
        img = img.convert("RGB")
    img.thumbnail((max_side, max_side), Image.Resampling.LANCZOS)
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


PROMPT_TEMPLATE = """你是台灣營造工地的品管助理。下面有 {n} 張施工自主檢查的照片，依序標為 照片1 ~ 照片{n}。
其中有些是「白板照」（拿著寫有工程名稱、抽查位置、日期、檢查項目、設計值/實測值的白板或夾板表格），
其他是一般的「檢查照片」（量尺、儀器、現場施工狀況等，沒有白板，或白板很小看不清楚）。

【第一步：判斷每張是不是白板照】
- has_board = true：畫面主體是白板／表格，而且上面的字讀得出來。
- has_board = false：其他所有照片。這類照片只要回 {{"photo": 編號, "has_board": false}}，不用填其他欄位。
{board_hint}
【第二步：只針對白板照判讀】
1. 只讀「這次寫上去」的字跡。擦不乾淨的淡色殘影、上次留下的模糊字一律忽略。
2. 印好的表格：打勾（✓、V、塗黑）的選項才算數，沒勾的列忽略；印好的空白欄位以手寫內容為準。
3. 日期多為民國年（例如 115.10.14 = 2026-10-14），請轉成西元 YYYY-MM-DD。
4. 抽查位置照實抄寫並整理成乾淨文字（例如「實驗棟 line 7-12/A-C」）。
5. 數值保留白板上的單位與正負號（GL-2.225m、2.5cm），不要自己換算或四捨五入。
6. 看不清楚的字不要猜，對應欄位留空字串，並在 notes 說明。

【對應資料庫】
- category_code：從下列 [C數字] 選最符合「抽查項目」的一個；都不符合就給 ""。
- item_code：在該分類底下選最符合這張白板檢查內容的 [C數字-數字]；都不符合就給 ""。
- desc / design / result：最後要寫進報告的文字。
  * 有對應到 item 時，desc 用資料庫的「說明」原文；design、result 依資料庫的「格式」，
    把白板上的數值填進去（格式是空的就直接寫白板上的值）。
  * 沒對應到 item 時，用白板內容自己寫簡短的 desc / design / result。
  * result 若白板有多個實測值（例如高程與超挖量），用「、」串在一起。

【資料庫】
{db_listing}

{extra_hint}
只回傳 JSON，不要任何其他文字。photos 陣列要包含全部 {n} 張，格式：
{{
  "photos": [
    {{
      "photo": 1,
      "has_board": true,
      "board": {{
        "project": "", "location": "", "date": "", "inspection_item": "", "seq": "",
        "design_values": "", "measured_values": "", "signers": ""
      }},
      "category_code": "", "item_code": "",
      "desc": "", "design": "", "result": "",
      "confidence": 0.0,
      "notes": ""
    }},
    {{"photo": 2, "has_board": false}}
  ]
}}
"""


def build_prompt(checks_db, n, current_category=None, force_board=False):
    hint = ""
    if current_category:
        hint = f"（使用者目前選的分類是「{current_category}」，若白板內容相符請優先使用此分類。）\n"
    board_hint = ""
    if force_board:
        board_hint = "（使用者已確認這張是白板照，請直接當白板判讀，has_board 給 true。）\n"
    return PROMPT_TEMPLATE.format(
        n=n, db_listing=build_db_listing(checks_db), extra_hint=hint, board_hint=board_hint
    )


# ------------------------------------------------------------------
# 呼叫模型（parts = [("text", str) | ("image", b64), ...]）
# ------------------------------------------------------------------
# Gemini 的「思考」會讓回應慢好幾秒，讀白板用不到，盡量關掉。
# 新舊模型參數不同，依序嘗試，成功的那組記下來之後直接用。
# includeThoughts：讓模型思考時也持續回傳內容，連線才不會因為「一直沒回應」被判逾時
_GEMINI_THINKING_OPTIONS = [
    {"thinkingConfig": {"thinkingLevel": "minimal", "includeThoughts": True}},  # Gemini 3 系列
    {"thinkingConfig": {"thinkingBudget": 0}},                                 # Gemini 2.5 系列
    {"thinkingConfig": {"includeThoughts": True}},                             # 其他會思考的模型
    {},                                                                         # 都不支援就不帶
]
_gemini_thinking_ok = {}


def _is_daily_quota(text):
    t = (text or "").lower()
    return "perday" in t or "per day" in t or "requests_per_day" in t


def _post_with_retry(url, retries=3, **kw):
    """遇到 429（免費額度每分鐘次數上限）或 5xx 時自動等一下重試"""
    for attempt in range(retries + 1):
        r = requests.post(url, **kw)
        if r.status_code == 429 and _is_daily_quota(r.text):
            return r  # 今天的額度用完了，重試也沒用，直接回報
        if r.status_code in (429, 500, 502, 503) and attempt < retries:
            wait = float(r.headers.get("retry-after", 0) or 0) or (2 ** attempt * 2)
            time.sleep(min(wait, 20))
            continue
        return r
    return r


def _gemini_stream(cfg, g_parts, gen_cfg, stats=None, deadline=150):
    """
    用串流方式呼叫 Gemini：模型一邊想一邊回傳，連線不會因為「思考太久沒回應」被判逾時。
    回傳 (status_code, 文字, 錯誤內容)；stats 會記錄耗時、模型版本、token 用量等診斷資訊。
    """
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{cfg['model']}:streamGenerateContent"
    body = {"contents": [{"parts": g_parts}], "generationConfig": gen_cfg}
    t0 = time.time()
    for attempt in range(4):
        r = requests.post(url, params={"key": cfg["api_key"], "alt": "sse"}, json=body,
                          stream=True, timeout=(15, 60))
        if r.status_code == 200:
            break
        txt = r.text
        if r.status_code in (429, 500, 502, 503) and attempt < 3 and not _is_daily_quota(txt):
            time.sleep(min(float(r.headers.get("retry-after", 0) or 0) or 2 ** attempt * 2, 20))
            continue
        if stats is not None:
            stats.update(status=r.status_code, secs=round(time.time() - t0, 1), error=txt[:300])
        return r.status_code, "", txt

    texts, finish, usage, version, first_chunk = [], None, {}, "", None
    for line in r.iter_lines(decode_unicode=True):
        if time.time() - t0 > deadline:
            r.close()
            raise TimeoutError(f"AI 超過 {deadline} 秒還沒回完")
        if not line or not line.startswith("data:"):
            continue
        if first_chunk is None:
            first_chunk = round(time.time() - t0, 1)
        try:
            chunk = json.loads(line[5:].strip())
        except json.JSONDecodeError:
            continue
        version = chunk.get("modelVersion", version)
        usage = chunk.get("usageMetadata", usage)
        for cand in chunk.get("candidates", [])[:1]:
            finish = cand.get("finishReason", finish)
            for prt in (cand.get("content") or {}).get("parts", []):
                if not prt.get("thought") and prt.get("text"):
                    texts.append(prt["text"])
    if stats is not None:
        stats.update(status=200, secs=round(time.time() - t0, 1), first_chunk=first_chunk,
                     model_version=version, finish=finish,
                     thoughts_tokens=usage.get("thoughtsTokenCount", 0),
                     output_tokens=usage.get("candidatesTokenCount", 0),
                     input_tokens=usage.get("promptTokenCount", 0))
    if finish == "MAX_TOKENS":
        raise ValueError("AI 回覆異常（輸出過長被截斷），請再按一次")
    if finish in ("SAFETY", "RECITATION", "PROHIBITED_CONTENT"):
        raise ValueError(f"AI 拒絕回覆（{finish}），請改用單張 🤖 辨識")
    return 200, "".join(texts), ""


def _call_gemini(cfg, parts, stats=None):
    g_parts = [
        {"text": v} if kind == "text" else {"inline_data": {"mime_type": "image/jpeg", "data": v}}
        for kind, v in parts
    ]
    # 不要把 temperature 調低：Gemini 3 系列官方建議維持預設，調低容易讓輸出重複鬼打牆。
    base_cfg = {"responseMimeType": "application/json", "maxOutputTokens": 8192}
    options = ([_gemini_thinking_ok[cfg["model"]]] if cfg["model"] in _gemini_thinking_ok
               else _GEMINI_THINKING_OPTIONS)
    last = (0, "", "")
    for opt in options:
        st_ = {} if stats is None else stats
        st_.clear()
        code, text, err = _gemini_stream(cfg, g_parts, {**base_cfg, **opt}, st_)
        st_["thinking_option"] = json.dumps(opt) if opt else "(預設)"
        last = (code, text, err)
        if code == 400 and opt and "api_key" not in err.lower() and "api key" not in err.lower():
            continue  # 這個模型不吃這種 thinking 參數，換下一種
        if code != 200:
            break
        _gemini_thinking_ok[cfg["model"]] = opt
        return text
    code, _, err = last
    resp = requests.models.Response()
    resp.status_code, resp._content = code, err.encode("utf-8")
    raise requests.HTTPError(response=resp)


def _call_claude(cfg, parts):
    c_parts = [
        {"type": "text", "text": v} if kind == "text"
        else {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": v}}
        for kind, v in parts
    ]
    r = _post_with_retry(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": cfg["api_key"],
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        },
        json={
            "model": cfg["model"],
            "max_tokens": 4000,
            "temperature": 0,
            "messages": [{"role": "user", "content": c_parts}],
        },
        timeout=(10, 90),
    )
    r.raise_for_status()
    data = r.json()
    return "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")


def _parse_json(text):
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text)
    m = re.search(r"[\{\[].*[\}\]]", text, re.S)
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


def _error_text(e):
    """把 API 錯誤翻成看得懂的中文"""
    if isinstance(e, requests.HTTPError) and e.response is not None:
        code, body = e.response.status_code, e.response.text or ""
        low = body.lower()
        if code == 429 and _is_daily_quota(body):
            return ("⛔ 今天的 AI 免費額度已用完（約台灣時間下午 3～4 點重置）。"
                    "今天請先手動填寫，或在 Google AI Studio 為這個專案啟用帳單。")
        if code == 429:
            return "⏳ 每分鐘使用次數超過上限，已自動重試仍失敗，請等 1 分鐘後再按一次。"
        if "api_key_invalid" in low or "api key not valid" in low or code in (401, 403):
            return "🔑 API Key 無效或沒有權限，請檢查 Streamlit Secrets 裡的 api_key 是否貼對。"
        if "credit balance" in low or "billing" in low:
            return "💳 API 帳戶餘額不足或尚未設定帳單，請到 API 後台確認。"
        if code == 404:
            return "❓ 找不到這個模型名稱，請檢查 Secrets 裡的 model（或直接刪掉那一行用預設值）。"
        if code >= 500:
            return "🌐 AI 服務暫時忙碌或故障，請稍後再試。"
        return f"API 錯誤 {code}：{body[:200]}"
    if isinstance(e, requests.ConnectTimeout):
        return "🌐 連不上 Google AI 伺服器（連線逾時），請稍後再試。"
    if isinstance(e, (requests.ReadTimeout, TimeoutError)):
        return "⏱️ AI 回應太慢逾時。請按側邊欄「🩺 測試 AI 連線」看卡在哪裡。"
    if isinstance(e, requests.Timeout):
        return "🌐 連線逾時，可能是網路不穩，請再按一次。"
    if isinstance(e, requests.ConnectionError):
        return "🌐 連不到 AI 服務，請檢查網路後再試。"
    if isinstance(e, json.JSONDecodeError):
        return "AI 回傳格式錯誤，請再按一次"
    return str(e)


def _recognize_batch(cfg, file_objs, checks_db, current_category, img_side, force_board):
    """一次 API 呼叫處理多張，回傳與輸入同順序的結果清單"""
    n = len(file_objs)
    try:
        parts = [("text", build_prompt(checks_db, n, current_category, force_board))]
        for i, f in enumerate(file_objs, start=1):
            parts.append(("text", f"照片{i}："))
            parts.append(("image", prepare_image_b64(f, img_side)))
        call = _call_claude if cfg["provider"] == "claude" else _call_gemini
        data = _parse_json(call(cfg, parts))
        items = data.get("photos", data) if isinstance(data, dict) else data
        if isinstance(items, dict):
            items = [items]

        out = [{"ok": True, "has_board": False} for _ in range(n)]
        for pos, it in enumerate(items or []):
            if not isinstance(it, dict):
                continue
            try:
                idx = int(it.get("photo", pos + 1)) - 1
            except (TypeError, ValueError):
                idx = pos
            if not 0 <= idx < n:
                continue
            it = dict(it)
            it["has_board"] = bool(it.get("has_board", False))
            if it["has_board"]:
                _resolve_codes(it, checks_db)
            it["ok"] = True
            out[idx] = it
        return out
    except Exception as e:
        err = _error_text(e)
        return [{"ok": False, "error": err} for _ in range(n)]


def recognize_many(cfg, file_objs, checks_db, current_category=None):
    """整組辨識：每 BATCH_SIZE 張打包成一次呼叫，多包平行送出"""
    chunks = [file_objs[i:i + BATCH_SIZE] for i in range(0, len(file_objs), BATCH_SIZE)]
    with ThreadPoolExecutor(max_workers=max(1, min(4, len(chunks)))) as ex:
        futures = [ex.submit(_recognize_batch, cfg, c, checks_db, current_category,
                             BATCH_IMG_SIDE, False) for c in chunks]
        results = []
        for f in futures:
            results.extend(f.result())
    return results


def recognize_one(cfg, file_obj, checks_db, current_category=None):
    """單張辨識：使用者手動指定這張是白板，用較高解析度"""
    return _recognize_batch(cfg, [file_obj], checks_db, current_category,
                            SINGLE_IMG_SIDE, True)[0]


def diagnose(cfg, checks_db, file_obj=None):
    """
    🩺 測試 AI 連線：一步一步測，回傳每一步的結果，用來找出慢在哪裡
    1. 純文字小請求：測 API Key、網路、模型名稱
    2. 一張照片完整辨識（有上傳照片才測）：測實際辨識速度
    """
    steps = []

    def run(name, fn):
        stats, t0 = {}, time.time()
        try:
            out = fn(stats)
            stats.setdefault("secs", round(time.time() - t0, 1))
            steps.append({"step": name, "ok": True, "result": str(out)[:200], **stats})
            return True
        except Exception as e:
            stats.setdefault("secs", round(time.time() - t0, 1))
            steps.append({"step": name, "ok": False, **stats, "error": _error_text(e)})
            return False

    if cfg["provider"] == "gemini":
        ok = run("1. 文字測試（API Key／網路／模型）",
                 lambda st_: _call_gemini(cfg, [("text", '只回傳 JSON：{"ok": true}')], st_))
    else:
        ok = run("1. 文字測試（API Key／網路／模型）",
                 lambda st_: _call_claude(cfg, [("text", '只回傳 JSON：{"ok": true}')]))
    if ok and file_obj is not None:
        def one(st_):
            img = prepare_image_b64(file_obj, SINGLE_IMG_SIDE)
            parts = [("text", build_prompt(checks_db, 1, None, True)), ("text", "照片1："), ("image", img)]
            if cfg["provider"] == "gemini":
                return _call_gemini(cfg, parts, st_)
            return _call_claude(cfg, parts)
        run("2. 一張白板完整辨識", one)
    return steps


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
