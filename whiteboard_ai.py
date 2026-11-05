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
    "gemini": "gemini-3.5-flash-lite",
    "claude": "claude-haiku-5-5",
}

BATCH_SIZE = 8          # 第一段「找白板」一次 API 呼叫最多放幾張
DETECT_IMG_SIDE = 768   # 第一段只判斷是不是白板，小圖就夠
SINGLE_IMG_SIDE = 2000  # 第二段逐張讀白板：高解析度，手寫小字才看得清楚
DEFAULT_READ_MODEL = "gemini-3.6-flash"   # 讀白板（手寫）預設用手寫效果較好的模型

# 工地白板常見用語：給 AI 辨認潦草手寫時參考（可自行增加）
SITE_VOCAB = [
    "行政棟", "實驗棟", "滯洪池A", "滯洪池B.C", "油槽", "line", "GL", "超挖", "高程",
    "一階", "二階", "三階", "四階", "儀器高", "讀數", "設計", "實測", "開挖", "祛水", "基地",
]


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

【第一步：判斷每張是不是白板特寫】
工地習慣會特別拍一張「白板特寫」；其他檢查照片裡雖然也常有白板入鏡，但只是配角。
- has_board = true：只有「白板特寫」才算——白板／表格是照片主角，佔畫面大部分（大約一半以上），
  而且字清楚到能逐字讀出來。
- has_board = false：其他所有照片，包括「白板有入鏡，但只佔畫面一小部分、放在旁邊、或字模糊」的
  檢查照片——就算你讀得出部分字也一樣算 false。
  這類照片只要回 {{"photo": 編號, "has_board": false}}，不用填其他欄位。
{board_hint}
【第二步：只針對白板照判讀】
0. 先把白板上「所有手寫內容」逐字照抄到 raw_text（照原樣、照位置由上到下，看不清楚的字用 □），
   抄完再根據 raw_text 填其他欄位。不要跳過這一步。
1. 只讀「這次寫上去」的字跡。擦不乾淨的淡色殘影、上次留下的模糊字一律忽略。
2. 印好的表格：打勾（✓、V、塗黑）的選項才算數，沒勾的列忽略；印好的空白欄位以手寫內容為準。
3. 日期多為民國年（例如 115.10.14 = 2026-10-14），請轉成西元 YYYY-MM-DD。
4. 抽查位置照實抄寫並整理成乾淨文字（例如「實驗棟 line 7-12/A-C」）。
5. 數值保留白板上的單位與正負號（GL-2.225m、2.5cm），不要自己換算或四捨五入。
6. 看不清楚的字不要猜，對應欄位留空字串，並在 notes 說明。

【手寫判讀技巧】
- 手寫數字常見混淆：1/7、4/9、5/6/S、0/6、2/Z、3/8、小數點與逗號、7 與 1 加一橫。
  用「合理範圍」交叉判斷：開挖高程大多在 GL-0.5 ～ GL-20 m；超挖量通常 0 ～ 10 cm；
  儀器高、讀數多為 GL+ 幾十～幾百 cm；同一張白板上的設計值、實測值、讀數應該彼此接近。
- 「GL」手寫常看起來像「6L」「GI」「G2」，一律當作 GL；「cm」「m」「mm」常寫得很潦草，依數值大小判斷。
- 中文手寫潦草時，優先對照下面的工地常用詞與資料庫用語來辨認，不要硬猜成無關的字：
  {vocab}
- 日期：{date_hint}
- uncertain：把你沒把握的值列出來（例如 "實測 GL-2.27m 的 7 可能是 1"），沒有就給空陣列。

【對應資料庫】
- category_code：從下列 [C數字] 選最符合「抽查項目」的一個；都不符合就給 ""。
- title：只根據白板的「抽查/查驗項目」那一欄，整理成「…施工自主檢查」的寫法。
  不要加上白板最上面的「工程名稱」（例如「衛生福利部防疫中心興建工程」）或機關、廠商名稱。
  例如抽查項目寫「基地開挖及祛水工程施工查驗」或「…施工自檢」→ title 給「基地開挖及祛水工程施工自主檢查」。
  不管有沒有對應到資料庫都要給。
- items：把白板上「每一個檢查項目」依白板由上到下的順序列出來，一項一筆。
  * 以資料庫的「說明」為單位：白板上的內容若對應到資料庫裡不同的說明（例如「開挖高程」與「超挖」
    各有一個說明），就拆成兩筆；若資料庫只有一個說明涵蓋兩個數值，就合成一筆。
  * item_code：該分類底下最符合的 [C數字-數字]；都不符合就給 ""。
  * desc / design / result：最後要寫進報告的文字。
    有對應到 item 時，desc 用資料庫的「說明」原文；design、result 依資料庫的「格式」，
    把白板上的數值填進去（格式是空的就直接寫白板上的值）。
    沒對應到 item 時，用白板內容自己寫簡短的 desc / design / result。

【資料庫】
{db_listing}

{extra_hint}
只回傳 JSON，不要任何其他文字。photos 陣列要包含全部 {n} 張，格式：
{{
  "photos": [
    {{
      "photo": 1,
      "has_board": true,
      "raw_text": "",
      "board": {{
        "project": "", "location": "", "date": "", "inspection_item": "", "seq": "",
        "design_values": "", "measured_values": "", "signers": ""
      }},
      "category_code": "",
      "title": "",
      "items": [
        {{"item_code": "", "desc": "", "design": "", "result": ""}}
      ],
      "confidence": 0.0,
      "uncertain": [],
      "notes": ""
    }},
    {{"photo": 2, "has_board": false}}
  ]
}}
"""


def build_prompt(checks_db, n, current_category=None, force_board=False, photo_dates=None):
    hint = ""
    if current_category:
        hint = f"（使用者目前選的分類是「{current_category}」，若白板內容相符請優先使用此分類。）\n"
    board_hint = ""
    if force_board:
        board_hint = "（使用者已確認這張是白板照，請直接當白板判讀，has_board 給 true。）\n"
    dates = [d for d in (photo_dates or []) if d]
    if dates:
        date_hint = (f"這張照片的拍攝日期是 {dates[0]}（民國 {dates[0].year - 1911}.{dates[0].month}.{dates[0].day}），"
                     "白板上的日期通常是同一天或前後幾天，讀不清楚時以此判斷。")
    else:
        date_hint = "白板日期多為民國年，例如 115.10.14。"
    return PROMPT_TEMPLATE.format(
        n=n, db_listing=build_db_listing(checks_db), extra_hint=hint, board_hint=board_hint,
        vocab="、".join(SITE_VOCAB), date_hint=date_hint,
    )


DETECT_PROMPT = """下面有 {n} 張工地自主檢查照片，依序標為 照片1 ~ 照片{n}。
工地習慣會特別拍一張「白板特寫」；其他檢查照片裡雖然也常有白板入鏡，但只是配角。
請找出哪幾張是「白板特寫」：白板／表格是照片主角，佔畫面大部分（大約一半以上），而且字清楚可讀。
白板只是入鏡、放在旁邊、佔畫面一小部分或字模糊的檢查照片，都不算。
只回傳 JSON：{{"boards": [白板特寫的照片編號，例如 1, 4；沒有就給空陣列]}}
"""


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
# 讀手寫時讓模型「稍微想一下」準確度較好；判斷是不是白板則越快越好
_GEMINI_THINKING_LOW = [
    {"thinkingConfig": {"thinkingLevel": "low", "includeThoughts": True}},     # Gemini 3 系列
    {"thinkingConfig": {"thinkingBudget": 1024, "includeThoughts": True}},     # Gemini 2.5 系列
    {"thinkingConfig": {"includeThoughts": True}},
    {},
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
    retries = []  # 記錄被 Google 擋下重試的狀況（429 次數上限／503 伺服器忙碌）
    for attempt in range(4):
        r = requests.post(url, params={"key": cfg["api_key"], "alt": "sse"}, json=body,
                          stream=True, timeout=(10, 40))
        if stats is not None:
            stats["retries"] = retries
        if r.status_code == 200:
            break
        txt = r.text
        if r.status_code in (500, 502, 503) and attempt < 1:
            wait = 2
            retries.append(f"{r.status_code}（等 {wait:.0f} 秒）")
            time.sleep(wait)
            continue
        if stats is not None:
            stats.update(status=r.status_code, secs=round(time.time() - t0, 1), error=txt[:300])
        return r.status_code, "", txt

    texts, finish, usage, version, first_chunk = [], None, {}, "", None
    # Google 的串流回應沒有標明編碼，requests 會誤當成 Latin-1，中文變亂碼 → 一定要指定 UTF-8
    r.encoding = "utf-8"
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


def _call_gemini(cfg, parts, stats=None, think="minimal"):
    g_parts = [
        {"text": v} if kind == "text" else {"inline_data": {"mime_type": "image/jpeg", "data": v}}
        for kind, v in parts
    ]
    # 不要把 temperature 調低：Gemini 3 系列官方建議維持預設，調低容易讓輸出重複鬼打牆。
    base_cfg = {"responseMimeType": "application/json", "maxOutputTokens": 8192}
    ck = (cfg["model"], think)
    options = ([_gemini_thinking_ok[ck]] if ck in _gemini_thinking_ok
               else (_GEMINI_THINKING_LOW if think == "low" else _GEMINI_THINKING_OPTIONS))
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
        _gemini_thinking_ok[ck] = opt
        return text
    code, _, err = last
    resp = requests.models.Response()
    resp.status_code, resp._content = code, err.encode("utf-8")
    raise requests.HTTPError(response=resp)


# 額度用完／太慢的模型先暫停一段時間（同一個 API Key 的所有使用者共用）
_model_block = {}


def _secs_to_pacific_midnight():
    """Google 每日額度在美國太平洋時間午夜重置"""
    try:
        from zoneinfo import ZoneInfo
        now = datetime.datetime.now(ZoneInfo("America/Los_Angeles"))
        nxt = (now + datetime.timedelta(days=1)).replace(hour=0, minute=0, second=5, microsecond=0)
        return max(60, (nxt - now).total_seconds())
    except Exception:
        return 6 * 3600


def _switch_reason(e):
    """這個錯誤要不要換模型？要的話回傳 (暫停秒數, 原因)"""
    if isinstance(e, requests.HTTPError) and e.response is not None:
        code, body = e.response.status_code, e.response.text or ""
        if code == 429 and _is_daily_quota(body):
            return _secs_to_pacific_midnight(), "今天額度用完"
        if code == 429:
            return 60, "每分鐘次數已滿"
        if code == 404:
            return 3600, "找不到這個模型"
        if code >= 500:
            return 60, "伺服器忙碌"
        return None
    if isinstance(e, (requests.ReadTimeout, TimeoutError)):
        return 300, "回應太慢"
    if isinstance(e, requests.ConnectionError):
        return 30, "連線失敗"
    return None


def _short(m):
    return m.replace("gemini-", "")


def call_with_fallback(cfg, parts, stats=None, think="minimal"):
    """
    依序嘗試模型：使用者選的 → 清單裡其他的。額度用完、太慢、找不到模型時自動換下一個。
    回傳 (文字, 實際用的模型, 換模型的說明清單)
    """
    if cfg["provider"] != "gemini":
        return _call_claude(cfg, parts), cfg["model"], []
    order = [cfg["model"]] + [m for m in GEMINI_MODEL_CHOICES if m != cfg["model"]]
    now = time.time()
    ready = [m for m in order if _model_block.get(m, (0, ""))[0] <= now]
    notes = [f"{_short(m)} {_model_block[m][1]}，暫停使用" for m in order if m not in ready]
    if not ready:                      # 全部都在暫停中：還是全部再試一次
        ready, notes = order, []
    last = None
    for m in ready:
        try:
            text = _call_gemini({**cfg, "model": m}, parts, stats, think)
            _model_block.pop(m, None)
            return text, m, notes
        except Exception as e:
            why = _switch_reason(e)
            if not why:
                raise
            _model_block[m] = (time.time() + why[0], why[1])
            notes.append(f"{_short(m)} {why[1]}")
            last = e
    raise last


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
    return json.loads(m.group(0), strict=False)


def _resolve_codes(result, checks_db):
    """
    把 C3 / C3-5 轉回真正的分類名稱與項目索引。
    result["items"] 每一筆會加上 item_index（0-based，對不到就是 None）。
    """
    cats = list(checks_db.keys())
    result["category"] = None

    m = re.match(r"C(\d+)", str(result.get("category_code", "")).strip(), re.I)
    if m and 1 <= int(m.group(1)) <= len(cats):
        result["category"] = cats[int(m.group(1)) - 1]

    items = result.get("items")
    if not isinstance(items, list) or not items:
        # 相容舊格式：只有一組 desc/design/result
        items = [{k: result.get(k, "") for k in ("item_code", "desc", "design", "result")}]
    clean = []
    for it in items:
        if not isinstance(it, dict):
            continue
        it = {k: str(it.get(k) or "").strip() for k in ("item_code", "desc", "design", "result")}
        it["item_index"] = None
        m = re.match(r"C(\d+)-(\d+)", it["item_code"], re.I)
        if m:
            ci, ii = int(m.group(1)), int(m.group(2))
            if 1 <= ci <= len(cats) and 1 <= ii <= len(checks_db[cats[ci - 1]]):
                if result["category"] is None:
                    result["category"] = cats[ci - 1]
                if result["category"] == cats[ci - 1]:
                    it["item_index"] = ii - 1
        if it["desc"] or it["design"] or it["result"]:
            clean.append(it)
    result["items"] = clean
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


def _read_board(cfg, file_obj, checks_db, current_category, photo_date=None):
    """第二段：單張白板高解析度判讀（用讀手寫的模型、讓模型多想一下）"""
    try:
        parts = [("text", build_prompt(checks_db, 1, current_category, True, [photo_date])),
                 ("text", "照片1："), ("image", prepare_image_b64(file_obj, SINGLE_IMG_SIDE))]
        text, used, notes = call_with_fallback(cfg, parts, think="low")
        data = _parse_json(text)
        items = data.get("photos", data) if isinstance(data, dict) else data
        if isinstance(items, dict):
            items = [items]
        it = dict((items or [{}])[0] or {})
        it["has_board"] = True
        _resolve_codes(it, checks_db)
        it.update(ok=True, model_used=used, switch_notes=notes)
        return it
    except Exception as e:
        return {"ok": False, "error": _error_text(e)}


def _detect_boards(cfg, file_objs):
    """第一段：小圖＋快速模型，只找出哪幾張是白板特寫。回傳 (白板索引集合, 實際模型, 換模型說明)"""
    parts = [("text", DETECT_PROMPT.format(n=len(file_objs)))]
    for i, f in enumerate(file_objs, start=1):
        parts.append(("text", f"照片{i}："))
        parts.append(("image", prepare_image_b64(f, DETECT_IMG_SIDE)))
    text, used, notes = call_with_fallback(cfg, parts)
    data = _parse_json(text)
    raw = data.get("boards", []) if isinstance(data, dict) else data
    idx = set()
    for b in raw or []:
        try:
            k = int(b) - 1
        except (TypeError, ValueError):
            continue
        if 0 <= k < len(file_objs):
            idx.add(k)
    return idx, used, notes


def _read_cfg(cfg, read_model):
    if cfg["provider"] != "gemini":
        return cfg
    return {**cfg, "model": read_model or DEFAULT_READ_MODEL}


def recognize_many(cfg, file_objs, checks_db, current_category=None, read_model=None, photo_dates=None):
    """
    整組辨識（兩段式）：
      1. 找白板：每 BATCH_SIZE 張打包成一次呼叫、小圖、快速模型
      2. 讀白板：只把白板特寫逐張送出，高解析度＋手寫效果好的模型，多張同時進行
    回傳與輸入同順序的結果清單
    """
    n = len(file_objs)
    photo_dates = list(photo_dates or [None] * n)
    out = [None] * n
    chunks = [list(range(i, min(i + BATCH_SIZE, n))) for i in range(0, n, BATCH_SIZE)]

    def detect(chunk):
        try:
            idx, used, notes = _detect_boards(cfg, [file_objs[i] for i in chunk])
            return chunk, {chunk[k] for k in idx}, used, notes, None
        except Exception as e:
            return chunk, set(), None, [], _error_text(e)

    boards, det_notes = [], []
    with ThreadPoolExecutor(max_workers=max(1, min(4, len(chunks)))) as ex:
        for chunk, found, used, notes, err in ex.map(detect, chunks):
            det_notes += notes
            for i in chunk:
                if err:
                    out[i] = {"ok": False, "error": err}
                elif i in found:
                    boards.append(i)
                else:
                    out[i] = {"ok": True, "has_board": False, "model_used": used, "switch_notes": notes}

    rcfg = _read_cfg(cfg, read_model)
    if boards:
        with ThreadPoolExecutor(max_workers=min(4, len(boards))) as ex:
            res = ex.map(lambda i: _read_board(rcfg, file_objs[i], checks_db, current_category,
                                               photo_dates[i]), boards)
            for i, r in zip(boards, res):
                if r.get("ok"):
                    r["switch_notes"] = list(dict.fromkeys(det_notes + (r.get("switch_notes") or [])))
                out[i] = r
    return out


def recognize_one(cfg, file_obj, checks_db, current_category=None, read_model=None, photo_date=None):
    """單張辨識：使用者手動指定這張是白板，直接進第二段"""
    return _read_board(_read_cfg(cfg, read_model), file_obj, checks_db, current_category, photo_date)


GEMINI_MODEL_CHOICES = [
    "gemini-3.5-flash-lite",
    "gemini-3.6-flash",
    "gemini-3.1-flash-lite",
    "gemini-flash-latest",      # 最新 Flash（目前指向 3.8 Flash），免費版常排隊很久，放最後
]
GEMINI_MODEL_LABELS = {
    "gemini-3.5-flash-lite": "3.5 Flash-Lite（最快・預設）",
    "gemini-3.6-flash": "3.6 Flash（手寫效果好）",
    "gemini-3.1-flash-lite": "3.1 Flash-Lite（備用）",
    "gemini-flash-latest": "最新 Flash（手寫效果最好，但免費版可能很慢）",
}


def ping_models(cfg, models=GEMINI_MODEL_CHOICES):
    """同時對幾個模型送一句話，比較哪個回得快"""
    def one(m):
        stats, t0 = {}, time.time()
        try:
            _call_gemini({**cfg, "model": m}, [("text", '只回傳 JSON：{"ok": true}')], stats)
            return {"model": m, "ok": True, "secs": round(time.time() - t0, 1),
                    "version": stats.get("model_version", ""), "retries": stats.get("retries", [])}
        except Exception as e:
            return {"model": m, "ok": False, "secs": round(time.time() - t0, 1), "error": _error_text(e),
                    "retries": stats.get("retries", [])}
    with ThreadPoolExecutor(max_workers=len(models)) as ex:
        return list(ex.map(one, models))


def diagnose(cfg, checks_db, file_obj=None, read_model=None):
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
                return _call_gemini(_read_cfg(cfg, read_model), parts, st_, "low")
            return _call_claude(cfg, parts)
        run("2. 一張白板完整辨識（讀手寫模型）", one)
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
