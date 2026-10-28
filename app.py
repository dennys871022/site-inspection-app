import streamlit as st
try:
    import pkg_resources
except ImportError:
    import setuptools

from docx import Document
from docx.shared import Cm, Pt
from docx.oxml.ns import qn
from docxcompose.composer import Composer
from PIL import Image
import io
import datetime
from datetime import timedelta, timezone
import os
import zipfile
import pandas as pd
import smtplib
import re
import json
from urllib.parse import unquote
import streamlit.components.v1 as components
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from whiteboard_ai import get_ai_config, recognize_many, recognize_one, parse_board_date

# ==========================================
# 0. 雲端資料庫設定
# ==========================================
GOOGLE_SHEETS_CSV_URL = "https://docs.google.com/spreadsheets/d/1ubR0wOJkOhA4IYyQ_Qq-LUldKwkEj084N45Ym04sKU8/export?format=csv"

# ==========================================
# 1. 核心功能函式庫
# ==========================================

def get_taiwan_date():
    utc_now = datetime.datetime.now(timezone.utc)
    return (utc_now + timedelta(hours=8)).date()

def get_paragraph_style(paragraph):
    style = {}
    if paragraph.runs:
        run = paragraph.runs[0]
        style['font_name'] = run.font.name
        style['font_size'] = run.font.size
        style['bold'] = run.bold
        style['italic'] = run.italic
        style['underline'] = run.underline
        style['color'] = run.font.color.rgb
        try:
            rPr = run._element.rPr
            if rPr is not None and rPr.rFonts is not None:
                style['eastAsia'] = rPr.rFonts.get(qn('w:eastAsia'))
        except: pass
    return style

def apply_style_to_run(run, style):
    if not style: return
    if style.get('font_name'): run.font.name = style.get('font_name')
    if style.get('font_size'): run.font.size = style['font_size']
    if style.get('bold') is not None: run.bold = style['bold']
    if style.get('italic') is not None: run.italic = style['italic']
    if style.get('underline') is not None: run.underline = style['underline']
    if style.get('color'): run.font.color.rgb = style['color']
    if style.get('eastAsia'):
        run._element.rPr.rFonts.set(qn('w:eastAsia'), style['eastAsia'])
    elif style.get('font_name') == 'Times New Roman':
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

def compress_image(image_file, max_width=800):
    img = Image.open(image_file)
    if img.mode == 'RGBA': img = img.convert('RGB')
    try:
        from PIL import ImageOps
        img = ImageOps.exif_transpose(img)
    except: pass
    ratio = max_width / float(img.size[0])
    if ratio < 1:
        h_size = int((float(img.size[1]) * float(ratio)))
        img = img.resize((max_width, h_size), Image.Resampling.LANCZOS)
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='JPEG', quality=75)
    img_byte_arr.seek(0)
    return img_byte_arr

def replace_text_content(doc, replacements):
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_paragraph_pure(paragraph, replacements)
    for paragraph in doc.paragraphs:
        replace_paragraph_pure(paragraph, replacements)

def replace_paragraph_pure(paragraph, replacements):
    if not paragraph.text: return
    original_text = paragraph.text
    needs_replace = False
    for key in replacements:
        if key in original_text:
            needs_replace = True
            break
    if needs_replace:
        saved_style = get_paragraph_style(paragraph)
        new_text = original_text
        for key, value in replacements.items():
            val_str = str(value) if value is not None else ""
            new_text = new_text.replace(key, val_str)
        paragraph.clear()
        new_run = paragraph.add_run(new_text)
        apply_style_to_run(new_run, saved_style)

def replace_placeholder_with_image(doc, placeholder, image_stream):
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if placeholder in paragraph.text:
                        align = paragraph.alignment
                        paragraph.clear()
                        paragraph.alignment = align
                        run = paragraph.add_run()
                        if image_stream:
                            run.add_picture(image_stream, width=Cm(8.0))
                        return

def remove_element(element):
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)

def truncate_doc_after_page_break(doc):
    body = doc.element.body
    break_index = -1
    for i, element in enumerate(body):
        if element.tag.endswith('p'):
            if 'w:br' in element.xml and 'type="page"' in element.xml:
                break_index = i
                break
    if break_index != -1:
        for i in range(len(body) - 1, break_index - 1, -1):
            if body[i].tag.endswith('sectPr'):
                continue
            remove_element(body[i])

def generate_single_page(template_bytes, context, photo_batch, start_no):
    doc = Document(io.BytesIO(template_bytes))
    text_replacements = {f"{{{k}}}": v for k, v in context.items()}
    replace_text_content(doc, text_replacements)
    
    for i in range(1, 9):
        img_key = f"{{img_{i}}}"
        info_key = f"{{info_{i}}}"
        idx = i - 1
        if idx < len(photo_batch):
            data = photo_batch[idx]
            replace_placeholder_with_image(doc, img_key, compress_image(data['file']))
            
            spacer = "\u3000" * 4 
            
            info_text = f"照片編號：{data['no']:02d}{spacer}日期：{data['date_str']}\n"
            info_text += f"說明：{data['desc']}\n"
            
            if data.get('design') and data['design'].strip():
                info_text += f"設計：{data['design']}\n"
                
            info_text += f"實測：{data['result']}"
            
            replace_text_content(doc, {info_key: info_text})
        else:
            pass 

    if len(photo_batch) <= 4:
        truncate_doc_after_page_break(doc)
    
    final_clean = {}
    for i in range(1, 9):
        final_clean[f"{{img_{i}}}"] = ""
        final_clean[f"{{info_{i}}}"] = ""
    replace_text_content(doc, final_clean)

    return doc

def generate_names(selected_type, base_date):
    clean_type = selected_type.split(' (EA')[0].split(' (EB')[0]
    suffix = "自主檢查"
    if "施工" in clean_type or "混凝土" in clean_type:
        suffix = "施工自主檢查"
        clean_type = clean_type.replace("-施工", "")
    elif "材料" in clean_type:
        suffix = "材料進場自主檢查"
        clean_type = clean_type.replace("-材料", "")
    elif "有價廢料" in clean_type:
        suffix = "有價廢料清運自主檢查"
        clean_type = clean_type.replace("-有價廢料", "")
    
    match = re.search(r'(\(.*\))', clean_type)
    extra_info = ""
    if match:
        extra_info = match.group(1) 
        clean_type = clean_type.replace(extra_info, "").strip() 
        
    full_item_name = f"{clean_type}{suffix}{extra_info}"
    
    roc_year = base_date.year - 1911
    roc_date_str = f"{roc_year}{base_date.month:02d}{base_date.day:02d}"
    file_name = f"{roc_date_str}{full_item_name}"
    return full_item_name, file_name

def generate_clean_filename_base(selected_type, base_date):
    _, file_name = generate_names(selected_type, base_date)
    return file_name

def send_email_via_secrets(doc_bytes, filename, receiver_email, receiver_name):
    try:
        sender_email = st.secrets["email"]["account"]
        sender_password = st.secrets["email"]["password"]
    except KeyError:
        return False, "❌ 找不到 Secrets 設定！請檢查 secrets.toml。"

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = f"[自動回報] {filename.replace('.docx', '')}"
    
    body = f"""收件人：{receiver_name}\n\n這是由系統自動生成的檢查表彙整：{filename}\n內含所有檢查項目。\n\n(由 Streamlit 雲端系統自動發送)"""
    msg.attach(MIMEText(body, 'plain'))
    part = MIMEApplication(doc_bytes, Name=filename)
    part['Content-Disposition'] = f'attachment; filename="{filename}"'
    msg.attach(part)
    
    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()
        return True, f"✅ 寄送成功！已寄給 {receiver_name} ({receiver_email})"
    except Exception as e:
        return False, f"❌ 寄送失敗: {str(e)}"

def fetch_google_sheets_db(csv_url):
    try:
        df = pd.read_csv(csv_url)
        df = df.fillna("")
        
        required_cols = ["分類", "說明", "設計", "實測"]
        for col in required_cols:
            if col not in df.columns:
                return False, f"表單缺少必填欄位：{col}"
        
        new_db = {}
        current_category = "未分類項目"
        
        for _, row in df.iterrows():
            cat_val = str(row["分類"]).strip()
            if cat_val:
                current_category = cat_val
                
            desc = str(row["說明"]).strip()
            design = str(row["設計"]).strip()
            result = str(row["實測"]).strip()
            
            if not desc:
                continue 
            
            if current_category not in new_db:
                new_db[current_category] = []
                
            new_db[current_category].append({
                "desc": desc,
                "design": design,
                "result": result
            })
            
        return True, new_db
    except Exception as e:
        return False, f"讀取失敗：{str(e)}"

# --- 狀態管理函式 ---
def make_thumb(file_obj, max_side=480):
    """手機原圖動輒 3~5MB，每次重新整理都顯示原圖會很慢，改顯示小圖"""
    try:
        from PIL import ImageOps
        img = Image.open(io.BytesIO(file_obj.getvalue()))
        img = ImageOps.exif_transpose(img)
        if img.mode != "RGB": img = img.convert("RGB")
        img.thumbnail((max_side, max_side))
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=70)
        return buf.getvalue()
    except Exception:
        return None

def init_group_photos(g_idx):
    if f"photos_{g_idx}" not in st.session_state:
        st.session_state[f"photos_{g_idx}"] = []

def get_photo_time(file_obj):
    """
    取得拍照時間，回傳 (時間, 同一天內的序號, 來源說明)
    1. 照片 EXIF 拍攝時間（手機原圖、相機）
    2. 檔名裡的時間，例如 IMG_20261014_153012.jpg、PXL_20261014_073012345.jpg
    3. LINE 相簿下載的檔名 ..._261014_11.jpg（日期＋相簿順序；LINE 會刪掉 EXIF）
    都沒有就回傳 (None, 0, "")，維持上傳順序
    """
    try:
        img = Image.open(io.BytesIO(file_obj.getvalue()))
        exif = img.getexif()
        raw = exif.get_ifd(0x8769).get(36867) or exif.get(306)  # DateTimeOriginal / DateTime
        if raw:
            return datetime.datetime.strptime(str(raw).strip()[:19], "%Y:%m:%d %H:%M:%S"), 0, "EXIF"
    except Exception:
        pass
    name = os.path.basename(getattr(file_obj, "name", "") or "")
    m = re.search(r"(20\d{2})(\d{2})(\d{2})[_\-T ]?(\d{2})(\d{2})(\d{2})", name)
    if m:
        try:
            return datetime.datetime(*map(int, m.groups())), 0, "檔名"
        except ValueError:
            pass
    m = re.search(r"_(\d{2})(\d{2})(\d{2})_(\d+)\.\w+$", name)
    if m:
        try:
            yy, mm, dd, seq = map(int, m.groups())
            return datetime.datetime(2000 + yy, mm, dd), seq, "LINE"
        except ValueError:
            pass
    return None, 0, ""

def _time_sort_key(p):
    ts = p.get('ts')
    return (ts is None, ts or datetime.datetime.max, p.get('seq', 0), p.get('up_idx', 0))

def sort_photos_by_time(g_idx):
    lst = st.session_state.get(f"photos_{g_idx}", [])
    lst.sort(key=_time_sort_key)
    st.session_state[f"manual_order_{g_idx}"] = False

def add_new_photos(g_idx, uploaded_files):
    init_group_photos(g_idx)
    current_list = st.session_state[f"photos_{g_idx}"]
    existing_ids = {p['id'] for p in current_list}
    st.session_state['_up_counter'] = st.session_state.get('_up_counter', 0)

    new_items = []
    for f in uploaded_files:
        file_id = f"{f.name}_{f.size}"
        if file_id not in existing_ids:
            ts, seq, src = get_photo_time(f)
            st.session_state['_up_counter'] += 1
            new_items.append({
                "id": file_id, "file": f, "thumb": make_thumb(f),
                "ts": ts, "seq": seq, "ts_src": src, "up_idx": st.session_state['_up_counter'],
                "desc": "", "design": "", "result": "", "selected_opt_index": 0
            })
            existing_ids.add(file_id)

    if st.session_state.get(f"manual_order_{g_idx}"):
        # 已經手動排過：不打亂原本的順序，新照片自己排好接在後面
        current_list.extend(sorted(new_items, key=_time_sort_key))
    else:
        current_list.extend(new_items)
        current_list.sort(key=_time_sort_key)

def move_photo(g_idx, index, direction):
    st.session_state[f"manual_order_{g_idx}"] = True
    lst = st.session_state[f"photos_{g_idx}"]
    new_index = index + direction
    if 0 <= new_index < len(lst):
        lst[index], lst[new_index] = lst[new_index], lst[index]

def delete_photo(g_idx, index):
    lst = st.session_state[f"photos_{g_idx}"]
    if 0 <= index < len(lst):
        del lst[index]

# ==========================================
# 2. 備用資料庫與常數設定
# ==========================================

RECIPIENTS = {
    "范嘉文": "ses543212004@fengyu.com.tw",
    "林憲睿": "dennys871022@fengyu.com.tw",
    "翁育玟": "Vicky1019@fengyu.com.tw",
    "林智捷": "ccl20010218@fengyu.com.tw",
    "趙健鈞": "kk919472770@fengyu.com.tw",
    "孫永明": "kevin891023@fengyu.com.tw",
    "林泓鈺": "henry30817@fengyu.com.tw",
    "黃元杰": "s10411097@fengyu.com.tw",
    "郭登慶": "tw850502@fengyu.com.tw",
    "歐冠廷": "canon1220@fengyu.com.tw",
    "黃彥榤": "ajh73684@fengyu.com.tw",
    "陳昱勳": "x85082399@fengyu.com.tw",
    "測試用 (寄給自己)": st.secrets["email"]["account"] if "email" in st.secrets else "test@example.com"
}

COMMON_SUB_CONTRACTORS = [
    "川峻工程有限公司",
    "世詮營造股份有限公司",
    "互國企業有限公司",
    "世和金屬股份有限公司",
    "宥辰興業股份有限公司",
    "亞東預拌混凝土股份有限公司",
    "自行輸入..." 
]

DEFAULT_CHECKS_DB = {
    "預設資料 (雲端連結失敗時顯示)": [
        {"desc": "這是一個預設項目", "design": "設定範例", "result": "實測範例"}
    ]
}

# ==========================================
# 3. 主程式介面邏輯
# ==========================================

st.set_page_config(page_title="工程自主檢查表生成器", layout="wide")
st.title("🏗️ 工程自主檢查表 (主控同步雲端版)")

def load_latest_db():
    if GOOGLE_SHEETS_CSV_URL.strip():
        success, result = fetch_google_sheets_db(GOOGLE_SHEETS_CSV_URL.strip())
        if success:
            return result
        else:
            st.error(f"雲端資料庫載入失敗：{result} (退回預設資料)")
            return DEFAULT_CHECKS_DB
    return DEFAULT_CHECKS_DB

if 'checks_db' not in st.session_state:
    st.session_state['checks_db'] = load_latest_db()

# Init Variables
if 'merged_doc_buffer' not in st.session_state: st.session_state['merged_doc_buffer'] = None
if 'merged_filename' not in st.session_state: st.session_state['merged_filename'] = ""
if 'saved_template' not in st.session_state: st.session_state['saved_template'] = None
if 'num_groups' not in st.session_state: st.session_state['num_groups'] = 1

DEFAULT_TEMPLATE_PATH = "template.docx"
if not st.session_state['saved_template'] and os.path.exists(DEFAULT_TEMPLATE_PATH):
    with open(DEFAULT_TEMPLATE_PATH, "rb") as f:
        st.session_state['saved_template'] = f.read()

# ==========================================
# ★ 連動邏輯：文字修改後自動推播給其他組 (加入 3 大格空格)
# ==========================================
def on_item_0_change():
    if "item_0" in st.session_state:
        base_name = st.session_state["item_0"]
        
        # 智慧過濾：移除結尾的空格與 #數字，確保我們拿到最純淨的主檔名
        base_name = re.sub(r'(\u3000|\s)*#\d+$', '', base_name)
            
        num = st.session_state.get('num_groups', 1)
        spacer = "\u3000" * 3  # 3個全形大空格
        
        for other_g in range(1, num):
            # 自動推播並加上 3個空格與自己的編號
            st.session_state[f"item_{other_g}"] = f"{base_name}{spacer}#{other_g + 1}"

def update_group_info(g_idx):
    base_date = st.session_state.get('global_date', datetime.date.today())
    selected_type = st.session_state[f"type_{g_idx}"]
    item_name, _ = generate_names(selected_type, base_date)
    
    spacer = "\u3000" * 3  # 3個全形大空格
    
    # ★ 自動加上 3個大空格 與 #1, #2 等編號
    st.session_state[f"item_{g_idx}"] = f"{item_name}{spacer}#{g_idx + 1}"
    
    def clear_group_data(idx):
        keys_to_clear = [k for k in st.session_state.keys() if f"_{idx}_" in k and (k.startswith("sel_") or k.startswith("desc_") or k.startswith("design_") or k.startswith("result_"))]
        for k in keys_to_clear: del st.session_state[k]
        if f"photos_{idx}" in st.session_state:
            for p in st.session_state[f"photos_{idx}"]:
                p['desc'] = ""; p['design'] = ""; p['result'] = ""; p['selected_opt_index'] = 0

    clear_group_data(g_idx)
    
    if g_idx == 0:
        current_num_groups = st.session_state.get('num_groups', 1)
        for other_g in range(1, current_num_groups):
            st.session_state[f"type_{other_g}"] = selected_type
            # 確保同步時，其他組別擁有正確的 空格+#2, #3 編號
            st.session_state[f"item_{other_g}"] = f"{item_name}{spacer}#{other_g + 1}"
            clear_group_data(other_g)

# ==========================================
# ★ AI 白板辨識：讀白板 → 對應工項/說明 → 自動填入
# ==========================================
AI_CFG = get_ai_config(st.secrets) if hasattr(st, "secrets") else None

def _sync_photo_fields_from_widgets(g_idx):
    """把畫面上已手動改過的文字先存回 photo dict，避免被覆蓋時遺失"""
    for p in st.session_state.get(f"photos_{g_idx}", []):
        for field in ("desc", "design", "result"):
            k = f"{field}_{g_idx}_{p['id']}"
            if k in st.session_state: p[field] = st.session_state[k]

def run_ai_on_group(g_idx, only_index=None):
    if not AI_CFG: return
    photos = st.session_state.get(f"photos_{g_idx}", [])
    if not photos: return
    _sync_photo_fields_from_widgets(g_idx)

    targets = list(range(len(photos))) if only_index is None else [only_index]
    db = st.session_state['checks_db']
    current_type = st.session_state.get(f"type_{g_idx}")
    msgs = []

    if only_index is None:
        with st.spinner(f"🤖 AI 正在從 {len(targets)} 張照片中找出白板並辨識..."):
            results = recognize_many(AI_CFG, [photos[i]['file'] for i in targets], db, current_type)
    else:
        with st.spinner("🤖 AI 正在辨識這張白板..."):
            results = [recognize_one(AI_CFG, photos[only_index]['file'], db, current_type)]

    valid = [(i, r) for i, r in zip(targets, results) if r.get("ok") and r.get("has_board")]

    # 1) 決定工項分類：整組辨識時用多數決；單張辨識不動分類
    final_type = current_type
    if only_index is None:
        votes = {}
        for _, r in valid:
            if r.get("category"): votes[r["category"]] = votes.get(r["category"], 0) + 1
        if votes:
            best = max(votes, key=votes.get)
            if best != current_type:
                final_type = best
                st.session_state[f"type_{g_idx}"] = best
                base_date = st.session_state.get('global_date', get_taiwan_date())
                item_name, _ = generate_names(best, base_date)
                spacer = "\u3000" * 3
                st.session_state[f"item_{g_idx}"] = f"{item_name}{spacer}#{g_idx + 1}"
                msgs.append(("info", f"工項已自動切換為：{best}"))

    # 2) 每張照片填入 說明/設計/實測
    errors = []
    for i, r in zip(targets, results):
        p = photos[i]
        if not r.get("ok"):
            errors.append((i, r.get("error", "未知錯誤")))
            continue
        if not r.get("has_board"):
            continue  # 一般檢查照片：完全不動，保留原本手動填的內容
        for field in ("desc", "design", "result"):
            val = str(r.get(field) or "").strip()
            p[field] = val
            st.session_state[f"{field}_{g_idx}_{p['id']}"] = val
        if r.get("category") == final_type and r.get("item_index") is not None:
            p['selected_opt_index'] = r["item_index"] + 1
        else:
            p['selected_opt_index'] = 0
        st.session_state.pop(f"sel_{g_idx}_{p['id']}", None)  # 讓下拉選單依新的 index 重建
        p['ai'] = {
            "vals": {f: p[f] for f in ("desc", "design", "result")},
            "opt": p['selected_opt_index'],
            "confidence": r.get("confidence", 0),
            "notes": r.get("notes", ""),
            "board": r.get("board", {}),
            "off_category": bool(r.get("category")) and r.get("category") != final_type,
        }

    # 3) 位置／日期帶入側邊欄（取第一張有讀到的）
    if st.session_state.get("ai_fill_meta", True):
        for _, r in valid:
            b = r.get("board") or {}
            loc = str(b.get("location") or "").strip()
            d = parse_board_date(b.get("date"))
            if loc: st.session_state["p_loc"] = loc
            if d: st.session_state["global_date"] = d
            if loc or d:
                msgs.append(("info", f"已帶入 位置：{loc or '(未讀到)'}　日期：{d or '(未讀到)'}"))
                break

    for err in dict.fromkeys(e for _, e in errors):  # 同一批失敗只顯示一次
        nos = "、".join(f"No.{i+1:02d}" for i, e in errors if e == err)
        msgs.append(("error", f"{nos} 辨識失敗：{err}"))

    board_nos = "、".join(f"No.{i+1:02d}" for i, _ in valid)
    if only_index is not None:
        if valid: msgs.insert(0, ("success", "✅ 已重新辨識這張白板，請確認內容"))
        elif not errors: msgs.insert(0, ("warning", "這張照片讀不到白板內容，請手動填寫"))
    elif valid:
        other = len(targets) - len(valid) - len(errors)
        msgs.insert(0, ("success", f"✅ 找到 {len(valid)} 張白板（{board_nos}）已自動填入；"
                                   f"其他 {other} 張檢查照片未更動，請用下拉選單填寫"))
    elif not errors:
        msgs.insert(0, ("warning", "這組照片裡沒有找到白板。若某張其實是白板，可按該張旁邊的 🤖 強制辨識"))
    st.session_state[f"ai_msg_{g_idx}"] = msgs

def restore_ai_result(g_idx, i):
    p = st.session_state[f"photos_{g_idx}"][i]
    ai = p.get('ai') or {}
    if not ai.get('vals'): return
    for f, v in ai['vals'].items():
        p[f] = v
        st.session_state[f"{f}_{g_idx}_{p['id']}"] = v
    p['selected_opt_index'] = ai.get('opt', 0)
    st.session_state.pop(f"sel_{g_idx}_{p['id']}", None)

def clear_all_data():
    for key in list(st.session_state.keys()):
        if key.startswith(('type_', 'item_', 'fname_', 'photos_', 'file_', 'sel_', 'desc_', 'design_', 'result_', 'ai_msg_')):
            del st.session_state[key]
    st.session_state['num_groups'] = 1
    st.session_state['merged_doc_buffer'] = None
    st.session_state['merged_filename'] = ""

# ==========================================
# ★ 記住常用設定：存在「各自手機/電腦的瀏覽器」，不同人互不影響
# ==========================================
PREFS_COOKIE = "sia_prefs"
PREF_DEFAULTS = {
    "p_name": "衛生福利部防疫中心興建工程",
    "p_cont": "豐譽營造股份有限公司",
    "sub_select": COMMON_SUB_CONTRACTORS[0],
    "p_sub_custom": "川峻工程有限公司",
    "recipient": list(RECIPIENTS.keys())[0],
    "ai_fill_meta": True,
}

def load_prefs_once():
    if st.session_state.get('_prefs_loaded'): return
    st.session_state['_prefs_loaded'] = True
    saved = {}
    try:
        raw = st.context.cookies.get(PREFS_COOKIE)
        if raw: saved = json.loads(unquote(raw))
    except Exception:
        saved = {}
    effective = {}
    for k, default in PREF_DEFAULTS.items():
        v = saved.get(k, default)
        if k == "sub_select" and v not in COMMON_SUB_CONTRACTORS: v = default
        if k == "recipient" and v not in RECIPIENTS: v = default
        effective[k] = v
        if k not in st.session_state: st.session_state[k] = v
    st.session_state['_prefs_saved'] = effective

def ensure_pref(k):
    """有些欄位不一定每次都顯示（例如「自行輸入」廠商），被 Streamlit 清掉時從已存設定補回"""
    if k not in st.session_state:
        st.session_state[k] = st.session_state.get('_prefs_saved', {}).get(k, PREF_DEFAULTS[k])

def save_prefs_if_changed():
    prev = st.session_state.get('_prefs_saved', {})
    current = {k: st.session_state.get(k, prev.get(k)) for k in PREF_DEFAULTS}
    if current == st.session_state.get('_prefs_saved'): return
    st.session_state['_prefs_saved'] = current
    value = json.dumps(json.dumps(current, ensure_ascii=False))
    components.html(f"""<script>
      (function() {{
        var c = "{PREFS_COOKIE}=" + encodeURIComponent({value}) + "; max-age=31536000; path=/; SameSite=Lax";
        try {{ window.parent.document.cookie = c; }} catch (e) {{ document.cookie = c; }}
      }})();
    </script>""", height=0)

load_prefs_once()

# Sidebar
with st.sidebar:
    st.header("1. 樣板設定")
    if st.session_state['saved_template']:
        st.success("✅ Word 樣板已載入")
    else:
        uploaded = st.file_uploader("上傳樣板", type=['docx'])
        if uploaded:
            st.session_state['saved_template'] = uploaded.getvalue()
            st.rerun()
            
    st.markdown("---")
    st.header("☁️ 雲端資料庫狀態")
    if GOOGLE_SHEETS_CSV_URL.strip():
        st.success("✅ 已綁定專屬試算表")
        if st.button("🔄 點我強制同步最新資料", use_container_width=True, type="primary"):
            with st.spinner("📥 正在抓取最新資料..."):
                st.session_state['checks_db'] = load_latest_db()
                st.success("更新完成！")
                st.rerun()
    else:
        st.warning("⚠️ 尚未設定 GOOGLE_SHEETS_CSV_URL。")
            
    st.markdown("---")
    st.button("🗑️ 清除所有填寫資料", on_click=clear_all_data, use_container_width=True)

    st.markdown("---")
    st.header("🤖 AI 白板辨識")
    if AI_CFG:
        st.success(f"✅ 已啟用（{AI_CFG['provider']} / {AI_CFG['model']}）")
        ensure_pref("ai_fill_meta")
        st.checkbox("辨識時一併帶入「施作位置」與「日期」", key="ai_fill_meta")
    else:
        st.info("尚未設定 API Key，請在 Secrets 加入 [ai] 區段")

    st.markdown("---")
    st.header("2. 專案資訊")
    st.caption("💾 工程名稱、廠商、收件人會記在這台裝置的瀏覽器，每個人各自記錄")
    p_name = st.text_input("工程名稱", key="p_name")
    p_cont = st.text_input("施工廠商", key="p_cont")
    sub_select = st.selectbox("協力廠商", COMMON_SUB_CONTRACTORS, key="sub_select")
    if sub_select == "自行輸入...":
        ensure_pref("p_sub_custom")
        p_sub = st.text_input("請輸入廠商名稱", key="p_sub_custom")
    else:
        p_sub = sub_select
    if "p_loc" not in st.session_state: st.session_state["p_loc"] = "北棟 1F"
    if "global_date" not in st.session_state: st.session_state["global_date"] = get_taiwan_date()
    p_loc = st.text_input("施作位置", key="p_loc")
    base_date = st.date_input("日期", key='global_date')

# Main Body
if st.session_state['saved_template']:
    num_groups = st.number_input("本次產生幾組檢查表？", min_value=1, value=st.session_state['num_groups'], key='num_groups_input')
    st.session_state['num_groups'] = num_groups
    all_groups_data = []

    for g in range(num_groups):
        st.markdown(f"---")
        st.subheader(f"📂 第 {g+1} 組")
        c1, c2, c3 = st.columns([2, 2, 1])
        db_options = list(st.session_state['checks_db'].keys())
        
        # ==========================================
        # ★ 剛新增組別時，自動預設帶入第一組的選項及名稱 (加入大空格)
        # ==========================================
        if g > 0 and f"type_{g}" not in st.session_state and "type_0" in st.session_state:
            st.session_state[f"type_{g}"] = st.session_state["type_0"]
            if "item_0" in st.session_state:
                base_name = st.session_state["item_0"]
                base_name = re.sub(r'(\u3000|\s)*#\d+$', '', base_name)
                spacer = "\u3000" * 3
                st.session_state[f"item_{g}"] = f"{base_name}{spacer}#{g + 1}"
            
        selected_type = c1.selectbox(f"選擇檢查工項", db_options, key=f"type_{g}", on_change=update_group_info, args=(g,))
        
        if f"item_{g}" not in st.session_state:
            update_group_info(g)
            
        # ★ 第一組綁定 on_change，只要修改就會同步全場
        if g == 0:
            g_item = c2.text_input(f"自檢項目名稱", key=f"item_{g}", on_change=on_item_0_change)
        else:
            g_item = c2.text_input(f"自檢項目名稱", key=f"item_{g}")
            
        roc_year = base_date.year - 1911
        date_display = f"{roc_year}.{base_date.month:02d}.{base_date.day:02d}"
        c3.text(f"日期: {date_display}")

        st.markdown("##### 📸 照片上傳與排序")
        uploader_key_name = f"uploader_key_{g}"
        if uploader_key_name not in st.session_state: st.session_state[uploader_key_name] = 0
        dynamic_key = f"uploader_{g}_{st.session_state[uploader_key_name]}"
        
        new_files = st.file_uploader(f"點擊此處選擇照片 (第 {g+1} 組)", type=['jpg','png','jpeg'], accept_multiple_files=True, key=dynamic_key)
        if new_files:
            add_new_photos(g, new_files)
            st.session_state[uploader_key_name] += 1
            st.rerun()
        
        if st.session_state.get(f"photos_{g}"):
            st.button("🕒 依拍照時間重新排序", key=f"tsort_{g}", on_click=sort_photos_by_time, args=(g,))
            if st.button("🔄 順序反了嗎？點我「一鍵反轉」照片順序", key=f"rev_{g}"):
                current_list = st.session_state[f"photos_{g}"]
                for p in current_list:
                    d_key = f"desc_{g}_{p['id']}"
                    if d_key in st.session_state: p['desc'] = st.session_state[d_key]
                    des_key = f"design_{g}_{p['id']}"
                    if des_key in st.session_state: p['design'] = st.session_state[des_key]
                    r_key = f"result_{g}_{p['id']}"
                    if r_key in st.session_state: p['result'] = st.session_state[r_key]
                    s_key = f"sel_{g}_{p['id']}"
                    if s_key in st.session_state: p['selected_opt_index'] = st.session_state[s_key]
                st.session_state[f"photos_{g}"].reverse()
                st.session_state[f"manual_order_{g}"] = True
                st.rerun()
        
        init_group_photos(g)
        photo_list = st.session_state[f"photos_{g}"]

        if photo_list and AI_CFG:
            st.button(f"🤖 AI 找出白板並自動填入（本組 {len(photo_list)} 張只算 1 次辨識）" if len(photo_list) <= 8
                      else f"🤖 AI 找出白板並自動填入（本組 {len(photo_list)} 張）", key=f"ai_all_{g}",
                      on_click=run_ai_on_group, args=(g,), type="primary", use_container_width=True)
        for level, text in st.session_state.pop(f"ai_msg_{g}", []):
            getattr(st, level)(text)
        
        if photo_list:
            check_items_list = st.session_state['checks_db'].get(selected_type, [])
            options = ["(請選擇...)"] + [item['desc'] for item in check_items_list]

            for i, photo_data in enumerate(photo_list):
                with st.container():
                    col_img, col_info, col_ctrl = st.columns([1.5, 3, 0.5])
                    pid = photo_data['id']
                    with col_img:
                        if 'thumb' not in photo_data: photo_data['thumb'] = make_thumb(photo_data['file'])
                        st.image(photo_data['thumb'] or photo_data['file'], use_container_width=True)
                        ts = photo_data.get('ts')
                        if ts and photo_data.get('ts_src') == "LINE":
                            ts_txt = f"　🕒 {ts:%m/%d} 第{photo_data.get('seq', 0)}張"
                        elif ts:
                            ts_txt = f"　🕒 {ts:%m/%d %H:%M}"
                        else:
                            ts_txt = "　🕒 無時間資訊"
                        st.caption(f"No. {i+1:02d}{ts_txt}")
                        ai = photo_data.get('ai')
                        if ai:
                            if ai.get('error'):
                                st.warning(f"🤖 {ai['error']}")
                            else:
                                conf = float(ai.get('confidence') or 0)
                                tag = "🟢" if conf >= 0.8 else ("🟡" if conf >= 0.5 else "🔴")
                                st.caption(f"{tag} AI 信心 {conf:.0%}" + (f"｜{ai['notes']}" if ai.get('notes') else ""))
                                if ai.get('off_category'):
                                    st.caption("⚠️ 這張白板的工項跟本組不同，可能要移到別組")
                                b = ai.get('board') or {}
                                labels = [("location", "位置"), ("date", "日期"), ("inspection_item", "項目"),
                                          ("seq", "編號"), ("design_values", "設計"), ("measured_values", "實測")]
                                lines = [f"{lb}：{b.get(k)}" for k, lb in labels if b.get(k)]
                                if lines:
                                    with st.expander("🔍 AI 讀到的白板"):
                                        st.text("\n".join(lines))
                                st.button("↩️ 還原 AI 結果", key=f"ai_undo_{g}_{i}",
                                          on_click=restore_ai_result, args=(g, i))
                    with col_info:
                        def on_select_change(pk=pid, gk=g):
                            k = f"sel_{gk}_{pk}"
                            if k not in st.session_state: return
                            new_idx = st.session_state[k]
                            dk, desk, rk = f"desc_{gk}_{pk}", f"design_{gk}_{pk}", f"result_{gk}_{pk}"
                            if isinstance(new_idx, int) and new_idx > 0 and new_idx <= len(check_items_list):
                                item_data = check_items_list[new_idx-1]
                                st.session_state[dk] = item_data['desc']
                                st.session_state[desk] = item_data['design']
                                st.session_state[rk] = item_data['result']
                            else:
                                st.session_state[dk] = ""
                                st.session_state[desk] = ""
                                st.session_state[rk] = ""

                        current_opt_idx = photo_data.get('selected_opt_index', 0)
                        if current_opt_idx > len(options): current_opt_idx = 0
                        st.selectbox("快速填寫", range(len(options)), format_func=lambda x: options[x], index=current_opt_idx, key=f"sel_{g}_{pid}", on_change=on_select_change, label_visibility="collapsed")

                        def on_text_change(field, pk=pid, idx=i, gk=g): 
                            val = st.session_state[f"{field}_{gk}_{pk}"]
                            st.session_state[f"photos_{gk}"][idx][field] = val
                            if field == 'sel': st.session_state[f"photos_{gk}"][idx]['selected_opt_index'] = val

                        desc_key = f"desc_{g}_{pid}"
                        design_key = f"design_{g}_{pid}"
                        result_key = f"result_{g}_{pid}"
                        if desc_key not in st.session_state: st.session_state[desc_key] = photo_data.get('desc', '')
                        if design_key not in st.session_state: st.session_state[design_key] = photo_data.get('design', '')
                        if result_key not in st.session_state: st.session_state[result_key] = photo_data.get('result', '')

                        st.text_input("說明", key=desc_key, on_change=on_text_change, args=('desc',))
                        st.text_input("設計 (可留空)", key=design_key, on_change=on_text_change, args=('design',))
                        st.text_input("實測", key=result_key, on_change=on_text_change, args=('result',))

                    with col_ctrl:
                        if st.button("⬆️", key=f"up_{g}_{i}"): move_photo(g, i, -1); st.rerun()
                        if st.button("⬇️", key=f"down_{g}_{i}"): move_photo(g, i, 1); st.rerun()
                        if st.button("❌", key=f"del_{g}_{i}"): delete_photo(g, i); st.rerun()
                        if AI_CFG:
                            st.button("🤖", key=f"ai_one_{g}_{i}", help="把這張當白板辨識（AI 沒認出白板、或讀錯時用）",
                                      on_click=run_ai_on_group, args=(g, i))
                    st.divider()

            g_photos_export = []
            for i, p in enumerate(photo_list):
                d_val = st.session_state.get(f"desc_{g}_{p['id']}", p['desc'])
                des_val = st.session_state.get(f"design_{g}_{p['id']}", p['design'])
                r_val = st.session_state.get(f"result_{g}_{p['id']}", p['result'])
                g_photos_export.append({
                    "file": p['file'], "no": i + 1, "date_str": date_display, 
                    "desc": d_val, "design": des_val, "result": r_val
                })

            all_groups_data.append({
                "group_id": g+1,
                "context": {
                    "project_name": p_name, "contractor": p_cont, "sub_contractor": p_sub,
                    "location": p_loc, "date": date_display, "check_item": g_item
                },
                "photos": g_photos_export
            })

    st.markdown("---")
    st.subheader("🚀 執行操作")
    default_filename = ""
    if "type_0" in st.session_state:
        default_filename = generate_clean_filename_base(st.session_state["type_0"], base_date)
    else:
        default_filename = f"自主檢查表_{get_taiwan_date()}"

    final_file_name_input = st.text_input("📝 最終 Word 檔名", value=default_filename)
    if not final_file_name_input.endswith(".docx"): final_file_name = final_file_name_input + ".docx"
    else: final_file_name = final_file_name_input

    ensure_pref("recipient")
    selected_name = st.selectbox("📬 收件人", list(RECIPIENTS.keys()), key="recipient")
    target_email = RECIPIENTS[selected_name]

    if st.button("步驟 1：生成報告資料 (單一 Word 檔)", type="primary", use_container_width=True):
        if not all_groups_data: st.error("⚠️ 請至少上傳一張照片並填寫資料")
        else:
            with st.spinner("📦 正在生成並合併 Word 檔案..."):
                master_doc = None
                composer = None
                for group in all_groups_data:
                    photos = group['photos']
                    context = group['context']
                    for page_idx, i in enumerate(range(0, len(photos), 8)):
                        batch = photos[i : i+8]
                        start_no = i + 1
                        current_doc = generate_single_page(st.session_state['saved_template'], context, batch, start_no)
                        if master_doc is None:
                            master_doc = current_doc
                            composer = Composer(master_doc)
                        else:
                            composer.append(current_doc)
                out_buffer = io.BytesIO()
                composer.save(out_buffer)
                st.session_state['merged_doc_buffer'] = out_buffer.getvalue()
                st.session_state['merged_filename'] = final_file_name
                st.success(f"✅ 彙整完成！檔名：{final_file_name}")

    if st.session_state['merged_doc_buffer']:
        col_mail, col_dl = st.columns(2)
        with col_mail:
            if st.button(f"📧 立即寄出 Word 檔給：{selected_name}", use_container_width=True):
                with st.spinner("📨 雲端發信中..."):
                    success, msg = send_email_via_secrets(st.session_state['merged_doc_buffer'], st.session_state['merged_filename'], target_email, selected_name)
                    if success: st.success(msg)
                    else: st.error(msg)
        with col_dl:
            st.download_button(label="📥 下載 Word 檔案", data=st.session_state['merged_doc_buffer'], file_name=st.session_state['merged_filename'], mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
else:
    st.info("👈 請先在左側確認 Word 樣板")

save_prefs_if_changed()
