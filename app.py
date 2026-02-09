import streamlit as st
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml.ns import qn
from docxcompose.composer import Composer # <--- 這是合併檔案的關鍵
from PIL import Image
import io
import datetime
from datetime import timedelta, timezone
import os
import zipfile
import pandas as pd
import smtplib
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# --- 0. 台灣時區設定 ---
def get_taiwan_date():
    utc_now = datetime.datetime.now(timezone.utc)
    return (utc_now + timedelta(hours=8)).date()

# --- 1. 設定收件人名單 ---
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

# --- 常用協力廠商名單 ---
COMMON_SUB_CONTRACTORS = [
    "川峻工程有限公司",
    "世銓營造股份有限公司",
    "互國企業有限公司",
    "世和金屬股份有限公司",
    "宥辰興業股份有限公司",
    "亞東預拌混凝土股份有限公司",
    "自行輸入..." 
]

# --- 2. 終極內建資料庫 ---
CHECKS_DB = {
    "拆除工程-施工 (EA26)": {
        "items": [
            "防護措施:公共管線及環境保護", "安全監測:初始值測量", "防塵作為:灑水或防塵網",
            "降噪作為:低噪音機具", "構造物拆除順序:由上而下", "保留構件:記號保護",
            "拆除物分類:回收/不可回收/有價", "車輛輪胎清潔", "安全監測數據查核",
            "地坪整平清潔", "廢棄物清運"
        ],
        "results": [
            "已完成相關防護措施，管線已封閉/遷移", "已完成初始值測量及設置", "現場已設置灑水或防塵網",
            "使用低噪音機具、非衝擊式工法", "依施工規劃由上而下拆除", "保留構件已標示並保護",
            "已依類別分類置放", "輪胎已清潔，無帶污泥出場", "傾斜計<1/937.5，沉陷點<2cm",
            "地坪已平整清潔", "依核定計畫書執行清運"
        ]
    },
    "拆除工程-有價廢料 (EB26)": {
        "items": [
            "廢鋼筋載運", "銅線/製品載運", "電線電纜(含皮)載運", "型鋼載運", 
            "鋁料載運", "載運車輛資訊", "重量查核(空車重)", "重量查核(總重)", "重量查核(有價物重)"
        ],
        "results": [
            "載運廢鋼筋 * 1 車", "載運銅製品 * 1 車", "載運電纜 * 1 車", "載運型鋼 * 1 車", 
            "載運鋁料 * 1 車", "車號：__________", "空車重:____kg", "總重:____kg", "有價物重:____kg"
        ]
    },
    "擋土排樁工程(排樁)-施工": {
        "items": [
            "放樣樁位檢測", "鑽掘垂直度", "鑽掘深度/入岩", "排樁直徑",
            "鋼筋籠(主筋/箍筋)", "鋼筋籠搭接/銲接", "鋼筋間隔器",
            "特密管埋置深度", "混凝土澆置(樁身)", "壓梁-鋼筋綁紮",
            "壓梁-模內尺寸", "壓梁-混凝土澆置", "壓梁-完成面高程", "澆置後清潔"
        ],
        "results": [
            "偏差 ≦3cm", "套管內≦1/300, 土內≦1/100", "設計深度≥14.5m, 入岩盤≥3m",
            "D≥80cm", "主筋#10(14支); 箍筋#4@10cm", "搭接#10=153cm; 銲接4cm",
            "@200cm", "埋置深度≥2M", "fc'=280kgf/cm2; 澆置不中斷",
            "主筋#7/#6; 箍筋#4@15cm", "60*80cm", "fc'=210kgf/cm2; 坍度20±4cm",
            "依施工圖施作 ±3cm", "表面平整、無汙染"
        ]
    },
    "擋土排樁工程(預壘樁)-施工": {
        "items": [
            "樁心檢測", "鑽掘垂直度", "預壘樁長度/直徑", "鋼筋籠(主筋/箍筋)",
            "鋼筋籠搭接/銲接", "水泥砂漿試體/壓力", "澆置間隔時間",
            "微型樁鑽掘(垂直/深度)", "微型樁注漿(水灰比)", "微型樁鋼管",
            "壓梁-鋼筋綁紮", "壓梁-模內尺寸", "壓梁-混凝土澆置", "澆置後清潔"
        ],
        "results": [
            "D40/D35: ±3cm", "≦1/100", "L≥6.3m; D=40/35cm", "主筋#8/#7; 箍筋#4@15cm",
            "搭接#8=139cm/#7=121cm; 銲接4cm", "5x5x5cm方塊; 壓力≥2.1kgf/cm2",
            "不得超過3分鐘", "10度±3度; L≥7m; 間距@45cm", "W/C=1:1; ≦10min",
            "L=7m; 間隔器@2m", "主筋#6; 箍筋#4@15cm", "D40:40x120 / D35:35x60",
            "fc'=210kgf/cm2; 坍度20±4cm", "表面平整、無汙染"
        ]
    },
    "擋土排樁工程(CCP止水樁)-施工": {
        "items": [
            "定位樁心檢測", "鑽掘垂直度", "止水樁長度", "止水樁直徑",
            "水泥漿水灰比", "注漿壓力值", "澆置後清潔"
        ],
        "results": [
            "±3cm", "≦1/40", "L≥14.5m (樁底至相鄰排樁頂)", "D≥30cm",
            "W/C=1:1", "≥180kgf/cm2", "水泥漿澆置後清潔"
        ]
    },
    "擋土排樁工程-材料": {
        "items": ["證明文件查核", "規格尺寸檢查", "外觀形狀檢查", "工地放置檢查", "取樣試驗"],
        "results": ["出廠證明/檢驗紀錄齊全", "符合契約規範及訂貨規格", "無碰撞變形、破損、裂痕", "分類置放並標幟、底部墊高", "依規範取樣/不取樣"]
    },
    "微型樁工程-施工 (EA53)": {
        "items": ["開挖前置:管線確認", "樁心檢測 (≦3cm)", "鑽掘垂直度 (0-5度)", "鑽掘尺寸 (深度/樁徑)", "鑽掘間距 (@60cm)", "水泥漿拌合比 (1:1)", "注漿作業 (≦10min)", "鋼管吊放安裝", "廢漿清除", "樁頂劣質打石", "帽梁鋼筋綁紮", "帽梁灌漿"],
        "results": ["確認無地下管線干擾", "樁心偏差 ≦3cm", "垂直度符合規定 (0-5度)", "深度≧16m; 樁徑≧15cm", "間距@60cm, 交錯施工", "水灰比 W/C=1:1", "時間≦10min，注漿至帽梁底部", "長度16m; 間隔器@2m", "已清除硬固廢漿", "劣質混凝土已打除", "主筋#6-4支, 箍筋#3@20cm", "強度 fc'=210kgf/cm2"]
    },
    "微型樁工程-材料 (EB53)": {
        "items": ["證明文件", "規格尺寸", "外觀形狀", "工地放置", "取樣試驗"],
        "results": ["出廠證明/檢驗紀錄齊全", "符合契約規範", "無碰撞變形", "分類堆置/標示", "依規範取樣"]
    },
    "假設工程-施工 (EA51)": {
        "items": ["放樣", "全阻式圍籬組立", "半阻式圍籬組立", "防溢座施作", "出入口地坪(鋼筋/澆置)", "大門安裝", "安全走廊", "警示燈設置", "洗車台尺寸檢查", "圍籬綠化維護"],
        "results": ["依施工圖說放樣", "間距/埋入深度符合規定", "間距/埋入深度符合規定", "混凝土210kgf/cm2", "厚度20cm; 雙層雙向#4@10cm", "尺寸及埋入深度符合規定", "高300寬150cm", "間距符合規定", "500x522cm; 沉沙池深170cm", "存活率90%以上"]
    },
    "假設工程-材料 (EB51)": {
        "items": ["證明文件", "外觀形狀", "工地放置", "預鑄水溝尺寸"],
        "results": ["出廠證明/檢驗紀錄齊全", "無碰撞變形、破損", "分類堆置/標示", "內溝寬30±5cm, 深40±5cm"]
    },
    "車道拓寬工程 (EA52)": {
        "items": ["碎石級配舖設", "鋼筋綁紮", "模板組立", "混凝土澆置(結構)", "粉刷面清潔", "基準灰誌製作", "馬賽克磚舖貼", "瀝青混凝土舖設"],
        "results": ["級配高度 20cm", "箍筋#4@20cm; 保護層4cm", "牆厚20cm; 垂直度±13mm", "強度 210kgf/cm2", "無殘餘雜物、凸出物", "間距不大於1M", "顏色與樣板相同", "密級配，無汙損浮起"]
    },
    "混凝土工程 (共用)": {
        "items": ["照明與雨天防護", "澆置前清潔濕潤", "模板振動器", "澆置時間控制", "坍度/流度檢查", "溫度檢查", "氯離子含量", "試體取樣", "振動搗實", "養護作業"],
        "results": ["照明充足，備有防雨材", "垃圾清除，模板濕潤", "備有至少二具", "拌合至澆置90分鐘內", "符合設計 (如 18±4cm)", "13~32度C", "小於 0.15 kg/m3", "每100m3取樣1組", "間距<50cm; 每次5-10秒", "灑水或覆蓋養護"]
    }
}

# --- 3. 核心功能 ---

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

def cleanup_template_for_short_report(doc, num_photos):
    if num_photos > 4:
        return 
    
    placeholders_to_remove = [f"{{img_{i}}}" for i in range(5, 9)] + \
                             [f"{{info_{i}}}" for i in range(5, 9)]
    
    for table in list(doc.tables): 
        for row in list(table.rows):
            row_text = ""
            for cell in row.cells:
                row_text += cell.text
            if any(ph in row_text for ph in placeholders_to_remove):
                remove_element(row._element)
                
    for paragraph in list(doc.paragraphs):
        if any(ph in paragraph.text for ph in placeholders_to_remove):
            remove_element(paragraph._element)
            
    for p in doc.paragraphs:
        if p.runs:
            for r in p.runs:
                if 'w:br' in r._element.xml and 'type="page"' in r._element.xml:
                    remove_element(r._element)

def generate_single_page(template_bytes, context, photo_batch, start_no):
    doc = Document(io.BytesIO(template_bytes))
    
    # 1. 文字替換
    text_replacements = {f"{{{k}}}": v for k, v in context.items()}
    replace_text_content(doc, text_replacements)
    
    # 2. 填入照片
    for i in range(1, 9):
        img_key = f"{{img_{i}}}"
        info_key = f"{{info_{i}}}"
        idx = i - 1
        if idx < len(photo_batch):
            data = photo_batch[idx]
            replace_placeholder_with_image(doc, img_key, compress_image(data['file']))
            spacer = "\u3000" * 6 
            info_text = f"照片編號：{data['no']:02d}{spacer}日期：{data['date_str']}\n"
            info_text += f"說明：{data['desc']}\n"
            info_text += f"實測：{data['result']}"
            replace_text_content(doc, {info_key: info_text})
        else:
            pass 

    # 3. 智慧縮減 (刪除多餘頁面)
    cleanup_template_for_short_report(doc, len(photo_batch))
    
    # 4. 清理剩餘佔位符
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

# --- Email 寄送功能 (更新為傳送單一 .docx) ---
def send_email_via_secrets(doc_bytes, filename, receiver_email, receiver_name):
    try:
        sender_email = st.secrets["email"]["account"]
        sender_password = st.secrets["email"]["password"]
    except KeyError:
        return False, "❌ 找不到 Secrets 設定！請檢查 secrets.toml。"

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = f"[自動回報] {filename.replace('.docx', '')}" # 標題去掉副檔名
    
    body = f"""
    收件人：{receiver_name}
    
    這是由系統自動生成的檢查表彙整：{filename}
    內含所有檢查項目。
    
    (由 Streamlit 雲端系統自動發送)
    """
    msg.attach(MIMEText(body, 'plain'))
    
    # 附件類型改為 Word
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

# --- 狀態管理 ---
def init_group_photos(g_idx):
    if f"photos_{g_idx}" not in st.session_state:
        st.session_state[f"photos_{g_idx}"] = []

def add_new_photos(g_idx, uploaded_files):
    init_group_photos(g_idx)
    current_list = st.session_state[f"photos_{g_idx}"]
    existing_ids = {p['id'] for p in current_list}
    
    for f in uploaded_files:
        file_id = f"{f.name}_{f.size}"
        if file_id not in existing_ids:
            current_list.append({
                "id": file_id, "file": f, "desc": "", "result": "", "selected_opt_index": 0 
            })
            existing_ids.add(file_id)

def move_photo(g_idx, index, direction):
    lst = st.session_state[f"photos_{g_idx}"]
    new_index = index + direction
    if 0 <= new_index < len(lst):
        lst[index], lst[new_index] = lst[new_index], lst[index]

def delete_photo(g_idx, index):
    lst = st.session_state[f"photos_{g_idx}"]
    if 0 <= index < len(lst):
        del lst[index]

# --- UI ---
st.set_page_config(page_title="工程自主檢查表生成器", layout="wide")
st.title("🏗️ 工程自主檢查表 (全功能整合版)")

# Init
if 'merged_doc_buffer' not in st.session_state: st.session_state['merged_doc_buffer'] = None
if 'merged_filename' not in st.session_state: st.session_state['merged_filename'] = ""
if 'saved_template' not in st.session_state: st.session_state['saved_template'] = None
if 'checks_db' not in st.session_state: st.session_state['checks_db'] = CHECKS_DB
if 'num_groups' not in st.session_state: st.session_state['num_groups'] = 1

DEFAULT_TEMPLATE_PATH = "template.docx"
if not st.session_state['saved_template'] and os.path.exists(DEFAULT_TEMPLATE_PATH):
    with open(DEFAULT_TEMPLATE_PATH, "rb") as f:
        st.session_state['saved_template'] = f.read()

# Callbacks
def update_all_filenames():
    base_date = st.session_state['global_date']
    num = st.session_state['num_groups']
    for g in range(num):
        type_key = f"type_{g}"
        if type_key in st.session_state:
            selected_type = st.session_state[type_key]
            item_name, file_name = generate_names(selected_type, base_date)
            st.session_state[f"item_{g}"] = item_name
            st.session_state[f"fname_{g}"] = file_name

def update_group_info(g_idx):
    base_date = st.session_state['global_date']
    selected_type = st.session_state[f"type_{g_idx}"]
    item_name, file_name = generate_names(selected_type, base_date)
    st.session_state[f"item_{g_idx}"] = item_name
    st.session_state[f"fname_{g_idx}"] = file_name
    
    keys_to_clear = [k for k in st.session_state.keys() if f"_{g_idx}_" in k and (k.startswith("sel_") or k.startswith("desc_") or k.startswith("result_"))]
    for k in keys_to_clear: del st.session_state[k]
    if f"photos_{g_idx}" in st.session_state:
        for p in st.session_state[f"photos_{g_idx}"]:
            p['desc'] = ""; p['result'] = ""; p['selected_opt_index'] = 0

def clear_all_data():
    for key in list(st.session_state.keys()):
        if key.startswith(('type_', 'item_', 'fname_', 'photos_', 'file_', 'sel_', 'desc_', 'result_')):
            del st.session_state[key]
    st.session_state['num_groups'] = 1
    st.session_state['merged_doc_buffer'] = None
    st.session_state['merged_filename'] = ""

# Sidebar
with st.sidebar:
    st.header("1. 樣板設定")
    if st.session_state['saved_template']:
        st.success("✅ 樣板已載入")
    else:
        uploaded = st.file_uploader("上傳樣板", type=['docx'])
        if uploaded:
            st.session_state['saved_template'] = uploaded.getvalue()
            st.rerun()
            
    with st.expander("🛠️ 擴充資料庫"):
        uploaded_db = st.file_uploader("上傳 Excel", type=['xlsx', 'csv'])
        if uploaded_db:
            try:
                if uploaded_db.name.endswith('csv'): df = pd.read_csv(uploaded_db)
                else: df = pd.read_excel(uploaded_db)
                new_db = CHECKS_DB.copy()
                for _, row in df.iterrows():
                    cat = str(row.iloc[0]).strip()
                    item = str(row.iloc[1]).strip()
                    res = str(row.iloc[2]).strip()
                    if cat not in new_db: new_db[cat] = {"items": [], "results": []}
                    new_db[cat]["items"].append(item)
                    new_db[cat]["results"].append(res)
                st.session_state['checks_db'] = new_db
                st.success("擴充成功")
            except: st.error("讀取失敗")
    
    st.markdown("---")
    st.button("🗑️ 清除所有填寫資料", type="primary", on_click=clear_all_data)

    st.markdown("---")
    st.header("2. 專案資訊")
    p_name = st.text_input("工程名稱", "衛生福利部防疫中心興建工程")
    p_cont = st.text_input("施工廠商", "豐譽營造股份有限公司")
    
    # --- 協力廠商 下拉選單 + 輸入 ---
    sub_select = st.selectbox("協力廠商", COMMON_SUB_CONTRACTORS)
    if sub_select == "自行輸入...":
        p_sub = st.text_input("請輸入廠商名稱", "川峻工程有限公司")
    else:
        p_sub = sub_select
    # -------------------------------------
    
    p_loc = st.text_input("施作位置", "北棟 1F")
    base_date = st.date_input("日期", get_taiwan_date(), key='global_date', on_change=update_all_filenames)

# Main
if st.session_state['saved_template']:
    
    num_groups = st.number_input("本次產生幾組檢查表？", min_value=1, value=st.session_state['num_groups'], key='num_groups_input')
    st.session_state['num_groups'] = num_groups
    
    all_groups_data = []

    for g in range(num_groups):
        st.markdown(f"---")
        st.subheader(f"📂 第 {g+1} 組")
        
        c1, c2, c3 = st.columns([2, 2, 1])
        db_options = list(st.session_state['checks_db'].keys())
        selected_type = c1.selectbox(f"選擇檢查工項", db_options, key=f"type_{g}", on_change=update_group_info, args=(g,))
        
        if f"item_{g}" not in st.session_state: update_group_info(g)
            
        g_item = c2.text_input(f"自檢項目名稱", key=f"item_{g}")
        roc_year = base_date.year - 1911
        date_display = f"{roc_year}.{base_date.month:02d}.{base_date.day:02d}"
        c3.text(f"日期: {date_display}")
        file_name_custom = st.text_input("自定義檔名", key=f"fname_{g}")

        st.markdown("##### 📸 照片上傳與排序 (支援一次多選)")
        
        # --- 多選上傳模式 (動態 Key) ---
        uploader_key_name = f"uploader_key_{g}"
        if uploader_key_name not in st.session_state:
            st.session_state[uploader_key_name] = 0
            
        dynamic_key = f"uploader_{g}_{st.session_state[uploader_key_name]}"
        
        new_files = st.file_uploader(
            f"點擊此處選擇照片 (第 {g+1} 組)", 
            type=['jpg','png','jpeg'], 
            accept_multiple_files=True, 
            key=dynamic_key
        )
        
        if new_files:
            add_new_photos(g, new_files)
            st.session_state[uploader_key_name] += 1
            st.rerun()
        # --------------------------------
        
        # --- 反轉按鈕 ---
        if st.session_state.get(f"photos_{g}"):
            if st.button("🔄 順序反了嗎？點我「一鍵反轉」照片順序", key=f"rev_{g}"):
                current_list = st.session_state[f"photos_{g}"]
                for p in current_list:
                    # Sync Description
                    d_key = f"desc_{g}_{p['id']}"
                    if d_key in st.session_state:
                        p['desc'] = st.session_state[d_key]
                    
                    # Sync Result
                    r_key = f"result_{g}_{p['id']}"
                    if r_key in st.session_state:
                        p['result'] = st.session_state[r_key]
                        
                    # Sync Selection
                    s_key = f"sel_{g}_{p['id']}"
                    if s_key in st.session_state:
                        p['selected_opt_index'] = st.session_state[s_key]

                st.session_state[f"photos_{g}"].reverse()
                st.rerun()
        # ----------------------------
        
        init_group_photos(g)
        photo_list = st.session_state[f"photos_{g}"]
        
        if photo_list:
            std_items = st.session_state['checks_db'][selected_type]["items"]
            std_results = st.session_state['checks_db'][selected_type]["results"]
            options = ["(請選擇...)"] + std_items

            for i, photo_data in enumerate(photo_list):
                with st.container():
                    col_img, col_info, col_ctrl = st.columns([1.5, 3, 0.5])
                    pid = photo_data['id']
                    
                    with col_img:
                        st.image(photo_data['file'], use_container_width=True)
                        st.caption(f"No. {i+1:02d}")
                    
                    with col_info:
                        def on_select_change(pk=pid, gk=g):
                            k = f"sel_{gk}_{pk}"
                            if k not in st.session_state: return
                            new_idx = st.session_state[k]
                            dk, rk = f"desc_{gk}_{pk}", f"result_{gk}_{pk}"
                            if isinstance(new_idx, int) and new_idx > 0 and new_idx <= len(std_items):
                                st.session_state[dk] = std_items[new_idx-1]
                                st.session_state[rk] = std_results[new_idx-1]
                            else:
                                st.session_state[dk] = ""
                                st.session_state[rk] = ""

                        current_opt_idx = photo_data.get('selected_opt_index', 0)
                        if current_opt_idx > len(options): current_opt_idx = 0

                        st.selectbox("快速填寫", range(len(options)), format_func=lambda x: options[x], index=current_opt_idx, key=f"sel_{g}_{pid}", on_change=on_select_change, label_visibility="collapsed")

                        def on_text_change(field, pk=pid, idx=i, gk=g): 
                            val = st.session_state[f"{field}_{gk}_{pk}"]
                            st.session_state[f"photos_{gk}"][idx][field_map[field]] = val
                            if field == 'sel': st.session_state[f"photos_{gk}"][idx]['selected_opt_index'] = val

                        field_map = {'desc': 'desc', 'result': 'result'}
                        desc_key, result_key = f"desc_{g}_{pid}", f"result_{g}_{pid}"
                        if desc_key not in st.session_state: st.session_state[desc_key] = photo_data['desc']
                        if result_key not in st.session_state: st.session_state[result_key] = photo_data['result']

                        st.text_input("說明", key=desc_key, on_change=on_text_change, args=('desc',))
                        st.text_input("實測", key=result_key, on_change=on_text_change, args=('result',))

                    with col_ctrl:
                        if st.button("⬆️", key=f"up_{g}_{i}"): move_photo(g, i, -1); st.rerun()
                        if st.button("⬇️", key=f"down_{g}_{i}"): move_photo(g, i, 1); st.rerun()
                        if st.button("❌", key=f"del_{g}_{i}"): delete_photo(g, i); st.rerun()
                    st.divider()

            g_photos_export = []
            for i, p in enumerate(photo_list):
                d_val = st.session_state.get(f"desc_{g}_{p['id']}", p['desc'])
                r_val = st.session_state.get(f"result_{g}_{p['id']}", p['result'])
                g_photos_export.append({
                    "file": p['file'], "no": i + 1, "date_str": date_display, "desc": d_val, "result": r_val
                })

            all_groups_data.append({
                "group_id": g+1, "file_prefix": file_name_custom,
                "context": {
                    "project_name": p_name, "contractor": p_cont, "sub_contractor": p_sub,
                    "location": p_loc, "date": date_display, "check_item": g_item
                },
                "photos": g_photos_export
            })

    # --- 最終操作區 ---
    st.markdown("---")
    st.subheader("🚀 執行操作")
    
    selected_name = st.selectbox("📬 收件人", list(RECIPIENTS.keys()))
    target_email = RECIPIENTS[selected_name]

    if st.button("步驟 1：生成報告資料 (單一 Word 檔)", type="primary", use_container_width=True):
        if not all_groups_data:
            st.error("⚠️ 請至少上傳一張照片並填寫資料")
        else:
            with st.spinner("📦 正在生成並合併 Word 檔案..."):
                # --- 重大修改：使用 Composer 合併檔案 ---
                master_doc = None
                composer = None
                
                for group in all_groups_data:
                    photos = group['photos']
                    context = group['context']
                    # 每一組可能因為照片多寡產生 1 或 2 頁 (或更多)
                    # 我們這裡假設每組只會用到一次 generate_single_page (處理 8 張)
                    # 如果單組超過 8 張，您原本的邏輯是切分 batch，這裡沿用
                    
                    for page_idx, i in enumerate(range(0, len(photos), 8)):
                        batch = photos[i : i+8]
                        start_no = i + 1
                        
                        # 生成這一頁的 Doc (已包含智慧縮減)
                        current_doc = generate_single_page(st.session_state['saved_template'], context, batch, start_no)
                        
                        if master_doc is None:
                            # 第一個生成的文檔當作主文檔
                            master_doc = current_doc
                            composer = Composer(master_doc)
                        else:
                            # 之後的文檔都附加到主文檔後面
                            # 注意：docxcompose 會自動處理分頁符號
                            composer.append(current_doc)
                
                # 儲存合併後的檔案
                out_buffer = io.BytesIO()
                composer.save(out_buffer)
                
                st.session_state['merged_doc_buffer'] = out_buffer.getvalue()
                
                # 設定合併後的檔名
                roc_year = base_date.year - 1911
                date_str = f"{roc_year}{base_date.month:02d}{base_date.day:02d}"
                st.session_state['merged_filename'] = f"自主檢查表彙整_{date_str}.docx"
                
                st.success("✅ 彙整完成！所有組別已合併為單一 Word 檔。")

    if st.session_state['merged_doc_buffer']:
        col_mail, col_dl = st.columns(2)
        
        with col_mail:
            if st.button(f"📧 立即寄出 Word 檔給：{selected_name}", use_container_width=True):
                with st.spinner("📨 雲端發信中..."):
                    success, msg = send_email_via_secrets(
                        st.session_state['merged_doc_buffer'], 
                        st.session_state['merged_filename'],
                        target_email,
                        selected_name
                    )
                    if success:
                        st.success(msg)
                    else:
                        st.error(msg)
        
        with col_dl:
            st.download_button(
                label="📥 下載 Word 檔案", 
                data=st.session_state['merged_doc_buffer'], 
                file_name=st.session_state['merged_filename'], 
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                use_container_width=True
            )

else:
    st.info("👈 請先在左側確認 Word 樣板")
