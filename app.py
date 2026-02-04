import streamlit as st
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml.ns import qn
from PIL import Image
import io
import datetime
import os
import zipfile
import pandas as pd

# --- 0. 終極內建資料庫 ---
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
            "鋁料載運", "載運車輛資訊", "重量查核"
        ],
        "results": [
            "載運廢鋼筋 * 1 車", "載運銅製品 * 1 車", "載運電纜 * 1 車", "載運型鋼 * 1 車", 
            "載運鋁料 * 1 車", "車號：__________", "總重:____kg / 淨重:____kg"
        ]
    },
    "微型樁工程-施工 (EA53)": {
        "items": [
            "開挖前置:管線確認", "樁心檢測 (≦3cm)", "鑽掘垂直度 (0-5度)",
            "鑽掘尺寸 (深度/樁徑)", "鑽掘間距 (@60cm)", "水泥漿拌合比 (1:1)",
            "注漿作業 (≦10min)", "鋼管吊放安裝", "廢漿清除", "樁頂劣質打石", 
            "帽梁鋼筋綁紮", "帽梁灌漿"
        ],
        "results": [
            "確認無地下管線干擾", "樁心偏差 ≦3cm", "垂直度符合規定 (0-5度)",
            "深度≧16m; 樁徑≧15cm", "間距@60cm, 交錯施工", "水灰比 W/C=1:1",
            "時間≦10min，注漿至帽梁底部", "長度16m; 間隔器@2m", "已清除硬固廢漿",
            "劣質混凝土已打除", "主筋#6-4支, 箍筋#3@20cm", "強度 fc'=210kgf/cm2"
        ]
    },
    "微型樁工程-材料 (EB53)": {
        "items": ["證明文件", "規格尺寸", "外觀形狀", "工地放置", "取樣試驗"],
        "results": ["出廠證明/檢驗紀錄齊全", "符合契約規範", "無碰撞變形", "分類堆置/標示", "依規範取樣"]
    },
    "排樁工程-施工 (EA54)": {
        "items": [
            "樁心定位檢測", "預壘樁鑽掘(長度/直徑)", "鋼筋籠製作(主筋/箍筋)",
            "鋼筋籠搭接與間隔", "水泥砂漿試體製作", "預壘樁灌漿高程",
            "微型樁鑽掘(垂直/深度)", "微型樁注漿/鋼管", "壓梁鋼筋綁紮", "壓梁混凝土澆置"
        ],
        "results": [
            "偏差 ±2cm 以內", "長度/直徑符合設計圖說", "主筋#8/#7; 箍筋#4 符合規定",
            "搭接≧8cm; 間隔片@200cm", "已製作方塊試體", "高程≧樁長; 壓力≧2.1kgf/cm2",
            "垂直度±5度; 深度≧7m", "水灰比1:1; 鋼管L=7m", "主筋#6; 箍筋#4@15cm", "強度 210kgf/cm2, 坍度20±4cm"
        ]
    },
    "排樁工程-材料 (EB54)": {
        "items": ["證明文件", "規格尺寸", "外觀形狀", "工地放置", "取樣試驗"],
        "results": ["出廠證明/檢驗紀錄齊全", "符合契約規範", "無碰撞變形", "分類堆置/標示", "依規範取樣"]
    },
    "假設工程-施工 (EA51)": {
        "items": [
            "放樣", "全阻式圍籬組立", "半阻式圍籬組立", "防溢座施作",
            "出入口地坪(鋼筋/澆置)", "大門安裝", "安全走廊", "警示燈設置",
            "洗車台尺寸檢查", "圍籬綠化維護"
        ],
        "results": [
            "依施工圖說放樣", "間距/埋入深度符合規定", "間距/埋入深度符合規定", "混凝土210kgf/cm2",
            "厚度20cm; 雙層雙向#4@10cm", "尺寸及埋入深度符合規定", "高300寬150cm",
            "間距符合規定", "500x522cm; 沉沙池深170cm", "存活率90%以上"
        ]
    },
    "假設工程-材料 (EB51)": {
        "items": ["證明文件", "外觀形狀", "工地放置", "預鑄水溝尺寸"],
        "results": ["出廠證明/檢驗紀錄齊全", "無碰撞變形、破損", "分類堆置/標示", "內溝寬30±5cm, 深40±5cm"]
    },
    "車道拓寬工程 (EA52)": {
        "items": [
            "碎石級配舖設", "鋼筋綁紮", "模板組立", "混凝土澆置(結構)",
            "粉刷面清潔", "基準灰誌製作", "馬賽克磚舖貼", "瀝青混凝土舖設"
        ],
        "results": [
            "級配高度 20cm", "箍筋#4@20cm; 保護層4cm", "牆厚20cm; 垂直度±13mm", "強度 210kgf/cm2",
            "無殘餘雜物、凸出物", "間距不大於1M", "顏色與樣板相同", "密級配，無汙損浮起"
        ]
    },
    "混凝土工程 (共用)": {
        "items": [
            "照明與雨天防護", "澆置前清潔濕潤", "模板振動器", "澆置時間控制",
            "坍度/流度檢查", "溫度檢查", "氯離子含量", "試體取樣", "振動搗實", "養護作業"
        ],
        "results": [
            "照明充足，備有防雨材", "垃圾清除，模板濕潤", "備有至少二具", "拌合至澆置90分鐘內",
            "符合設計 (如 18±4cm)", "13~32度C", "小於 0.15 kg/m3", "每100m3取樣1組",
            "間距<50cm; 每次5-10秒", "灑水或覆蓋養護"
        ]
    }
}

# --- 1. 樣式與影像處理 ---

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
        except:
            pass
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
    if img.mode == 'RGBA':
        img = img.convert('RGB')
    try:
        from PIL import ImageOps
        img = ImageOps.exif_transpose(img)
    except:
        pass
    ratio = max_width / float(img.size[0])
    if ratio < 1:
        h_size = int((float(img.size[1]) * float(ratio)))
        img = img.resize((max_width, h_size), Image.Resampling.LANCZOS)
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='JPEG', quality=75)
    img_byte_arr.seek(0)
    return img_byte_arr

# --- 2. 替換邏輯 (純淨樣式) ---

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
            
            # 日期前 6 個全形空白
            spacer = "\u3000" * 6 
            info_text = f"照片編號：{data['no']:02d}{spacer}日期：{data['date_str']}\n"
            info_text += f"說明：{data['desc']}\n"
            info_text += f"實測：{data['result']}"
            
            replace_text_content(doc, {info_key: info_text})
        else:
            replace_text_content(doc, {img_key: ""})
            replace_text_content(doc, {info_key: ""})
    return doc

# --- 4. 智慧命名邏輯 ---

def generate_auto_names(selected_type, base_date):
    """
    根據選擇的工項，自動生成符合標準的名稱。
    格式：[工項名稱][類型]自主檢查
    檔名：[日期][工項名稱][類型]自主檢查
    """
    # 解析選單字串，例如 "拆除工程-施工 (EA26)"
    # 取出 "拆除工程"
    main_name = selected_type.split('-')[0]
    
    # 判斷後綴
    suffix = "自主檢查"
    if "施工" in selected_type:
        suffix = "施工自主檢查"
    elif "材料" in selected_type:
        suffix = "材料進場自主檢查"
    elif "有價廢料" in selected_type:
        suffix = "有價廢料清運自主檢查"
    elif "混凝土" in selected_type:
        # 特例處理
        suffix = "施工自主檢查"
        
    full_item_name = f"{main_name}{suffix}"
    
    # 日期字串 (民國年無分隔符)
    roc_year = base_date.year - 1911
    roc_date_str = f"{roc_year}{base_date.month:02d}{base_date.day:02d}"
    
    file_name = f"{roc_date_str}{full_item_name}"
    
    return full_item_name, file_name

# --- 5. Streamlit UI ---

st.set_page_config(page_title="工程自主檢查表生成器", layout="wide")
st.title("🏗️ 工程自主檢查表 (標準命名版)")

# Init
if 'zip_buffer' not in st.session_state: st.session_state['zip_buffer'] = None
if 'saved_template' not in st.session_state: st.session_state['saved_template'] = None
if 'checks_db' not in st.session_state: st.session_state['checks_db'] = CHECKS_DB

DEFAULT_TEMPLATE_PATH = "template.docx"
if not st.session_state['saved_template'] and os.path.exists(DEFAULT_TEMPLATE_PATH):
    with open(DEFAULT_TEMPLATE_PATH, "rb") as f:
        st.session_state['saved_template'] = f.read()

# --- Callback ---
def update_group_defaults(g_idx, base_date):
    """類別或日期改變時，更新名稱"""
    type_key = f"type_{g_idx}"
    item_key = f"item_{g_idx}"
    fname_key = f"fname_{g_idx}"
    
    selected_type = st.session_state[type_key]
    
    # 呼叫命名邏輯
    item_name, file_name = generate_auto_names(selected_type, base_date)
    
    st.session_state[item_key] = item_name
    st.session_state[fname_key] = file_name

def update_photo_defaults(g_idx, p_no):
    """照片選單改變時，更新說明"""
    sel_key = f"sel_{g_idx}_{p_no}"
    desc_key = f"d_{g_idx}_{p_no}"
    res_key = f"r_{g_idx}_{p_no}"
    type_key = f"type_{g_idx}"
    
    selected_opt = st.session_state[sel_key]
    current_type = st.session_state[type_key]
    
    if selected_opt != "(請選擇...)":
        items = st.session_state['checks_db'][current_type]["items"]
        results = st.session_state['checks_db'][current_type]["results"]
        if selected_opt in items:
            idx = items.index(selected_opt)
            st.session_state[desc_key] = items[idx]
            st.session_state[res_key] = results[idx]
    else:
        st.session_state[desc_key] = ""
        st.session_state[res_key] = ""

# --- Sidebar ---
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
            except:
                st.error("讀取失敗")

    st.markdown("---")
    st.header("2. 專案資訊")
    p_name = st.text_input("工程名稱", "衛生福利部防疫中心興建工程")
    p_cont = st.text_input("施工廠商", "豐譽營造股份有限公司")
    p_sub = st.text_input("協力廠商", "川峻工程有限公司")
    p_loc = st.text_input("施作位置", "北棟 1F")
    
    # 日期選擇 (綁定 Rerun，讓所有組別檔名自動更新)
    base_date = st.date_input("日期", datetime.date.today())

# --- Main ---
if st.session_state['saved_template']:
    
    num_groups = st.number_input("本次產生幾組檢查表？", min_value=1, value=1)
    all_groups_data = []

    for g in range(num_groups):
        st.markdown(f"---")
        st.subheader(f"📂 第 {g+1} 組")
        
        c1, c2, c3 = st.columns([2, 2, 1])
        
        # 1. 選擇工項
        db_options = list(st.session_state['checks_db'].keys())
        selected_type = c1.selectbox(
            f"選擇檢查工項", 
            db_options, 
            key=f"type_{g}",
            on_change=update_group_defaults,
            args=(g, base_date)
        )
        
        # 初次載入或重新整理時，確保檔名正確
        if f"item_{g}" not in st.session_state:
            update_group_defaults(g, base_date)
            
        # 2. 自動產生的欄位
        g_item = c2.text_input(f"自檢項目名稱 {{check_item}}", key=f"item_{g}")
        
        roc_year = base_date.year - 1911
        date_display = f"{roc_year}.{base_date.month:02d}.{base_date.day:02d}"
        c3.text(f"日期: {date_display}")
        
        # 3. 檔名自定義
        file_name_custom = st.text_input("自定義檔名 (下載時使用)", key=f"fname_{g}")

        # 4. 照片上傳
        g_files = st.file_uploader(f"上傳照片", type=['jpg','png','jpeg'], accept_multiple_files=True, key=f"file_{g}")
        
        if g_files:
            g_photos = []
            std_items = st.session_state['checks_db'][selected_type]["items"]
            
            for i in range(0, len(g_files), 2):
                row_cols = st.columns(2)
                for j in range(2):
                    if i + j >= len(g_files): break
                    file = g_files[i+j]
                    no = i + j + 1
                    
                    with row_cols[j]:
                        img_col, input_col = st.columns([1, 2])
                        with img_col:
                            st.image(file, use_container_width=True)
                            st.caption(f"No. {no}")
                        
                        with input_col:
                            options = ["(請選擇...)"] + std_items
                            def_idx = no if no <= len(std_items) else 0
                            
                            if f"d_{g}_{no}" not in st.session_state:
                                st.session_state[f"d_{g}_{no}"] = ""
                                st.session_state[f"r_{g}_{no}"] = ""
                            
                            selected_opt = st.selectbox(
                                "快速選擇", options, index=def_idx, 
                                label_visibility="collapsed", 
                                key=f"sel_{g}_{no}",
                                on_change=update_photo_defaults,
                                args=(g, no)
                            )
                            
                            if st.session_state[f"d_{g}_{no}"] == "" and selected_opt != "(請選擇...)":
                                update_photo_defaults(g, no)

                            d_val = st.text_input("說明", key=f"d_{g}_{no}")
                            r_val = st.text_input("實測", key=f"r_{g}_{no}")
                            
                            g_photos.append({
                                "file": file, "no": no, "date_str": date_display,
                                "desc": d_val, "result": r_val
                            })
                        st.divider()

            all_groups_data.append({
                "group_id": g+1,
                "file_prefix": file_name_custom,
                "context": {
                    "project_name": p_name, "contractor": p_cont, 
                    "sub_contractor": p_sub, "location": p_loc, 
                    "date": date_display, "check_item": g_item
                },
                "photos": g_photos
            })

    # 生成按鈕
    st.markdown("---")
    if st.button("🚀 立即生成並下載", type="primary", use_container_width=True):
        if not all_groups_data:
            st.error("請上傳照片")
        else:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for group in all_groups_data:
                    photos = group['photos']
                    context = group['context']
                    file_prefix = group['file_prefix']
                    
                    for page_idx, i in enumerate(range(0, len(photos), 8)):
                        batch = photos[i : i+8]
                        start_no = i + 1
                        doc = generate_single_page(st.session_state['saved_template'], context, batch, start_no)
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        suffix = f"_{page_idx+1}" if len(photos) > 8 else ""
                        fname = f"{file_prefix}{suffix}.docx"
                        zf.writestr(fname, doc_io.getvalue())
            
            st.session_state['zip_buffer'] = zip_buffer.getvalue()
            st.success("✅ 完成！")

    if st.session_state['zip_buffer']:
        st.download_button(
            label="📥 下載 ZIP 檔",
            data=st.session_state['zip_buffer'],
            file_name=f"自檢表_{datetime.date.today()}.zip",
            mime="application/zip",
            use_container_width=True
        )
else:
    st.info("👈 請先在左側確認 Word 樣板")
