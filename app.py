import streamlit as st
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image
import io
import datetime
import os
import zipfile

# --- 0. 標準化檢查項目資料庫 (依據您上傳的 PDF 建立) ---
STANDARD_CHECKS = {
    "通用/自訂": {
        "items": ["現場既有雜物整理", "依施工計畫執行", "其他"],
        "results": ["現場既有雜物整理", "與計畫相符", "符合規定"]
    },
    "拆除工程 (EA26)": {
        "items": [
            "防塵作為:灑水或防塵網",
            "降噪作為:低噪音型機具",
            "構造物拆除順序(由上而下)",
            "保留構造不得損傷",
            "拆除物分類(可回收/不可回收/有價)",
            "車輛輪胎清潔",
            "安全監測(傾斜計/沉陷點)",
            "廢棄物清運(依核定計畫)",
            "地坪裝修材剃除"
        ],
        "results": [
            "備有灑水車/防塵網抑塵",
            "使用低噪音機具(大鋼牙破碎)",
            "依施工規劃由上而下拆除",
            "保留構造無損傷",
            "已依類別分類置放",
            "備有專人清潔輪胎，無帶汙泥出場",
            "監測數值在安全範圍內",
            "依核定計畫書執行清運",
            "地坪裝修材已剃除乾淨"
        ]
    },
    "微型樁工程 (EA53)": {
        "items": [
            "開挖前置作業(管線確認)",
            "樁心檢測 (≦3cm)",
            "鑽掘垂直度 (TYPE I: 0-5°)",
            "鑽掘深度 (L≧16m)",
            "鑽掘樁徑 (ψ≧15cm)",
            "鑽掘間距 (@60cm 交錯)",
            "水泥漿拌合比 (W/C=1:1)",
            "注漿時間 (≦10min)",
            "鋼管吊放 (L=16m, 間隔器@2m)",
            "廢漿清除",
            "樁頂劣質打石",
            "帽梁鋼筋綁紮 (#6-4支, #3@20cm)"
        ],
        "results": [
            "確認開挖區域無地下管線",
            "樁心偏差符合規定 (≦3cm)",
            "垂直度符合規定",
            "鑽掘深度符合設計 (16m)",
            "樁徑實測符合規定",
            "間距符合規定 (@60cm)",
            "拌合比例正確",
            "注漿連續，時間符合規定",
            "鋼管長度及間隔器安裝正確",
            "廢漿已清除",
            "劣質混凝土已打除",
            "鋼筋綁紮符合設計圖說"
        ]
    },
    "有價廢料載運 (EB26)": {
        "items": [
            "廢鋼筋載運",
            "銅線/銅製品載運",
            "電線電纜(裹外皮)載運",
            "型鋼載運",
            "鋁料載運",
            "空車重量查核",
            "載運後總重查核",
            "有價廢料淨重確認"
        ],
        "results": [
            "載運廢鋼筋 * 1車",
            "載運銅製品 * 1車",
            "載運電纜 * 1車",
            "載運型鋼 * 1車",
            "載運鋁料 * 1車",
            "空車重量: _____ kg",
            "載運總重: _____ kg",
            "有價物淨重: _____ kg"
        ]
    }
}

# --- 1. 樣式複製核心 ---

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

# --- 2. 替換邏輯 ---

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
            
            spacer = "\u3000" * 6
            info_text = f"照片編號：{data['no']:02d}{spacer}日期：{data['date_str']}\n"
            info_text += f"說明：{data['desc']}\n"
            info_text += f"實測：{data['result']}"
            
            replace_text_content(doc, {info_key: info_text})
        else:
            replace_text_content(doc, {img_key: ""})
            replace_text_content(doc, {info_key: ""})
    return doc

# --- 4. Streamlit UI ---

st.set_page_config(page_title="工程自主檢查表生成器", layout="wide")
st.title("🏗️ 工程自主檢查表 (智能選單版)")

# Init
if 'zip_buffer' not in st.session_state: st.session_state['zip_buffer'] = None
if 'saved_template' not in st.session_state: st.session_state['saved_template'] = None

DEFAULT_TEMPLATE_PATH = "template.docx"
if not st.session_state['saved_template'] and os.path.exists(DEFAULT_TEMPLATE_PATH):
    with open(DEFAULT_TEMPLATE_PATH, "rb") as f:
        st.session_state['saved_template'] = f.read()

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
            
    st.markdown("---")
    st.header("2. 專案資訊")
    p_name = st.text_input("工程名稱", "衛生福利部防疫中心興建工程")
    p_cont = st.text_input("施工廠商", "豐譽營造股份有限公司")
    p_sub = st.text_input("協力廠商", "川峻工程有限公司")
    p_loc = st.text_input("施作位置", "北棟 1F")
    base_date = st.date_input("日期", datetime.date.today())

# Main
if st.session_state['saved_template']:
    
    # 群組設定
    num_groups = st.number_input("本次產生幾組檢查表？", min_value=1, value=1)
    all_groups_data = []

    for g in range(num_groups):
        st.markdown(f"---")
        st.subheader(f"📂 第 {g+1} 組檢查")
        
        # 1. 選擇檢查類型 (決定下拉選單內容)
        c1, c2, c3 = st.columns([2, 2, 1])
        
        # 讓使用者選擇這組是要檢查什麼
        check_type = c1.selectbox(
            f"選擇檢查類別", 
            list(STANDARD_CHECKS.keys()), 
            index=1 if g==0 else 0, # 預設選第二個(拆除)方便測試
            key=f"type_{g}"
        )
        
        # 自動帶入對應的預設項目名稱
        default_item_name = check_type.split(" ")[0] + "自主檢查"
        g_item = c2.text_input(f"自檢項目名稱 {{check_item}}", value=default_item_name, key=f"item_{g}")
        
        # 日期
        roc_year = base_date.year - 1911
        g_date_str = f"{roc_year}.{base_date.month:02d}.{base_date.day:02d}"
        c3.text(f"日期: {g_date_str}")

        # 2. 上傳照片
        g_files = st.file_uploader(f"上傳照片 (第 {g+1} 組)", type=['jpg','png','jpeg'], accept_multiple_files=True, key=f"file_{g}")
        
        if g_files:
            st.info(f"已上傳 {len(g_files)} 張照片。請使用下方選單快速填寫。")
            
            g_photos = []
            
            # 取得該類別的標準清單
            std_items = STANDARD_CHECKS[check_type]["items"]
            std_results = STANDARD_CHECKS[check_type]["results"]
            
            # 兩欄排列照片編輯器
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
                            # --- 關鍵功能：下拉選單 ---
                            # 加一個 "自訂" 選項
                            options = ["(請選擇檢查項目...)"] + std_items
                            selected_opt = st.selectbox(
                                "快速選擇", 
                                options, 
                                label_visibility="collapsed", 
                                key=f"sel_{g}_{no}"
                            )
                            
                            # 根據選擇自動填入文字
                            current_desc = ""
                            current_res = ""
                            
                            if selected_opt != "(請選擇檢查項目...)":
                                idx = std_items.index(selected_opt)
                                current_desc = std_items[idx]
                                current_res = std_results[idx]
                            
                            # 允許使用者手動修改 (如果沒選，就留白讓使用者打)
                            d_val = st.text_input("說明", value=current_desc, key=f"d_{g}_{no}", placeholder="說明")
                            r_val = st.text_input("實測", value=current_res, key=f"r_{g}_{no}", placeholder="實測")
                            
                            g_photos.append({
                                "file": file, "no": no, "date_str": g_date_str,
                                "desc": d_val, "result": r_val
                            })
                        st.divider()

            all_groups_data.append({
                "group_id": g+1,
                "context": {
                    "project_name": p_name, "contractor": p_cont, 
                    "sub_contractor": p_sub, "location": p_loc, 
                    "date": g_date_str, "check_item": g_item
                },
                "photos": g_photos
            })

    # 生成按鈕
    st.markdown("---")
    if st.button("🚀 立即生成並下載報告", type="primary", use_container_width=True):
        if not all_groups_data:
            st.error("請至少完成一組照片上傳")
        else:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for group in all_groups_data:
                    g_id = group['group_id']
                    photos = group['photos']
                    context = group['context']
                    # 檔名處理 (移除不合法字元)
                    safe_name = context['check_item'].replace("/", "_").replace("\\", "_")
                    
                    # 分頁處理
                    for page_idx, i in enumerate(range(0, len(photos), 8)):
                        batch = photos[i : i+8]
                        start_no = i + 1
                        doc = generate_single_page(st.session_state['saved_template'], context, batch, start_no)
                        
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        
                        suffix = f"_{page_idx+1}" if len(photos) > 8 else ""
                        fname = f"組別{g_id}_{safe_name}{suffix}.docx"
                        zf.writestr(fname, doc_io.getvalue())
            
            st.session_state['zip_buffer'] = zip_buffer.getvalue()
            st.success("✅ 報告生成完畢！")

    if st.session_state['zip_buffer']:
        st.download_button(
            label="📥 下載所有報告 (.zip)",
            data=st.session_state['zip_buffer'],
            file_name=f"檢查報告_{datetime.date.today()}.zip",
            mime="application/zip",
            use_container_width=True
        )
