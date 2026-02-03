import streamlit as st
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml.ns import qn
from PIL import Image
import io
import datetime
import os
import zipfile  # 新增：用於打包多個檔案

# --- 1. 樣式複製核心 (保持不變，確保格式完美) ---

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

# --- 3. 單頁生成核心 ---

def generate_single_page(template_bytes, context, photo_batch, start_no):
    """生成單一頁面的 Word 檔 (處理 1~8 張)"""
    doc = Document(io.BytesIO(template_bytes))
    
    # 1. 填入基本資料
    text_replacements = {f"{{{k}}}": v for k, v in context.items()}
    replace_text_content(doc, text_replacements)
    
    # 2. 填入照片
    # 樣板固定只有 {img_1}~{img_8}
    for i in range(1, 9):
        img_key = f"{{img_{i}}}"
        info_key = f"{{info_{i}}}"
        
        idx = i - 1
        if idx < len(photo_batch):
            data = photo_batch[idx]
            
            # 填入圖片
            replace_placeholder_with_image(doc, img_key, compress_image(data['file']))
            
            # 填入文字 (計算連續編號)
            current_no = start_no + idx
            spacer = "\u3000" * 7 # 日期對齊用
            
            info_text = f"照片編號：{current_no:02d}{spacer}日期：{data['date_str']}\n"
            info_text += f"說明：{data['desc']}\n"
            info_text += f"實測：{data['result']}"
            
            replace_text_content(doc, {info_key: info_text})
        else:
            # 沒照片就清空
            replace_text_content(doc, {img_key: ""})
            replace_text_content(doc, {info_key: ""})
            
    return doc

# --- 4. Streamlit UI ---

st.set_page_config(page_title="自主檢查表生成器", layout="wide")
st.title("🏗️ 工程自主檢查表 (多組自動分頁版)")

# 初始化
if 'zip_buffer' not in st.session_state:
    st.session_state['zip_buffer'] = None
if 'saved_template' not in st.session_state:
    st.session_state['saved_template'] = None
    
# 自動載入
DEFAULT_TEMPLATE_PATH = "template.docx"
if not st.session_state['saved_template'] and os.path.exists(DEFAULT_TEMPLATE_PATH):
    with open(DEFAULT_TEMPLATE_PATH, "rb") as f:
        st.session_state['saved_template'] = f.read()

# --- 側邊欄設定 ---
with st.sidebar:
    st.header("1. 樣板設定")
    if st.session_state['saved_template']:
        st.success(f"✅ 樣板就緒")
    else:
        st.warning("⚠️ 請上傳 template.docx")
        uploaded = st.file_uploader("上傳樣板", type=['docx'])
        if uploaded:
            st.session_state['saved_template'] = uploaded.getvalue()
            st.rerun()

    st.markdown("---")
    st.header("2. 通用專案資訊")
    p_name = st.text_input("工程名稱 {project_name}", "衛生福利部防疫中心興建工程")
    p_cont = st.text_input("施工廠商 {contractor}", "豐譽營造股份有限公司")
    p_sub = st.text_input("協力廠商 {sub_contractor}", "川峻工程有限公司")
    p_loc = st.text_input("施作位置 {location}", "北棟 1F")
    base_date = st.date_input("預設檢查日期", datetime.date.today())

# --- 主畫面區 ---
if st.session_state['saved_template']:
    
    # 設定組數
    num_groups = st.number_input("📋 請問今天要產生幾組檢查表？", min_value=1, value=1, step=1)
    
    all_groups_data = [] # 儲存所有要生成的資料
    
    # 動態產生輸入表單
    for g in range(num_groups):
        with st.expander(f"📂 第 {g+1} 組檢查設定", expanded=(g==0)):
            c1, c2 = st.columns([2, 1])
            # 讓每組可以有不同的項目名稱
            g_item = c1.text_input(f"自檢項目 (第 {g+1} 組) {{check_item}}", 
                                   value=f"拆除工程施工自主檢查 #{g+1}", key=f"item_{g}")
            g_date = c2.date_input(f"日期", value=base_date, key=f"date_{g}")
            
            # 民國年
            roc_year = g_date.year - 1911
            g_date_str = f"{roc_year}.{g_date.month:02d}.{g_date.day:02d}"
            
            # 照片上傳
            g_files = st.file_uploader(f"上傳第 {g+1} 組照片 (超過 8 張會自動分頁)", 
                                       type=['jpg','png','jpeg'], accept_multiple_files=True, key=f"file_{g}")
            
            if g_files:
                st.info(f"已選擇 {len(g_files)} 張照片，將自動產生 {(len(g_files)-1)//8 + 1} 頁 Word 檔。")
                
                # 照片詳細資訊編輯 (批次)
                # 為了版面整潔，這裡只提供一個統一設定，或展開編輯
                with st.expander("✏️ 編輯照片說明 (選填)", expanded=False):
                    g_photos = []
                    for i, f in enumerate(g_files):
                        st.markdown(f"**照片 {i+1}** ({f.name})")
                        col_a, col_b = st.columns(2)
                        desc = col_a.text_input("說明", value="現場既有雜物整理", key=f"d_{g}_{i}")
                        res = col_b.text_input("實測", value="現場既有雜物整理", key=f"r_{g}_{i}")
                        g_photos.append({
                            "file": f, "desc": desc, "result": res, "date_str": g_date_str
                        })
                
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
    if st.button("🚀 開始生成所有報告", type="primary"):
        if not all_groups_data:
            st.error("請至少上傳一組照片！")
        else:
            # 建立 ZIP 檔案
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                
                for group in all_groups_data:
                    g_id = group['group_id']
                    photos = group['photos']
                    context = group['context']
                    item_name = context['check_item'].replace("/", "_") # 檔名安全處理
                    
                    # 計算分頁 (每 8 張一頁)
                    # chunk size = 8
                    for page_idx, i in enumerate(range(0, len(photos), 8)):
                        batch = photos[i : i+8]
                        start_no = i + 1 # 這一頁的起始編號 (例如第2頁從9開始)
                        
                        # 生成這一頁
                        doc = generate_single_page(st.session_state['saved_template'], context, batch, start_no)
                        
                        # 存成 Bytes
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        
                        # 檔名邏輯：如果有分頁，加上 (Page X)
                        page_suffix = f"_Page{page_idx+1}" if len(photos) > 8 else ""
                        file_name = f"Group{g_id}_{item_name}{page_suffix}.docx"
                        
                        # 加入 ZIP
                        zf.writestr(file_name, doc_io.getvalue())
            
            st.session_state['zip_buffer'] = zip_buffer.getvalue()
            st.success("✅ 全部生成完畢！請下載 ZIP 檔。")

    # 下載按鈕
    if st.session_state['zip_buffer']:
        st.download_button(
            label="📥 下載所有報告 (.zip)",
            data=st.session_state['zip_buffer'],
            file_name=f"檢查報告_{datetime.date.today()}.zip",
            mime="application/zip"
        )

else:
    st.info("👈 請先確認樣板已載入")
