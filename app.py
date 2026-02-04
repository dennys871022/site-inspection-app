import streamlit as st
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml.ns import qn
from PIL import Image
import io
import datetime
import os
import zipfile

# --- 核心邏輯 (維持不變) ---

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
            
            # 使用 6 個全形空白對齊
            spacer = "\u3000" * 6
            info_text = f"照片編號：{data['no']:02d}{spacer}日期：{data['date_str']}\n"
            info_text += f"說明：{data['desc']}\n"
            info_text += f"實測：{data['result']}"
            
            replace_text_content(doc, {info_key: info_text})
        else:
            replace_text_content(doc, {img_key: ""})
            replace_text_content(doc, {info_key: ""})
    return doc

# --- Streamlit UI (效率優化版) ---

st.set_page_config(page_title="自主檢查表生成器", layout="wide")
st.title("🚀 工程自主檢查表 (極速預覽版)")

# State Init
if 'zip_buffer' not in st.session_state: st.session_state['zip_buffer'] = None
if 'saved_template' not in st.session_state: st.session_state['saved_template'] = None

# Auto Load Template
DEFAULT_TEMPLATE_PATH = "template.docx"
if not st.session_state['saved_template'] and os.path.exists(DEFAULT_TEMPLATE_PATH):
    with open(DEFAULT_TEMPLATE_PATH, "rb") as f:
        st.session_state['saved_template'] = f.read()

# Sidebar
with st.sidebar:
    st.header("1. 設定")
    if st.session_state['saved_template']:
        st.success("✅ 樣板已載入")
    else:
        uploaded = st.file_uploader("上傳樣板", type=['docx'])
        if uploaded:
            st.session_state['saved_template'] = uploaded.getvalue()
            st.rerun()
            
    st.markdown("---")
    st.header("2. 專案資訊")
    p_name = st.text_input("工程名稱 {project_name}", "衛生福利部防疫中心興建工程")
    p_cont = st.text_input("施工廠商 {contractor}", "豐譽營造股份有限公司")
    p_sub = st.text_input("協力廠商 {sub_contractor}", "川峻工程有限公司")
    p_loc = st.text_input("施作位置 {location}", "北棟 1F")
    base_date = st.date_input("日期", datetime.date.today())

# Main Area
if st.session_state['saved_template']:
    
    st.info("💡 只要輸入一次「預設說明」，所有照片都會自動套用，除非您手動修改。")
    
    # --- 群組管理 ---
    num_groups = st.number_input("本次產生幾組檢查表？", min_value=1, value=1)
    all_groups_data = []

    for g in range(num_groups):
        st.markdown(f"### 📂 第 {g+1} 組")
        
        # 1. 快速設定區
        c1, c2, c3, c4 = st.columns([2, 1, 1.5, 1.5])
        g_item = c1.text_input(f"自檢項目", value=f"拆除工程施工自主檢查", key=f"item_{g}")
        
        # 日期轉換
        roc_year = base_date.year - 1911
        g_date_str = f"{roc_year}.{base_date.month:02d}.{base_date.day:02d}"
        
        # 預設值設定 (加速關鍵)
        def_desc = c3.text_input("預設說明 (套用全部)", value="現場既有雜物整理", key=f"def_d_{g}")
        def_res = c4.text_input("預設實測 (套用全部)", value="現場既有雜物整理", key=f"def_r_{g}")

        # 2. 照片上傳
        g_files = st.file_uploader(f"上傳第 {g+1} 組照片", type=['jpg','png','jpeg'], accept_multiple_files=True, key=f"file_{g}")
        
        if g_files:
            st.write(f"共 {len(g_files)} 張照片")
            
            g_photos = []
            # 使用 Expander 預設展開，但排版緊湊
            with st.expander("📸 檢視與微調照片 (已自動填入預設值)", expanded=True):
                # 建立一個容器，每行顯示 2 張圖
                for i in range(0, len(g_files), 2):
                    row_cols = st.columns(2)
                    for j in range(2):
                        if i + j >= len(g_files): break
                        
                        file = g_files[i+j]
                        no = i + j + 1
                        
                        with row_cols[j]:
                            # --- 預覽與編輯區 (左右並排) ---
                            img_col, input_col = st.columns([1, 2])
                            
                            with img_col:
                                # 顯示縮圖
                                st.image(file, use_container_width=True)
                                st.caption(f"No. {no}")
                            
                            with input_col:
                                # 輸入框
                                d_val = st.text_input(f"說明", value=def_desc, key=f"d_{g}_{no}", label_visibility="collapsed", placeholder="說明")
                                r_val = st.text_input(f"實測", value=def_res, key=f"r_{g}_{no}", label_visibility="collapsed", placeholder="實測")
                                
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
    if st.button("🚀 立即生成並下載", type="primary", use_container_width=True):
        if not all_groups_data:
            st.error("請上傳照片")
        else:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for group in all_groups_data:
                    g_id = group['group_id']
                    photos = group['photos']
                    context = group['context']
                    item_safe_name = context['check_item'].replace("/", "_")
                    
                    # 自動分頁邏輯 (每8張一頁)
                    for page_idx, i in enumerate(range(0, len(photos), 8)):
                        batch = photos[i : i+8]
                        start_no = i + 1
                        doc = generate_single_page(st.session_state['saved_template'], context, batch, start_no)
                        
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        
                        suffix = f"_{page_idx+1}" if len(photos) > 8 else ""
                        fname = f"{g_date_str}_{p_loc}_{item_safe_name}{suffix}.docx"
                        zf.writestr(fname, doc_io.getvalue())
            
            st.session_state['zip_buffer'] = zip_buffer.getvalue()
            st.success("✅ 完成！")

    if st.session_state['zip_buffer']:
        st.download_button(
            label="📥 下載 ZIP 壓縮檔",
            data=st.session_state['zip_buffer'],
            file_name=f"檢查報告_{datetime.date.today()}.zip",
            mime="application/zip",
            use_container_width=True
        )
