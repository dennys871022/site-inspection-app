import streamlit as st
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from PIL import Image
import io
import datetime

# --- 1. 基礎設定 ---

def set_font_style(run, font_name='標楷體', size=12):
    """設定字型"""
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)

def compress_image(image_file, max_width=800):
    """圖片壓縮與轉向處理"""
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

# --- 2. 核心功能：精準填空 ---

def replace_text_content(doc, replacements):
    """
    通用文字替換：將 {key} 換成 value
    適用於：工程名稱、位置、說明文字等
    """
    # 遍歷表格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_paragraph(paragraph, replacements)
    # 遍歷一般段落
    for paragraph in doc.paragraphs:
        replace_paragraph(paragraph, replacements)

def replace_paragraph(paragraph, replacements):
    for key, value in replacements.items():
        if key in paragraph.text:
            # 這裡使用簡單替換，保留段落格式
            if value is None: value = ""
            paragraph.text = paragraph.text.replace(key, str(value))
            # 重新設定字型 (因為替換後格式有時會跑掉)
            for run in paragraph.runs:
                set_font_style(run, size=11)

def replace_placeholder_with_image(doc, placeholder, image_stream):
    """
    找到 {img_X} 並替換成圖片
    """
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if placeholder in paragraph.text:
                        # 1. 清空佔位符文字
                        paragraph.text = "" 
                        # 2. 插入圖片
                        run = paragraph.add_run()
                        if image_stream:
                            # 圖片寬度固定 8cm (配合 A4 兩欄)
                            run.add_picture(image_stream, width=Cm(8.0))
                        return # 找到一個就停，避免重複

# --- 3. 主流程 ---

def generate_fixed_report(template_file, context, photo_data):
    doc = Document(template_file)
    
    # 1. 填入基本資料 (工程名稱等)
    # 將 {key} 轉換為 {value}
    text_replacements = {f"{{{k}}}": v for k, v in context.items()}
    replace_text_content(doc, text_replacements)
    
    # 2. 填入照片與說明 (迴圈處理 1~8)
    for i in range(1, 9): # 假設最多 8 張
        img_key = f"{{img_{i}}}"   # 對應 Word 裡的 {img_1}
        info_key = f"{{info_{i}}}" # 對應 Word 裡的 {info_1}
        
        # 檢查是否有這張照片
        data_idx = i - 1
        if data_idx < len(photo_data):
            # 有資料：填入圖片與文字
            data = photo_data[data_idx]
            
            # (A) 處理圖片
            replace_placeholder_with_image(doc, img_key, compress_image(data['file']))
            
            # (B) 處理文字 (組合成字串)
            # 格式：
            # 照片編號：01          日期：115.02.03
            # 說明：xxx
            # 實測：xxx
            info_text = f"照片編號：{data['no']:02d}　　　　日期：{data['date_str']}\n"
            info_text += f"說明：{data['desc']}\n"
            info_text += f"實測：{data['result']}"
            
            # 使用文字替換功能填入
            replace_text_content(doc, {info_key: info_text})
            
        else:
            # 沒資料：清空佔位符 (留白)
            replace_text_content(doc, {img_key: ""})
            replace_text_content(doc, {info_key: ""})
            
    return doc

# --- 4. Streamlit UI ---

st.set_page_config(page_title="自主檢查表生成器", layout="wide")
st.title("🏗️ 工程自主檢查表 (定位點填空版)")

if 'doc_buffer' not in st.session_state:
    st.session_state['doc_buffer'] = None
if 'doc_name' not in st.session_state:
    st.session_state['doc_name'] = ""

with st.sidebar:
    st.header("1. 上傳樣板")
    st.info("請確認 Word 表格內已預先填好 `{img_1}`...`{img_8}` 及 `{info_1}`...`{info_8}`")
    template_file = st.file_uploader("Word 樣板", type=['docx'])
    
    st.markdown("---")
    st.header("2. 專案資訊")
    with st.form("info"):
        p_name = st.text_input("工程名稱 {project_name}", "衛生福利部防疫中心興建工程")
        p_cont = st.text_input("施工廠商 {contractor}", "豐譽營造股份有限公司")
        p_sub = st.text_input("協力廠商 {sub_contractor}", "川峻工程有限公司")
        p_loc = st.text_input("施作位置 {location}", "北棟 1F")
        p_item = st.text_input("自檢項目 {check_item}", "拆除工程施工自主檢查(精細拆除) #1")
        check_date = st.date_input("檢查日期", datetime.date.today())
        st.form_submit_button("確認")

    roc_year = check_date.year - 1911
    date_str = f"{roc_year}.{check_date.month:02d}.{check_date.day:02d}"

if template_file:
    st.header("3. 照片上傳 (最多 8 張)")
    files = st.file_uploader("選擇照片", type=['jpg','png','jpeg'], accept_multiple_files=True)
    
    photo_data = []
    if files:
        with st.form("photos"):
            cols = st.columns(2)
            # 限制處理最多 8 張，避免錯誤
            process_files = files[:8]
            
            for i, f in enumerate(process_files):
                with cols[i%2]:
                    st.image(f, width=200)
                    no = st.number_input(f"編號", min_value=1, value=i+1, key=f"n{i}")
                    desc = st.text_input(f"說明", value="現場既有雜物整理", key=f"d{i}")
                    res = st.text_input(f"實測", value="現場既有雜物整理", key=f"r{i}")
                    photo_data.append({"file":f, "no":no, "date_str":date_str, "desc":desc, "result":res})
            
            if st.form_submit_button("🚀 生成 Word 報告"):
                ctx = {
                    "project_name": p_name, "contractor": p_cont, 
                    "sub_contractor": p_sub, "location": p_loc, 
                    "date": date_str, "check_item": p_item
                }
                try:
                    doc = generate_fixed_report(template_file, ctx, photo_data)
                    bio = io.BytesIO()
                    doc.save(bio)
                    st.session_state['doc_buffer'] = bio.getvalue()
                    st.session_state['doc_name'] = f"{date_str}_{p_loc}_檢查表.docx"
                    st.success("✅ 生成成功！")
                except Exception as e:
                    st.error(f"錯誤: {e}")

        if st.session_state['doc_buffer']:
            st.download_button("📥 下載 Word 檔", st.session_state['doc_buffer'], st.session_state['doc_name'], "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
