import streamlit as st
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml.ns import qn
from PIL import Image
import io
import datetime

# --- 1. 基礎設定 (樣式繼承核心) ---

def set_font_style(run, font_name='標楷體', size=None):
    """
    設定字型：
    1. 英數使用 Times New Roman
    2. 中文強制使用 標楷體
    3. 字體大小 (Size) 預設為 None -> 代表不修改，直接繼承樣板原本的大小
    """
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    if size:
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
    """通用文字替換：將 {key} 換成 value"""
    # 遍歷所有表格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_paragraph(paragraph, replacements)
    # 遍歷一般段落
    for paragraph in doc.paragraphs:
        replace_paragraph(paragraph, replacements)

def replace_paragraph(paragraph, replacements):
    """
    在段落中尋找並替換文字。
    【關鍵修正】優先嘗試 Run Level 替換，保留原本的字體大小與粗細。
    """
    if not paragraph.text:
        return

    for key, value in replacements.items():
        if key in paragraph.text:
            val_str = str(value) if value is not None else ""
            
            # 策略 A: 嘗試在單一 Run (樣式區塊) 中找到完整關鍵字
            # 這樣可以完美保留該關鍵字原本設定的 粗體/大小/顏色
            replaced_in_run = False
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, val_str)
                    # 這裡 size=None，表示不更動原本的大小
                    set_font_style(run, size=None) 
                    replaced_in_run = True
            
            # 策略 B: 如果關鍵字被 Word 切割(例如跨越格式)，則重寫整個段落
            if not replaced_in_run:
                # 為了避免格式跑掉，我們直接替換文字，並嘗試套用第一個 Run 的格式
                paragraph.text = paragraph.text.replace(key, val_str)
                for run in paragraph.runs:
                    # 同樣 size=None，只確保中文字型顯示正常
                    set_font_style(run, size=None)

def replace_placeholder_with_image(doc, placeholder, image_stream):
    """找到 {img_X} 並替換成圖片"""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if placeholder in paragraph.text:
                        paragraph.text = "" # 清空文字
                        run = paragraph.add_run()
                        if image_stream:
                            # 圖片寬度固定 8cm (適應一般 A4 兩欄寬度)
                            run.add_picture(image_stream, width=Cm(8.0))
                        return 

# --- 3. 主流程 ---

def generate_fixed_report(template_file, context, photo_data):
    doc = Document(template_file)
    
    # 1. 填入基本資料 (工程名稱等)
    # 這裡的 replace_paragraph 已經修正，會保留原本表格裡的字體大小
    text_replacements = {f"{{{k}}}": v for k, v in context.items()}
    replace_text_content(doc, text_replacements)
    
    # 2. 填入照片與說明 (處理 1~8 張)
    for i in range(1, 9):
        img_key = f"{{img_{i}}}"
        info_key = f"{{info_{i}}}"
        
        data_idx = i - 1
        if data_idx < len(photo_data):
            data = photo_data[data_idx]
            
            # (A) 填入圖片
            replace_placeholder_with_image(doc, img_key, compress_image(data['file']))
            
            # (B) 填入文字 (日期往右調整)
            # 因為 replace_paragraph 現在會繼承格式，
            # 所以這裡組出的字串會套用您在 Word 樣板裡對 {info_X} 設定的大小。
            info_text = f"照片編號：{data['no']:02d}　　　　　　日期：{data['date_str']}\n"
            info_text += f"說明：{data['desc']}\n"
            info_text += f"實測：{data['result']}"
            
            # 這裡傳入 replacements 字典
            replace_text_content(doc, {info_key: info_text})
            
        else:
            # 沒資料則清空佔位符
            replace_text_content(doc, {img_key: ""})
            replace_text_content(doc, {info_key: ""})
            
    return doc

# --- 4. Streamlit UI ---

st.set_page_config(page_title="自主檢查表生成器", layout="wide")
st.title("🏗️ 工程自主檢查表 (樣式完美繼承版)")

if 'doc_buffer' not in st.session_state:
    st.session_state['doc_buffer'] = None
if 'doc_name' not in st.session_state:
    st.session_state['doc_name'] = ""

with st.sidebar:
    st.header("1. 上傳樣板")
    st.info("💡 Word 樣板中的 `{project_name}` 或 `{info_1}` 設為幾號字，產出就是幾號字。")
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
            # 限制最多處理 8 張
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
