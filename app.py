import streamlit as st
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image
import io
import datetime

# --- 1. 核心工具函數 ---

def set_font_style(run, font_name='標楷體', size=12, bold=False):
    """設定中英文字型 (Times New Roman + 標楷體)"""
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.bold = bold

def replace_text_in_tables(doc, context):
    """替換全文件(含表格)內的文字變數"""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_paragraph_text(paragraph, context)
    for paragraph in doc.paragraphs:
        replace_paragraph_text(paragraph, context)

def replace_paragraph_text(paragraph, context):
    for key, value in context.items():
        placeholder = f"{{{key}}}"
        if placeholder in paragraph.text:
            paragraph.text = paragraph.text.replace(placeholder, str(value))
            for run in paragraph.runs:
                set_font_style(run, size=12)

def set_cell_border(cell, top=None, bottom=None, left=None, right=None, insideH=None, insideV=None):
    """強制設定儲存格邊框"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for border_name, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        if val:
            edge = OxmlElement(f'w:{border_name}')
            edge.set(qn('w:val'), val)
            edge.set(qn('w:sz'), '4')
            edge.set(qn('w:space'), '0')
            edge.set(qn('w:color'), 'auto')
            tcPr.append(edge)

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

# --- 2. 關鍵修復：表格列增生邏輯 ---

def process_photo_table(doc, photo_data):
    """找到含有 {photo_table} 的表格列，並在該處增生照片列"""
    target_table = None
    target_row_index = -1
    
    # 1. 尋找定位點
    for table in doc.tables:
        for i, row in enumerate(table.rows):
            # 檢查整列文字
            row_text = "".join([c.text for c in row.cells])
            if "{photo_table}" in row_text:
                target_table = table
                target_row_index = i
                break
        if target_table:
            break
            
    if not target_table:
        st.warning("⚠️ 找不到 {photo_table} 定位點，請檢查 Word 樣板。")
        return 
        
    # 2. 計算需要的總列數
    total_photos = len(photo_data)
    rows_needed = (total_photos + 1) // 2
    
    # 3. 準備第一列 (清除原本的定位字)
    first_row = target_table.rows[target_row_index]
    for cell in first_row.cells:
        cell.text = ""
        for p in cell.paragraphs: p.text = ""

    # 4. 開始填入照片
    for r in range(rows_needed):
        # 決定要填入哪一列
        if r == 0:
            current_row = first_row
        else:
            # 在表格最後新增一列 (會繼承表格寬度)
            current_row = target_table.add_row()
        
        start_photo_idx = r * 2
        
        for col in range(2): # 左右兩欄
            photo_idx = start_photo_idx + col
            
            # 防呆：確保格子存在
            if col >= len(current_row.cells): continue
                
            cell = current_row.cells[col]
            set_cell_border(cell, top="single", bottom="single", left="single", right="single")
            
            if photo_idx >= total_photos: continue 
                
            data = photo_data[photo_idx]
            
            # --- 內容填寫區 (這裡控制排版) ---
            
            # A. 圖片
            p_img = cell.paragraphs[0]
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                run = p_img.add_run()
                # 這裡設定圖片寬度，約 8.5cm 適合 A4 兩欄
                run.add_picture(compress_image(data['file']), width=Cm(8.5))
            except:
                p_img.add_run("[圖片錯誤]")
            
            # B. 文字 (模仿你的範例格式)
            p_info = cell.add_paragraph()
            p_info.paragraph_format.space_before = Pt(4)
            p_info.paragraph_format.space_after = Pt(2)
            
            # 第一行: 照片編號 + 日期 (中間用全形空白調整間距)
            # 你的範例：照片編號：01              日期：115.02.03
            text_line1 = f"照片編號：{data['no']:02d}　　　　　日期：{data['date_str']}\n"
            run1 = p_info.add_run(text_line1)
            set_font_style(run1, size=11)
            
            # 第二行: 說明
            text_line2 = f"說明：{data['desc']}\n"
            run2 = p_info.add_run(text_line2)
            set_font_style(run2, size=11)
            
            # 第三行: 實測
            text_line3 = f"實測：{data['result']}"
            run3 = p_info.add_run(text_line3)
            set_font_style(run3, size=11)

# --- 3. 主程式邏輯 ---

def generate_report(template_file, context, photo_data):
    doc = Document(template_file)
    replace_text_in_tables(doc, context)
    process_photo_table(doc, photo_data)
    return doc

# --- 4. UI ---

st.set_page_config(page_title="自主檢查表生成器", layout="wide")
st.title("🏗️ 工程自主檢查表自動生成系統 (最終修復版)")

if 'doc_buffer' not in st.session_state:
    st.session_state['doc_buffer'] = None
if 'doc_name' not in st.session_state:
    st.session_state['doc_name'] = ""

with st.sidebar:
    st.header("1. 上傳樣板")
    st.info("請確保 Word 表格內留有一行 `{photo_table}`")
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
    st.header("3. 照片上傳")
    files = st.file_uploader("選擇照片", type=['jpg','png','jpeg'], accept_multiple_files=True)
    
    photo_data = []
    if files:
        with st.form("photos"):
            cols = st.columns(2)
            for i, f in enumerate(files):
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
                    doc = generate_report(template_file, ctx, photo_data)
                    bio = io.BytesIO()
                    doc.save(bio)
                    st.session_state['doc_buffer'] = bio.getvalue()
                    st.session_state['doc_name'] = f"{date_str}_{p_loc}_檢查表.docx"
                    st.success("✅ 生成成功！")
                except Exception as e:
                    st.error(f"錯誤: {e}")

        if st.session_state['doc_buffer']:
            st.download_button("📥 下載 Word 檔", st.session_state['doc_buffer'], st.session_state['doc_name'], "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
