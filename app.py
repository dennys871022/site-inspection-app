import streamlit as st
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image
import io
import datetime

# --- 1. 核心工具函數 ---

def set_font_style(run, font_name='標楷體', size=12, bold=False):
    """設定中英文字型"""
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.bold = bold

def replace_text_in_tables(doc, context):
    """替換表格內的文字變數"""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in context.items():
                        placeholder = f"{{{key}}}"
                        if placeholder in paragraph.text:
                            paragraph.text = paragraph.text.replace(placeholder, str(value))
                            for run in paragraph.runs:
                                set_font_style(run, size=12)

def set_cell_border(cell, **kwargs):
    """設定表格邊框"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        if border_name in kwargs:
            edge = OxmlElement(f'w:{border_name}')
            edge.set(qn('w:val'), kwargs.get(border_name))
            edge.set(qn('w:sz'), '4')
            edge.set(qn('w:space'), '0')
            edge.set(qn('w:color'), 'auto')
            tcPr.append(edge)

def compress_image(image_file, max_width=800):
    """壓縮圖片"""
    img = Image.open(image_file)
    if img.mode == 'RGBA':
        img = img.convert('RGB')
    ratio = max_width / float(img.size[0])
    if ratio < 1:
        h_size = int((float(img.size[1]) * float(ratio)))
        img = img.resize((max_width, h_size), Image.Resampling.LANCZOS)
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='JPEG', quality=70)
    img_byte_arr.seek(0)
    return img_byte_arr

def move_table_after(table, paragraph):
    """
    【關鍵技術】將新建立的表格 (table) 移動到指定段落 (paragraph) 的後面。
    這樣才能精準控制表格位置，不會永遠跑到文件最後面。
    """
    tbl, p = table._tbl, paragraph._p
    p.addnext(tbl)

# --- 2. 業務邏輯：在指定位置生成照片表格 ---

def generate_report(template_file, context, photo_data):
    doc = Document(template_file)
    
    # 1. 替換基本資料 (Project Info)
    replace_text_in_tables(doc, context)
    
    # 2. 尋找定位點 {photo_table} 並插入照片表格
    target_paragraph = None
    
    # 搜尋所有段落尋找定位點
    for paragraph in doc.paragraphs:
        if "{photo_table}" in paragraph.text:
            target_paragraph = paragraph
            paragraph.text = "" # 清空定位點文字，只留位置
            break
            
    # 如果找不到定位點，就預設加在最後面
    if target_paragraph is None:
        # 如果沒找到，加一個新段落當作目標
        target_paragraph = doc.add_paragraph() 
    
    # 3. 建立照片表格 (暫時建立在記憶體中，等下移動)
    table = doc.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    # 設定欄寬
    for i in range(2):
        table.add_column(Cm(8.5))

    # 填入照片資料 (支援 8 張或更多)
    for i in range(0, len(photo_data), 2):
        row_cells = table.add_row().cells
        
        for j in range(2):
            idx = i + j
            if idx >= len(photo_data):
                break
            
            cell = row_cells[j]
            data = photo_data[idx]
            
            # 圖片
            p_img = cell.paragraphs[0]
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                run = p_img.add_run()
                run.add_picture(compress_image(data['file']), width=Cm(8.0))
            except:
                p_img.add_run("[圖片錯誤]")

            # 文字
            info_text = f"照片編號：{data['no']:02d}              日期：{data['date_str']}\n"
            info_text += f"說明：{data['desc']}\n"
            info_text += f"實測：{data['result']}"
            
            p_text = cell.add_paragraph(info_text)
            p_text.paragraph_format.space_before = Pt(4)
            p_text.paragraph_format.space_after = Pt(8)
            for run in p_text.runs:
                set_font_style(run, size=11)
            
            set_cell_border(cell, top="single", bottom="single", left="single", right="single")
    
    # 【關鍵步驟】將做好的表格搬移到定位點後面
    move_table_after(table, target_paragraph)
    
    return doc

# --- 3. Streamlit UI ---

st.set_page_config(page_title="自主檢查表自動生成系統", layout="wide")
st.title("🏗️ 工程自主檢查表自動生成系統 (定位點版)")

# 初始化
if 'generated_doc' not in st.session_state:
    st.session_state['generated_doc'] = None
if 'file_name' not in st.session_state:
    st.session_state['file_name'] = ""

with st.sidebar:
    st.header("1. 系統設定")
    st.info("💡 請上傳 Word 樣板，並確保裡面有 `{photo_table}` 定位字串。")
    template_file = st.file_uploader("上傳 Word 樣板", type=['docx'])
    
    st.markdown("---")
    st.header("2. 專案資訊")
    
    with st.form("info_form"):
        p_name = st.text_input("工程名稱 {project_name}", "衛生福利部防疫中心興建工程")
        p_cont = st.text_input("施工廠商 {contractor}", "豐譽營造股份有限公司")
        p_sub_cont = st.text_input("協力廠商 {sub_contractor}", "川峻工程有限公司")
        p_loc = st.text_input("施作位置 {location}", "北棟 1F")
        p_item = st.text_input("自檢項目 {check_item}", "拆除工程施工自主檢查(精細拆除) #1")
        check_date = st.date_input("檢查日期", datetime.date.today())
        st.form_submit_button("確認資訊")

    roc_year = check_date.year - 1911
    date_str = f"{roc_year}.{check_date.month:02d}.{check_date.day:02d}"

if template_file:
    st.header(f"3. 現場照片上傳 ({p_item})")
    st.markdown("💡 系統支援 **8 張** (或更多) 照片，請一次選取上傳，系統會自動排版。")
    
    uploaded_photos = st.file_uploader("請選擇照片", type=['jpg', 'png', 'jpeg'], accept_multiple_files=True)
    
    photo_data = []
    
    if uploaded_photos:
        st.markdown("---")
        with st.form("photo_form"):
            st.write("📸 照片資訊編輯")
            cols = st.columns(2)
            for i, file in enumerate(uploaded_photos):
                col = cols[i % 2]
                with col:
                    st.image(file, width=300)
                    # 自動編號 1-8
                    no = st.number_input(f"編號", min_value=1, value=i+1, key=f"n{i}")
                    desc = st.text_input(f"說明", value="現場既有雜物整理", key=f"d{i}")
                    res = st.text_input(f"實測", value="現場既有雜物整理", key=f"r{i}")
                    
                    photo_data.append({
                        "file": file, "no": no, "date_str": date_str, "desc": desc, "result": res
                    })
            
            generate_clicked = st.form_submit_button("🚀 生成 Word 報告")

        if generate_clicked:
            try:
                context = {
                    "project_name": p_name,
                    "contractor": p_cont,
                    "sub_contractor": p_sub_cont,
                    "location": p_loc,
                    "date": date_str,
                    "check_item": p_item
                }
                
                # 呼叫生成函數
                doc = generate_report(template_file, context, photo_data)
                
                bio = io.BytesIO()
                doc.save(bio)
                st.session_state['generated_doc'] = bio.getvalue()
                st.session_state['file_name'] = f"{date_str}_{p_loc}_檢查表.docx"
                
                st.success("✅ 報告生成成功！請下載。")
            except Exception as e:
                st.error(f"錯誤: {e}")

        if st.session_state['generated_doc']:
            st.download_button(
                label="📥 下載 Word 檔",
                data=st.session_state['generated_doc'],
                file_name=st.session_state['file_name'],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
