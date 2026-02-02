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
    """設定中英文字型 (解決 Word 中文顯示問題)"""
    run.font.name = 'Times New Roman'  # 英數預設
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name) # 中文強制設定
    run.font.size = Pt(size)
    run.bold = bold

def replace_text_in_tables(doc, context):
    """在 Word 表格中尋找 {keywords} 並替換"""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in context.items():
                        placeholder = f"{{{key}}}"
                        if placeholder in paragraph.text:
                            # 簡單替換
                            paragraph.text = paragraph.text.replace(placeholder, str(value))
                            # 重新套用字型
                            for run in paragraph.runs:
                                set_font_style(run, size=12)

def set_cell_border(cell, **kwargs):
    """設定儲存格邊框"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        if border_name in kwargs:
            edge = OxmlElement(f'w:{border_name}')
            edge.set(qn('w:val'), kwargs.get(border_name))
            edge.set(qn('w:sz'), '4') # 線條粗細
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

# --- 2. 業務邏輯：生成照片表格 ---

def add_photo_table(doc, photo_data):
    """依照工程慣例 (2欄xN列) 插入照片表格"""
    # 建立表格：2 欄
    table = doc.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    # 設定欄寬 (假設 A4 頁面，每欄約 8.5cm)
    for i in range(2):
        table.add_column(Cm(8.5))

    # 遍歷照片資料 (每 2 張一列)
    for i in range(0, len(photo_data), 2):
        row_cells = table.add_row().cells
        
        for j in range(2):
            idx = i + j
            if idx >= len(photo_data):
                break # 如果照片是奇數張，跳出
            
            cell = row_cells[j]
            data = photo_data[idx]
            
            # (1) 插入圖片
            p_img = cell.paragraphs[0]
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                run = p_img.add_run()
                # 這裡限制寬度 8.0cm 確保不會撐爆表格
                run.add_picture(compress_image(data['file']), width=Cm(8.0))
            except Exception as e:
                p_img.add_run(f"[圖片錯誤: {e}]")

            # (2) 插入文字資訊
            info_text = f"照片編號：{data['no']:02d}              日期：{data['date_str']}\n"
            info_text += f"說明：{data['desc']}\n"
            info_text += f"實測：{data['result']}"
            
            p_text = cell.add_paragraph(info_text)
            p_text.paragraph_format.space_before = Pt(4)
            p_text.paragraph_format.space_after = Pt(8)
            
            # 設定文字樣式
            for run in p_text.runs:
                set_font_style(run, size=11)
            
            # 設定邊框 (Single 線條)
            set_cell_border(cell, top="single", bottom="single", left="single", right="single")

# --- 3. Streamlit 使用者介面 ---

st.set_page_config(page_title="自主檢查表自動生成系統", layout="wide")
st.title("🏗️ 工程自主檢查表自動生成系統 (Template 版)")

# 初始化 session state 用來存檔
if 'generated_doc' not in st.session_state:
    st.session_state['generated_doc'] = None
if 'file_name' not in st.session_state:
    st.session_state['file_name'] = ""

# --- 側邊欄：設定與樣板上傳 ---
with st.sidebar:
    st.header("1. 系統設定")
    
    st.info("💡 請上傳您的 Word 底稿 (.docx)")
    template_file = st.file_uploader("上傳 Word 樣板", type=['docx'])
    
    st.markdown("---")
    st.header("2. 專案資訊輸入")
    
    # 使用 Form 來避免輸入一個字就重新整理
    with st.form("project_info_form"):
        # 這裡對應 Word 裡的 {keyword}
        p_name = st.text_input("工程名稱 {project_name}", "衛生福利部防疫中心興建工程")
        p_cont = st.text_input("施工廠商 {contractor}", "豐譽營造股份有限公司")
        # --- 新增：協力廠商 ---
        p_sub_cont = st.text_input("協力廠商 {sub_contractor}", "川峻工程有限公司") 
        
        p_loc = st.text_input("施作位置 {location}", "北棟 1F")
        p_item = st.text_input("自檢項目 {check_item}", "拆除工程施工自主檢查(精細拆除) #1")
        
        # 日期處理
        check_date = st.date_input("檢查日期", datetime.date.today())
        
        st.form_submit_button("確認基本資料") # 這按鈕只是為了讓 Form 運作，主要觸發在下方

    # 預先計算民國年日期字串
    roc_year = check_date.year - 1911
    date_str = f"{roc_year}.{check_date.month:02d}.{check_date.day:02d}"

# --- 主畫面：照片處理 ---
if template_file:
    st.header(f"3. 現場照片上傳 ({p_item})")
    st.markdown("💡 您可以一次選取 **8 張** (或更多) 照片上傳，系統會自動依序編號 1-8。")
    
    uploaded_photos = st.file_uploader("請選擇照片", type=['jpg', 'png', 'jpeg'], accept_multiple_files=True)
    
    photo_data = []
    
    if uploaded_photos:
        st.markdown("---")
        # 照片編輯表單
        with st.form("photo_form"):
            st.write("📸 照片資訊快速編輯")
            
            # 使用 Grid 排版，每列 2 張，方便檢視
            cols = st.columns(2)
            
            for i, file in enumerate(uploaded_photos):
                col = cols[i % 2] # 決定左邊還是右邊
                with col:
                    st.image(file, width=300)
                    
                    # 自動計算編號：1, 2, 3... 8
                    current_no = i + 1 
                    
                    c1, c2 = st.columns([1, 2])
                    # 讓使用者可以改編號，但預設就是 1,2,3...8
                    no = c1.number_input(f"編號", min_value=1, value=current_no, key=f"n{i}")
                    
                    # 預設文字邏輯 (可選)
                    default_desc = "現場既有雜物整理"
                    default_res = "現場既有雜物整理"
                    
                    desc = c2.text_input(f"說明", value=default_desc, key=f"d{i}")
                    res = st.text_input(f"實測", value=default_res, key=f"r{i}")
                    
                    photo_data.append({
                        "file": file,
                        "no": no,
                        "date_str": date_str,
                        "desc": desc,
                        "result": res
                    })
                    st.markdown("---")
            
            # Form 提交按鈕
            generate_clicked = st.form_submit_button("🚀 生成 Word 報告")

        # --- 處理邏輯 (在 Form 之外處理下載按鈕) ---
        if generate_clicked:
            try:
                # 1. 讀取樣板
                doc = Document(template_file)
                
                # 2. 準備替換的資料 (包含新增的協力廠商)
                context = {
                    "project_name": p_name,
                    "contractor": p_cont,
                    "sub_contractor": p_sub_cont, # 新增
                    "location": p_loc,
                    "date": date_str,
                    "check_item": p_item
                }
                
                # 3. 執行替換
                replace_text_in_tables(doc, context)
                
                # 4. 插入照片表格 (8張照片會自動產生4列)
                # 加標題
                p = doc.add_paragraph()
                run = p.add_run("檢 查 照 片")
                set_font_style(run, size=14, bold=True)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                add_photo_table(doc, photo_data)
                
                # 5. 存入 Session State
                bio = io.BytesIO()
                doc.save(bio)
                st.session_state['generated_doc'] = bio.getvalue()
                st.session_state['file_name'] = f"{date_str}_{p_loc}_{p_item}_自主檢查表.docx"
                
                st.success("✅ 報告生成完畢！請點擊下方按鈕下載。")
            
            except Exception as e:
                st.error(f"生成失敗: {e}")

        # --- 下載按鈕 (獨立於 Form 之外) ---
        if st.session_state['generated_doc'] is not None:
            st.download_button(
                label="📥 下載 Word 檔",
                data=st.session_state['generated_doc'],
                file_name=st.session_state['file_name'],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

else:
    st.info("👈 請先在左側上傳 Word 樣板 (.docx) 以開始使用。")
