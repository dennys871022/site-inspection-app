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

# --- 1. 核心工具函數 (專家級設定) ---

def set_font_style(run, font_name='標楷體', size=12, bold=False):
    """設定中英文字型 (解決 Word 中文顯示問題)"""
    run.font.name = 'Times New Roman'  # 英數預設
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name) # 中文強制設定
    run.font.size = Pt(size)
    run.bold = bold

def replace_text_in_tables(doc, context):
    """
    在 Word 表格中尋找 {keywords} 並替換成使用者輸入的資料。
    這是達成「格式一模一樣」的關鍵：直接改原本的字，不動表格結構。
    """
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in context.items():
                        placeholder = f"{{{key}}}"  # 例如 {project_name}
                        if placeholder in paragraph.text:
                            # 簡單替換 (保留原本段落格式)
                            paragraph.text = paragraph.text.replace(placeholder, str(value))
                            # 重新套用字型 (因為替換後可能會跑掉)
                            for run in paragraph.runs:
                                set_font_style(run, size=12)

def set_cell_border(cell, **kwargs):
    """
    (進階) 使用 OXML 設定儲存格邊框，確保畫出來的表格跟原檔一樣有格線。
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    
    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        if border_name in kwargs:
            edge = OxmlElement(f'w:{border_name}')
            edge.set(qn('w:val'), kwargs.get(border_name)) # single, double, nil
            edge.set(qn('w:sz'), '4') # 線條粗細
            edge.set(qn('w:space'), '0')
            edge.set(qn('w:color'), 'auto')
            tcPr.append(edge)

def compress_image(image_file, max_width=800):
    """壓縮圖片，避免 Word 檔過大"""
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
    """
    在文件末尾依照工程慣例 (2欄xN列) 插入照片表格。
    格式模仿：[照片] -> [編號/日期] -> [說明]
    """
    # 新增分頁 (如果需要)
    # doc.add_page_break() 
    
    # 建立表格：2 欄 (依照你的範例照片，通常是一排兩張)
    table = doc.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    # 設定欄寬 (假設 A4 寬度扣掉邊界，每欄約 8.5cm)
    for i in range(2):
        table.add_column(Cm(8.5)) # 這行可能需要根據 python-docx 版本微調，通常是用 cell.width

    # 遍歷照片資料
    for i in range(0, len(photo_data), 2):
        row_cells = table.add_row().cells
        
        # 處理這一列的 1~2 張照片
        for j in range(2):
            idx = i + j
            if idx >= len(photo_data):
                break
            
            cell = row_cells[j]
            data = photo_data[idx]
            
            # (1) 插入圖片段落
            p_img = cell.paragraphs[0]
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                run = p_img.add_run()
                run.add_picture(compress_image(data['file']), width=Cm(8.0))
            except Exception as e:
                p_img.add_run(f"[圖片錯誤: {e}]")

            # (2) 插入文字資訊段落
            # 格式參考：照片編號：01  日期：114.11.26
            info_text = f"照片編號：{data['no']:02d}    日期：{data['date_str']}\n"
            info_text += f"說明：{data['desc']}\n"
            info_text += f"實測：{data['result']}"
            
            p_text = cell.add_paragraph(info_text)
            p_text.paragraph_format.space_before = Pt(4)
            p_text.paragraph_format.space_after = Pt(8)
            
            # 設定文字樣式
            for run in p_text.runs:
                set_font_style(run, size=11)
            
            # 設定邊框 (讓它看起來像正式表格)
            set_cell_border(cell, top="single", bottom="single", left="single", right="single")

# --- 3. Streamlit 使用者介面 ---

st.set_page_config(page_title="自主檢查表自動生成系統", layout="wide")
st.title("🏗️ 工程自主檢查表自動生成系統 (Template 版)")

# --- 側邊欄：設定與樣板上傳 ---
with st.sidebar:
    st.header("1. 系統設定")
    
    # A. 樣板上傳區 (關鍵功能)
    st.info("💡 為了確保格式「一模一樣」，請上傳你的 Word 底稿。")
    template_file = st.file_uploader("上傳 Word 樣板 (.docx)", type=['docx'])
    
    if not template_file:
        st.warning("⚠️ 請先上傳樣板文件以開始使用。")
        st.markdown("""
        **如何製作樣板？**
        打開你的 Word 檔，把要替換的地方改成：
        - `{project_name}` (工程名稱)
        - `{contractor}` (施工廠商)
        - `{location}` (施作位置)
        - `{date}` (日期)
        - `{check_item}` (自檢項目)
        """)
    
    st.markdown("---")
    st.header("2. 專案資訊輸入")
    # 這裡的 key 要對應 Word 樣板裡的 {key}
    p_name = st.text_input("工程名稱 {project_name}", "衛生福利部防疫中心興建工程")
    p_cont = st.text_input("施工廠商 {contractor}", "豐譽營造股份有限公司")
    p_loc = st.text_input("施作位置 {location}", "北棟")
    
    # 日期處理 (轉民國年)
    check_date = st.date_input("檢查日期")
    roc_year = check_date.year - 1911
    date_str = f"{roc_year}.{check_date.month:02d}.{check_date.day:02d}"
    st.text(f"日期預覽：{date_str}")
    
    p_item = st.text_input("自檢項目 {check_item}", "拆除工程施工自主檢查")
    p_content = st.text_area("檢查內容 (選填)", "1. 防塵作為\n2. 保留構造不得損傷")

# --- 主畫面：照片處理 ---
if template_file:
    st.header("3. 現場照片上傳")
    uploaded_photos = st.file_uploader("上傳照片", type=['jpg', 'png', 'jpeg'], accept_multiple_files=True)
    
    photo_data = []
    
    if uploaded_photos:
        with st.form("photo_form"):
            st.write("照片資訊編輯")
            cols = st.columns(2)
            for i, file in enumerate(uploaded_photos):
                col = cols[i % 2]
                with col:
                    st.image(file, width=200)
                    c1, c2 = st.columns([1, 2])
                    no = c1.number_input(f"編號", min_value=1, value=i+1, key=f"n{i}")
                    desc = c2.text_input(f"說明", value="依施工計畫執行", key=f"d{i}")
                    res = st.text_input(f"實測", value="與計畫相符", key=f"r{i}")
                    
                    photo_data.append({
                        "file": file,
                        "no": no,
                        "date_str": date_str, # 使用上面算好的日期
                        "desc": desc,
                        "result": res
                    })
                    st.markdown("---")
            
            submit = st.form_submit_button("🚀 生成 Word 報告")
            
            if submit:
                # 1. 讀取使用者上傳的樣板
                doc = Document(template_file)
                
                # 2. 準備要替換的資料字典
                context = {
                    "project_name": p_name,
                    "contractor": p_cont,
                    "location": p_loc,
                    "date": date_str,
                    "check_item": p_item
                    # 如果樣板有 {check_content} 也可以替換
                }
                
                # 3. 執行文字替換 (保留原格式)
                replace_text_in_tables(doc, context)
                
                # 4. 在文件最後加入照片表格
                # 先加一個分頁符號，讓照片從新的一頁開始 (可選)
                # doc.add_page_break() 
                # 加標題
                p = doc.add_paragraph()
                run = p.add_run("檢 查 照 片")
                set_font_style(run, size=14, bold=True)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # 插入照片表格
                add_photo_table(doc, photo_data)
                
                # 5. 輸出檔案
                bio = io.BytesIO()
                doc.save(bio)
                
                out_name = f"{date_str}_{p_loc}_自主檢查表.docx"
                
                st.success("✅ 報告生成完畢！")
                st.download_button(
                    label="📥 下載 Word 檔",
                    data=bio.getvalue(),
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
