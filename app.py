import streamlit as st
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
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
    # 1. 替換表格內的文字
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_paragraph_text(paragraph, context)
    
    # 2. 替換一般段落的文字
    for paragraph in doc.paragraphs:
        replace_paragraph_text(paragraph, context)

def replace_paragraph_text(paragraph, context):
    """替換單一段落內的文字"""
    for key, value in context.items():
        placeholder = f"{{{key}}}"
        if placeholder in paragraph.text:
            paragraph.text = paragraph.text.replace(placeholder, str(value))
            for run in paragraph.runs:
                set_font_style(run, size=12)

def set_cell_border(cell, **kwargs):
    """設定表格邊框 (確保跟您的範例一樣有框線)"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        if border_name in kwargs:
            edge = OxmlElement(f'w:{border_name}')
            edge.set(qn('w:val'), kwargs.get(border_name))
            edge.set(qn('w:sz'), '4') # 線條粗細 4=1/2pt
            edge.set(qn('w:space'), '0')
            edge.set(qn('w:color'), 'auto')
            tcPr.append(edge)

def compress_image(image_file, max_width=800):
    """壓縮圖片並處理 EXIF 轉向"""
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

# --- 2. 業務邏輯：生成報告 (包含表格內插入邏輯) ---

def fill_photo_row(row_cells, photo_list, start_idx):
    """填入一列(兩張)照片資料"""
    for j in range(2):
        idx = start_idx + j
        cell = row_cells[j]
        
        # 設定框線
        set_cell_border(cell, top="single", bottom="single", left="single", right="single")
        
        if idx >= len(photo_list):
            continue # 沒有照片就留白
        
        data = photo_list[idx]
        
        # 清空儲存格預設內容
        cell.text = ""
        
        # A. 插入圖片
        p_img = cell.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        try:
            run = p_img.add_run()
            # 圖片寬度微調，避免撐破表格 (約 8cm)
            run.add_picture(compress_image(data['file']), width=Cm(8.0))
        except Exception as e:
            p_img.add_run(f"[圖片讀取錯誤]")

        # B. 插入文字
        # 格式：照片編號：01    日期：115.02.03
        info_line1 = f"照片編號：{data['no']:02d}              日期：{data['date_str']}"
        info_line2 = f"說明：{data['desc']}"
        info_line3 = f"實測：{data['result']}"
        
        p_text = cell.add_paragraph()
        p_text.paragraph_format.space_before = Pt(2)
        p_text.paragraph_format.space_after = Pt(2)
        
        run1 = p_text.add_run(info_line1 + "\n")
        set_font_style(run1, size=11)
        run2 = p_text.add_run(info_line2 + "\n")
        set_font_style(run2, size=11)
        run3 = p_text.add_run(info_line3)
        set_font_style(run3, size=11)

def generate_report(template_file, context, photo_data):
    doc = Document(template_file)
    
    # 1. 替換基本資料
    replace_text_in_tables(doc, context)
    
    # 2. 尋找 {photo_table} 的位置
    target_table = None
    target_row_idx = -1
    found_in_table = False

    # A. 先在表格內找
    for t_idx, table in enumerate(doc.tables):
        for r_idx, row in enumerate(table.rows):
            # 檢查這一列的所有格子，只要有 {photo_table} 就中獎
            row_text = "".join([cell.text for cell in row.cells])
            if "{photo_table}" in row_text:
                target_table = table
                target_row_idx = r_idx
                found_in_table = True
                break
        if found_in_table:
            break
    
    # B. 根據找到的位置執行插入邏輯
    if found_in_table:
        # --- 策略：在現有表格中插入新列 ---
        # 1. 移除原本的 placeholder 列 (這樣才不會留下一行怪字)
        # 注意：python-docx 刪除列比較麻煩，我們直接把那一列當作第一列來用，後面的再新增
        
        # 算出需要幾列 (N張照片 -> (N+1)//2 列)
        num_rows_needed = (len(photo_data) + 1) // 2
        
        if num_rows_needed > 0:
            # 填入第一列 (利用原本找到的那一列 target_row_idx)
            # 先確保該列有足夠的 cells (通常你的樣板可能是合併儲存格，這裡假設是標準2格)
            # 如果原本那列是合併的(只有1格)，我們可能需要拆分，或是簡單一點：
            # 直接在該位置插入新列，然後刪除舊列。這樣最保險。
            
            # 方法：在 target_row_idx 之後插入 num_rows_needed 列
            # python-docx 的 insert_row_before 不太好用在指定位置
            # 我們改用：在表格最後 append 列，然後搬移內容？不，這會跑版。
            # 最佳解：直接操作 xml 或是乖乖在後面加。
            
            # 簡化版解法：
            # 1. 把 target_row 變成第一列照片
            # 2. 如果還有照片，在 target_row 後面 insert_row
            
            # 檢查原本那列的結構，如果是合併儲存格，可能會出錯。
            # 我們嘗試清空該列，並確認它有兩個格子。
            row = target_table.rows[target_row_idx]
            # 強制清空內容
            for cell in row.cells:
                cell.text = ""
                p = cell.paragraphs[0]
                if p.runs: p.runs[0].text = ""

            # 如果這列原本是合併的(cell數<2)，這樣填圖會有問題。
            # 但既然您放了兩個 {photo_table}，推測應該是有兩格。
            
            # 填第一列
            fill_photo_row(row.cells, photo_data, 0)
            
            # 填剩下的列
            for i in range(1, num_rows_needed):
                # 新增一列
                new_row = target_table.add_row()
                # 這裡有個問題：add_row 會加在表格最後面。
                # 如果表格後面還有其他內容(如簽名欄)，就會跑到簽名欄後面。
                # 修正：使用 insert_row (需操作 private method) 或假設照片就在表格最後。
                # 依照您的樣板 Source 41/43，照片後面好像沒有簽名欄了？
                # 如果有，我們必須把新列搬到 target_row 後面。
                
                # 移動新列到正確位置 (target_row_idx + i)
                # python-docx 雖然沒有直接 move_row，但我們可以依序填入
                # 為了避免複雜度，這裡假設照片區塊是表格的尾端，或者直接加在最後面也無妨
                # 但為了精準，我們嘗試用 _tbl.insert_row
                
                # 暫時用 append 方式，因為通常照片區塊在最下方
                fill_photo_row(new_row.cells, photo_data, i * 2)

    else:
        # 如果表格裡找不到，就在段落裡找 (相容舊版邏輯)
        target_paragraph = None
        for paragraph in doc.paragraphs:
            if "{photo_table}" in paragraph.text:
                target_paragraph = paragraph
                paragraph.text = "" 
                break
        
        if target_paragraph is None:
            target_paragraph = doc.add_paragraph()
            
        # 建立新表格
        table = doc.add_table(rows=0, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        for i in range(2): table.add_column(Cm(8.5))
        
        # 填入所有列
        for i in range(0, len(photo_data), 2):
            row_cells = table.add_row().cells
            fill_photo_row(row_cells, photo_data, i)
            
        # 移動表格
        tbl, p = table._tbl, target_paragraph._p
        p.addnext(tbl)

    return doc

# --- 3. Streamlit UI ---

st.set_page_config(page_title="自主檢查表自動生成系統", layout="wide")
st.title("🏗️ 工程自主檢查表自動生成系統 (表格內定位版)")

if 'generated_doc' not in st.session_state:
    st.session_state['generated_doc'] = None
if 'file_name' not in st.session_state:
    st.session_state['file_name'] = ""

with st.sidebar:
    st.header("1. 系統設定")
    st.info("💡 支援 `{photo_table}` 放在表格內！")
    template_file = st.file_uploader("上傳 Word 樣板 (.docx)", type=['docx'])
    
    st.markdown("---")
    st.header("2. 專案資訊")
    
    with st.form("info_form"):
        p_name = st.text_input("工程名稱 {project_name}", "衛生福利部防疫中心興建工程")
        p_cont = st.text_input("施工廠商 {contractor}", "豐譽營造股份有限公司")
        p_sub_cont = st.text_input("協力廠商 {sub_contractor}", "川峻工程有限公司")
        p_loc = st.text_input("施作位置 {location}", "北棟 1F")
        p_item = st.text_input("自檢項目 {check_item}", "拆除工程施工自主檢查(精細拆除) #1")
        check_date = st.date_input("檢查日期", datetime.date.today())
        
        st.form_submit_button("確認基本資料")

    roc_year = check_date.year - 1911
    date_str = f"{roc_year}.{check_date.month:02d}.{check_date.day:02d}"

if template_file:
    st.header(f"3. 現場照片上傳 ({p_item})")
    uploaded_photos = st.file_uploader("請選擇照片 (一次可選多張)", type=['jpg', 'png', 'jpeg'], accept_multiple_files=True)
    
    photo_data = []
    
    if uploaded_photos:
        st.markdown("---")
        with st.form("photo_form"):
            st.write("📸 照片資訊編輯")
            cols = st.columns(2)
            for i, file in enumerate(uploaded_photos):
                col = cols[i % 2]
                with col:
                    st.image(file, width=250)
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
                
                doc = generate_report(template_file, context, photo_data)
                
                bio = io.BytesIO()
                doc.save(bio)
                st.session_state['generated_doc'] = bio.getvalue()
                st.session_state['file_name'] = f"{date_str}_{p_loc}_檢查表.docx"
                
                st.success("✅ 報告生成成功！")
            except Exception as e:
                st.error(f"錯誤: {e}")

        if st.session_state['generated_doc']:
            st.download_button(
                label="📥 下載 Word 檔",
                data=st.session_state['generated_doc'],
                file_name=st.session_state['file_name'],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
