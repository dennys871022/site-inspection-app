import streamlit as st
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml.ns import qn
from PIL import Image
import io
import datetime

# --- 1. 核心工具：只設定字型家族，不改大小粗細 ---

def ensure_chinese_font(run):
    """
    僅設定中文字型為標楷體，英文字型為 Times New Roman。
    絕不修改字體大小 (Size) 或粗體 (Bold)，完全繼承樣板設定。
    """
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

def compress_image(image_file, max_width=800):
    """圖片處理：壓縮與轉向"""
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

# --- 2. 替換邏輯：原地替換，保留格式 ---

def replace_text_in_paragraph(paragraph, replacements):
    """
    在段落中進行文字替換。
    優先嘗試保留 Run 的格式。
    """
    if not paragraph.text:
        return

    for key, value in replacements.items():
        if key in paragraph.text:
            value = str(value) if value is not None else ""
            
            # 策略 A: 嘗試在單一 Run 中找到完整關鍵字 (最能保留格式)
            replaced = False
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, value)
                    ensure_chinese_font(run) # 只確保中文顯示正常
                    replaced = True
            
            # 策略 B: 如果關鍵字被 Word 切割在不同 Run 中，則重寫整個段落文字
            # (會繼承段落的第一個 Run 的格式，通常是足夠的)
            if not replaced:
                paragraph.text = paragraph.text.replace(key, value)
                for run in paragraph.runs:
                    ensure_chinese_font(run)

def replace_placeholder_with_image_in_paragraph(paragraph, placeholder, image_stream):
    """
    找到段落中的 {img_X} 並原地換成圖片。
    """
    if placeholder in paragraph.text:
        # 1. 清空該段落的文字 (把 {img_1} 刪掉)
        paragraph.text = "" 
        
        # 2. 在該段落加入圖片 Run
        # 這樣圖片就會遵循該段落的對齊設定 (例如置中)
        run = paragraph.add_run()
        if image_stream:
            # 圖片寬度固定 8cm (適應一般表格欄寬)
            run.add_picture(image_stream, width=Cm(8.0))

# --- 3. 主流程 ---

def generate_fixed_report(template_file, context, photo_data):
    doc = Document(template_file)
    
    # 1. 準備全域取代資料 (工程名稱、廠商等)
    # 格式：{project_name} -> 值
    text_replacements = {f"{{{k}}}": v for k, v in context.items()}
    
    # 2. 執行全域文字替換 (包含基本資料表格)
    # 遍歷所有表格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_text_in_paragraph(paragraph, text_replacements)
                    
    # 遍歷所有一般段落
    for paragraph in doc.paragraphs:
        replace_text_in_paragraph(paragraph, text_replacements)
    
    # 3. 處理照片與說明 (針對 {img_X} 和 {info_X})
    # 我們需要遍歷文檔中的所有段落(含表格內)，找到這些特定的佔位符
    
    # 為了效率，我們先建立好每一張照片的取代資料
    img_map = {}  # { "{img_1}": image_stream, ... }
    info_map = {} # { "{info_1}": text_content, ... }
    
    for i in range(1, 9): # 支援 1~8
        img_key = f"{{img_{i}}}"
        info_key = f"{{info_{i}}}"
        
        data_idx = i - 1
        if data_idx < len(photo_data):
            # 有資料
            data = photo_data[data_idx]
            img_map[img_key] = compress_image(data['file'])
            
            # 組合說明文字
            info_text = f"照片編號：{data['no']:02d}　　　　日期：{data['date_str']}\n"
            info_text += f"說明：{data['desc']}\n"
            info_text += f"實測：{data['result']}"
            info_map[info_key] = info_text
        else:
            # 沒資料 -> 設為 None 或空字串，稍後清除
            img_map[img_key] = None
            info_map[info_key] = ""

    # 4. 再次遍歷文件，執行照片與說明的精準替換
    # (必須遍歷所有表格儲存格，因為您的定位點在表格裡)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    # 檢查是否有圖片佔位符
                    for k, img_stream in img_map.items():
                        if k in paragraph.text:
                            replace_placeholder_with_image_in_paragraph(paragraph, k, img_stream)
                    
                    # 檢查是否有文字佔位符 (使用之前的文字替換邏輯)
                    replace_text_in_paragraph(paragraph, info_map)

    return doc

# --- 4. Streamlit UI ---

st.set_page_config(page_title="自主檢查表生成器", layout="wide")
st.title("🏗️ 工程自主檢查表 (樣式繼承版)")

if 'doc_buffer' not in st.session_state:
    st.session_state['doc_buffer'] = None
if 'doc_name' not in st.session_state:
    st.session_state['doc_name'] = ""

with st.sidebar:
    st.header("1. 上傳樣板")
    st.info("請確認 Word 表格內已預先填好 `{img_1}`...`{img_8}` 及 `{info_1}`...`{info_8}`，並調整好您想要的大小與位置。")
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
            # 限制 8 張
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
