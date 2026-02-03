import streamlit as st
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml.ns import qn
from PIL import Image
import io
import datetime

# --- 1. 樣式複製核心工具 (關鍵修正) ---

def get_run_style(run):
    """
    【關鍵功能】記錄原本 Word 樣板裡文字的格式
    包含：字型名稱、中文字型、大小、粗體、斜體、底線、顏色
    """
    style = {}
    style['name'] = run.font.name
    style['size'] = run.font.size
    style['bold'] = run.bold
    style['italic'] = run.italic
    style['underline'] = run.underline
    style['color'] = run.font.color.rgb
    
    # 嘗試獲取中文字型設定 (East Asia Font)
    try:
        rPr = run._element.rPr
        if rPr is not None and rPr.rFonts is not None:
            style['eastAsia'] = rPr.rFonts.get(qn('w:eastAsia'))
        else:
            style['eastAsia'] = None
    except:
        style['eastAsia'] = None
        
    return style

def apply_run_style(run, style):
    """
    【關鍵功能】將記錄下來的格式，套用到新的文字上
    """
    if style.get('name'): run.font.name = style.get('name')
    if style.get('size'): run.font.size = style.get('size')
    if style.get('bold') is not None: run.bold = style.get('bold')
    if style.get('italic') is not None: run.italic = style.get('italic')
    if style.get('underline') is not None: run.underline = style.get('underline')
    if style.get('color'): run.font.color.rgb = style.get('color')
    
    # 套用中文字型
    if style.get('eastAsia'):
        run._element.rPr.rFonts.set(qn('w:eastAsia'), style.get('eastAsia'))

def compress_image(image_file, max_width=800):
    """圖片壓縮與轉向"""
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

# --- 2. 替換邏輯：先備份樣式，再替換文字 ---

def replace_text_content(doc, replacements):
    """通用文字替換"""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_paragraph_strict(paragraph, replacements)
    for paragraph in doc.paragraphs:
        replace_paragraph_strict(paragraph, replacements)

def replace_paragraph_strict(paragraph, replacements):
    """
    嚴格保留格式的替換邏輯：
    1. 嘗試在單一 Run 替換 (最完美)。
    2. 若失敗，則重寫段落，但強制套用「第一個 Run」的原始樣式。
    """
    if not paragraph.text:
        return

    original_text = paragraph.text
    # 檢查是否有需要替換的關鍵字
    needs_replace = False
    for key in replacements:
        if key in original_text:
            needs_replace = True
            break
    
    if not needs_replace:
        return

    # 策略 A: 嘗試簡單替換 (不破壞 Run 結構)
    # 如果關鍵字剛好在一個 Run 裡面，直接換掉文字，格式會自動保留
    for run in paragraph.runs:
        for key, value in replacements.items():
            if key in run.text:
                if value is None: value = ""
                run.text = run.text.replace(key, str(value))
                # 成功替換後，不需要做其他事，格式原本就在
    
    # 再次檢查是否還有殘留的 Key (代表 Key 被 Word 切割在不同 Run 之間)
    remaining_text = paragraph.text
    still_has_key = False
    for key in replacements:
        if key in remaining_text:
            still_has_key = True
            break
            
    # 策略 B: 如果關鍵字被切割，必須重寫段落，但要「複製樣式」
    if still_has_key:
        # 1. 備份第一個 Run 的樣式 (通常是我們想要的樣式)
        saved_style = {}
        if paragraph.runs:
            saved_style = get_run_style(paragraph.runs[0])
        
        # 2. 執行全段落文字替換
        new_text = original_text
        for key, value in replacements.items():
            if value is None: value = ""
            new_text = new_text.replace(key, str(value))
            
        # 3. 清空舊內容
        paragraph.clear() 
        # (clear() 會保留段落屬性如置中，但刪除所有 run)
        
        # 4. 加入新文字並套用備份的樣式
        new_run = paragraph.add_run(new_text)
        apply_run_style(new_run, saved_style)

def replace_placeholder_with_image(doc, placeholder, image_stream):
    """找到 {img_X} 並替換成圖片"""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if placeholder in paragraph.text:
                        # 備份對齊方式 (通常已經設定好)
                        alignment = paragraph.alignment
                        paragraph.text = "" 
                        paragraph.alignment = alignment
                        
                        run = paragraph.add_run()
                        if image_stream:
                            # 圖片寬度固定 8cm
                            run.add_picture(image_stream, width=Cm(8.0))
                        return 

# --- 3. 主流程 ---

def generate_fixed_report(template_file, context, photo_data):
    doc = Document(template_file)
    
    # 1. 填入基本資料
    # 格式：{key} -> value
    text_replacements = {f"{{{k}}}": v for k, v in context.items()}
    replace_text_content(doc, text_replacements)
    
    # 2. 填入照片與說明 (1~8)
    for i in range(1, 9):
        img_key = f"{{img_{i}}}"
        info_key = f"{{info_{i}}}"
        
        data_idx = i - 1
        if data_idx < len(photo_data):
            data = photo_data[data_idx]
            
            # (A) 填入圖片
            replace_placeholder_with_image(doc, img_key, compress_image(data['file']))
            
            # (B) 填入文字 (日期往右調整)
            # 這裡加入了 8 個全形空白，讓日期更靠右
            info_text = f"照片編號：{data['no']:02d}　　　　　　　　日期：{data['date_str']}\n"
            info_text += f"說明：{data['desc']}\n"
            info_text += f"實測：{data['result']}"
            
            replace_text_content(doc, {info_key: info_text})
            
        else:
            # 沒資料則清空
            replace_text_content(doc, {img_key: ""})
            replace_text_content(doc, {info_key: ""})
            
    return doc

# --- 4. Streamlit UI ---

st.set_page_config(page_title="自主檢查表生成器", layout="wide")
st.title("🏗️ 工程自主檢查表 (樣式完美複製版)")

if 'doc_buffer' not in st.session_state:
    st.session_state['doc_buffer'] = None
if 'doc_name' not in st.session_state:
    st.session_state['doc_name'] = ""

with st.sidebar:
    st.header("1. 上傳樣板")
    st.info("請確認 Word 樣板內的 `{project_name}` 或 `{info_1}` 已經設定好您要的字體大小與粗細。程式會直接複製它。")
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
