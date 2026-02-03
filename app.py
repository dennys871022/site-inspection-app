import streamlit as st
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image
import io
import datetime

# --- 1. 智慧樣式複製工具 (關鍵核心) ---

def get_paragraph_style(paragraph):
    """
    抓取段落中「第一個文字區塊(Run)」的樣式。
    這是為了確保當我們替換文字後，能把原本的大小、粗細、字型都貼回去。
    """
    style = {}
    if paragraph.runs:
        run = paragraph.runs[0]
        style['font_name'] = run.font.name
        style['font_size'] = run.font.size
        style['bold'] = run.bold
        style['italic'] = run.italic
        style['underline'] = run.underline
        style['color'] = run.font.color.rgb
        # 抓取中文字型設定
        try:
            rPr = run._element.rPr
            if rPr is not None and rPr.rFonts is not None:
                style['eastAsia'] = rPr.rFonts.get(qn('w:eastAsia'))
        except:
            pass
    return style

def apply_style_to_run(run, style):
    """將備份的樣式強制套用到新的文字上"""
    if not style: return

    # 1. 套用基本屬性
    if style.get('font_name'): run.font.name = style.get('font_name')
    if style.get('font_size'): run.font.size = style['font_size']
    if style.get('bold') is not None: run.bold = style['bold']
    if style.get('italic') is not None: run.italic = style['italic']
    if style.get('underline') is not None: run.underline = style['underline']
    if style.get('color'): run.font.color.rgb = style['color']
    
    # 2. 套用中文字型 (標楷體等)
    if style.get('eastAsia'):
        run._element.rPr.rFonts.set(qn('w:eastAsia'), style['eastAsia'])
    elif style.get('font_name') == 'Times New Roman':
        # 防呆：如果原本沒設中文字型，但英數是 Times，預設中文給標楷體，比較好看
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

def compress_image(image_file, max_width=800):
    """圖片處理"""
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

# --- 2. 替換邏輯 (智慧版) ---

def smart_replace_text(doc, replacements):
    """
    遍歷整份文件進行替換。
    使用「樣式複製」策略，確保格式 100% 不變。
    """
    # 處理所有表格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    process_paragraph(paragraph, replacements)
    
    # 處理一般段落
    for paragraph in doc.paragraphs:
        process_paragraph(paragraph, replacements)

def process_paragraph(paragraph, replacements):
    """單一段落處理邏輯"""
    if not paragraph.text:
        return

    original_text = paragraph.text
    needs_replace = False
    
    # 檢查是否有任何關鍵字命中
    for key in replacements:
        if key in original_text:
            needs_replace = True
            break
            
    if needs_replace:
        # 1. 先備份樣式 (從第一個 Run 抓，通常代表整段的格式)
        saved_style = get_paragraph_style(paragraph)
        
        # 2. 進行文字替換
        new_text = original_text
        for key, value in replacements.items():
            val_str = str(value) if value is not None else ""
            new_text = new_text.replace(key, val_str)
            
        # 3. 清空舊內容 (保留段落本身的對齊屬性)
        paragraph.clear()
        
        # 4. 填入新文字並「蓋回」原本的樣式
        new_run = paragraph.add_run(new_text)
        apply_style_to_run(new_run, saved_style)

def replace_img_placeholder(doc, placeholder, image_stream):
    """圖片替換邏輯"""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if placeholder in paragraph.text:
                        # 備份段落對齊 (置中/靠左)
                        align = paragraph.alignment
                        paragraph.clear()
                        paragraph.alignment = align
                        
                        run = paragraph.add_run()
                        if image_stream:
                            # 圖片寬度固定 8cm，確保表格整齊
                            run.add_picture(image_stream, width=Cm(8.0))
                        return

# --- 3. 執行流程 ---

def generate_report(template_bytes, context, photo_data):
    doc = Document(io.BytesIO(template_bytes))
    
    # 1. 準備文字替換表 (基本資料)
    text_map = {f"{{{k}}}": v for k, v in context.items()}
    
    # 2. 準備照片資料 (1~8)
    for i in range(1, 9):
        img_key = f"{{img_{i}}}"
        info_key = f"{{info_{i}}}"
        
        idx = i - 1
        if idx < len(photo_data):
            data = photo_data[idx]
            
            # (A) 圖片替換
            replace_img_placeholder(doc, img_key, compress_image(data['file']))
            
            # (B) 文字說明替換
            # 這裡使用 6 個全形空白調整日期位置
            spacer = "\u3000" * 6 
            info_text = f"照片編號：{data['no']:02d}{spacer}日期：{data['date_str']}\n"
            info_text += f"說明：{data['desc']}\n"
            info_text += f"實測：{data['result']}"
            
            text_map[info_key] = info_text
        else:
            # 無照片 -> 清空佔位符
            text_map[img_key] = ""
            text_map[info_key] = "" # 清空說明文字
    
    # 3. 一次性執行所有文字替換 (包含基本資料 & 照片說明)
    smart_replace_text(doc, text_map)
    
    return doc

# --- 4. Streamlit UI ---

st.set_page_config(page_title="自主檢查表生成器", layout="wide")
st.title("🏗️ 工程自主檢查表 (樣式鎖定版)")

# Session State 初始化
if 'saved_template' not in st.session_state:
    st.session_state['saved_template'] = None
if 'template_name' not in st.session_state:
    st.session_state['template_name'] = ""
if 'doc_buffer' not in st.session_state:
    st.session_state['doc_buffer'] = None
if 'doc_name' not in st.session_state:
    st.session_state['doc_name'] = ""

with st.sidebar:
    st.header("1. 樣板管理")
    if st.session_state['saved_template']:
        st.success(f"📂 使用中：{st.session_state['template_name']}")
        st.info("若需更換樣板，請直接上傳新檔案即可。")
    
    uploaded = st.file_uploader("上傳 Word 樣板", type=['docx'])
    if uploaded:
        st.session_state['saved_template'] = uploaded.getvalue()
        st.session_state['template_name'] = uploaded.name
        st.rerun()

    st.markdown("---")
    st.header("2. 專案資訊")
    with st.form("info_form"):
        p_name = st.text_input("工程名稱 {project_name}", "衛生福利部防疫中心興建工程")
        p_cont = st.text_input("施工廠商 {contractor}", "豐譽營造股份有限公司")
        p_sub = st.text_input("協力廠商 {sub_contractor}", "川峻工程有限公司")
        p_loc = st.text_input("施作位置 {location}", "北棟 1F")
        p_item = st.text_input("自檢項目 {check_item}", "拆除工程施工自主檢查(精細拆除) #1")
        check_date = st.date_input("檢查日期", datetime.date.today())
        st.form_submit_button("更新資訊")

    # 日期計算
    roc_year = check_date.year - 1911
    date_str = f"{roc_year}.{check_date.month:02d}.{check_date.day:02d}"

# 主畫面
if st.session_state['saved_template']:
    st.header("3. 照片上傳區 (支援 1~8 張)")
    
    files = st.file_uploader("請選擇照片", type=['jpg','png','jpeg'], accept_multiple_files=True)
    
    photo_data = []
    if files:
        with st.form("photo_form"):
            cols = st.columns(2)
            process_files = files[:8] # 最多取前8張
            
            for i, f in enumerate(process_files):
                with cols[i%2]:
                    st.image(f, width=200)
                    no = st.number_input(f"編號", min_value=1, value=i+1, key=f"n{i}")
                    desc = st.text_input(f"說明", value="現場既有雜物整理", key=f"d{i}")
                    res = st.text_input(f"實測", value="現場既有雜物整理", key=f"r{i}")
                    photo_data.append({
                        "file": f, "no": no, "date_str": date_str, 
                        "desc": desc, "result": res
                    })
            
            if st.form_submit_button("🚀 生成 Word 報告"):
                ctx = {
                    "project_name": p_name, "contractor": p_cont, 
                    "sub_contractor": p_sub, "location": p_loc, 
                    "date": date_str, "check_item": p_item
                }
                try:
                    doc = generate_report(st.session_state['saved_template'], ctx, photo_data)
                    bio = io.BytesIO()
                    doc.save(bio)
                    st.session_state['doc_buffer'] = bio.getvalue()
                    st.session_state['doc_name'] = f"{date_str}_{p_loc}_檢查表.docx"
                    st.success("✅ 生成成功！")
                except Exception as e:
                    st.error(f"發生錯誤: {e}")

        if st.session_state['doc_buffer']:
            st.download_button("📥 下載 Word 檔", st.session_state['doc_buffer'], st.session_state['doc_name'], "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
else:
    st.warning("👈 請先在左側上傳 Word 樣板 (.docx) 才能開始使用。")
