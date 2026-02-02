import streamlit as st
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from PIL import Image
import io
import datetime

# --- 工具函數區 ---

def set_font(run, font_name='標楷體', size=12):
    """設定中文字型與大小"""
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)

def get_roc_date(date_obj):
    """將西元年轉換為民國年格式 (e.g., 115.01.13)"""
    roc_year = date_obj.year - 1911
    return f"{roc_year}.{date_obj.month:02d}.{date_obj.day:02d}"

def compress_image(image_file, max_width=800):
    """壓縮圖片以縮小 Word 檔案大小"""
    img = Image.open(image_file)
    # 如果是 RGBA (透明背景) 轉為 RGB
    if img.mode == 'RGBA':
        img = img.convert('RGB')
    
    # 等比例縮放
    ratio = max_width / float(img.size[0])
    if ratio < 1:
        h_size = int((float(img.size[1]) * float(ratio)))
        img = img.resize((max_width, h_size), Image.Resampling.LANCZOS)
    
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='JPEG', quality=70) # 壓縮品質 70%
    img_byte_arr.seek(0)
    return img_byte_arr

# --- Word 生成核心邏輯 ---

def generate_docx(project_info, photo_data):
    doc = Document()
    
    # 設定版面邊界 (依照一般工程報告習慣微調)
    section = doc.sections[0]
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

    # --- 1. 建立表頭資訊 (Header Table) ---
    # 根據你的範例，這是一個 6 列的表格
    table = doc.add_table(rows=6, cols=4)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # 定義欄位名稱與對應的值
    headers = [
        ("工程名稱", project_info['project_name'], 3), # (標題, 內容, 合併欄位數)
        ("洽辦機關", project_info['agency'], 3),
        ("代辦機關", project_info['sub_agency'], 3),
        ("設計監造", project_info['designer'], 3),
        ("施工廠商", project_info['contractor'], 3),
    ]

    # 填入前 5 列 (固定格式)
    for i, (label, value, span) in enumerate(headers):
        row = table.rows[i]
        # 第一格：標題
        cell_label = row.cells[0]
        p = cell_label.paragraphs[0]
        run = p.add_run(label)
        set_font(run, size=14)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell_label.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        
        # 第二格：內容 (合併後面的儲存格)
        cell_value = row.cells[1]
        # 合併儲存格邏輯
        if span > 0:
            cell_value.merge(row.cells[1+span-1])
        
        p = cell_value.paragraphs[0]
        run = p.add_run(value)
        set_font(run, size=14)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell_value.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # 第 6 列：位置、日期、項目 (比較複雜，手動處理)
    row_6 = table.rows[5]
    
    # 抽查位置
    row_6.cells[0].text = "抽/查驗位置"
    set_font(row_6.cells[0].paragraphs[0].runs[0], size=12)
    row_6.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    row_6.cells[1].text = project_info['location']
    set_font(row_6.cells[1].paragraphs[0].runs[0], size=12)
    row_6.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 抽查日期 (標題在 cell 2, 日期在 cell 3) -> 這裡你的範例有點不同，我依照通用邏輯調整
    # 你的範例是：位置 | (內容) | 日期 | (內容)
    row_6.cells[2].text = "抽/查驗日期"
    set_font(row_6.cells[2].paragraphs[0].runs[0], size=12)
    row_6.cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    roc_date_str = get_roc_date(project_info['date'])
    row_6.cells[3].text = roc_date_str
    set_font(row_6.cells[3].paragraphs[0].runs[0], size=12)
    row_6.cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 項目欄位 (新增一列給項目名稱)
    row_item = table.add_row()
    row_item.cells[0].text = "抽/查驗項目"
    set_font(row_item.cells[0].paragraphs[0].runs[0], size=12)
    row_item.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    row_item.cells[1].merge(row_item.cells[3])
    row_item.cells[1].text = project_info['check_item']
    set_font(row_item.cells[1].paragraphs[0].runs[0], size=14)
    row_item.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 檢查內容 (這裡預留給標準檢核項目，如果要自動化這部分，需要更多資料庫邏輯)
    # 暫時插入一個空白列代表檢查內容區域
    row_content = table.add_row()
    row_content.height = Cm(4) # 預留高度
    row_content.cells[0].text = "抽/查驗情形"
    row_content.cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    row_content.cells[1].merge(row_content.cells[3])
    row_content.cells[1].text = project_info['check_content'] # 使用者輸入的檢查項目內容
    
    # 換頁，開始放照片
    doc.add_page_break()

    # --- 2. 照片區 (Photo Section) ---
    # 標題
    p_title = doc.add_paragraph()
    run = p_title.add_run("檢 查 照 片")
    set_font(run, size=16)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 照片表格：每列 1 張或 2 張，你的範例是一次一張大圖配說明，或左右兩張
    # 為了版面整齊，我們採用「一列兩張」的矩陣模式 (最常見且省紙)
    # 或是依照你的檔案 Source 34，是一張圖配下方詳細說明。
    
    # 採用通用模式：建立一個大表格來排版
    # 邏輯：每張照片佔據一個區塊：[照片] (換行) [編號/日期] (換行) [說明] (換行) [實測]
    
    # 為了讓排版最漂亮，我們使用 2 欄的表格，每欄放一張照片的完整資訊
    photo_table = doc.add_table(rows=0, cols=2)
    photo_table.autofit = False 
    photo_table.allow_autofit = False
    
    # 設定欄寬 (總寬度約 17cm，每欄 8.5cm)
    for col in photo_table.columns:
        col.width = Cm(8.5)

    current_row = None
    
    for i, p_data in enumerate(photo_data):
        # 每 2 張照片換一列
        if i % 2 == 0:
            current_row = photo_table.add_row()
        
        # 決定是左欄還是右欄
        cell = current_row.cells[i % 2]
        
        # 1. 插入段落放置圖片
        p_img = cell.paragraphs[0]
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        try:
            compressed_img = compress_image(p_data['file'])
            run = p_img.add_run()
            run.add_picture(compressed_img, width=Cm(8.0)) # 限制圖片寬度
        except Exception as e:
            p_img.add_run(f"[圖片讀取失敗: {e}]")

        # 2. 插入文字資訊表格 (嵌套表格或直接文字)
        # 直接用文字排版比較穩定
        info_text = (
            f"照片編號：{p_data['no']:02d}    日期：{roc_date_str}\n"
            f"說明：{p_data['desc']}\n"
            f"實測：{p_data['result']}"
        )
        p_info = cell.add_paragraph(info_text)
        p_info.paragraph_format.space_before = Pt(4)
        # 設定中文字型
        for run in p_info.runs:
            set_font(run, size=10)

    return doc

# --- Streamlit UI 介面 ---

st.set_page_config(page_title="工程自主檢查表產生器", page_icon="🏗️", layout="wide")

st.title("🏗️ 工程施工自主檢查表產生系統")
st.markdown("---")

# 側邊欄：全域設定
with st.sidebar:
    st.header("📝 專案基本資料")
    default_project = "衛生福利部防疫中心興建工程"
    project_name = st.text_input("工程名稱", value=default_project)
    contractor = st.text_input("施工廠商", value="豐譽營造股份有限公司")
    agency = st.text_input("洽辦機關", value="衛生福利部疾病管制署")
    sub_agency = st.text_input("代辦機關", value="內政部國土管理署")
    designer = st.text_input("設計監造", value="劉培森建築師事務所")
    
    st.markdown("---")
    st.header("📅 檢查資訊")
    check_date = st.date_input("檢查日期", datetime.date.today())
    location = st.text_input("施作位置 (e.g., 北棟 6F)", value="北棟")
    check_item = st.text_input("自檢項目 (e.g., 拆除工程)", value="拆除工程施工自主檢查(精細拆除)")
    
    st.markdown("---")
    check_content = st.text_area("檢查標準/內容 (顯示於表頭)", 
                                 value="1. 現場既有雜物整理\n2. 室裝材分類拆除集中\n3. 依可回收,不可回收,有價物分類",
                                 height=100)

# 主畫面：照片上傳與編輯
st.header("📸 照片上傳與說明")
uploaded_files = st.file_uploader("請上傳現場照片 (支援多選)", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)

photo_data = []

if uploaded_files:
    st.info(f"已上傳 {len(uploaded_files)} 張照片。請在下方填寫詳細資訊。")
    
    # 使用 Form 避免每打一個字就重整一次頁面
    with st.form("photo_details_form"):
        # 使用 Grid 排版，每行顯示 2 張照片的編輯區
        cols = st.columns(2)
        
        for i, file in enumerate(uploaded_files):
            col = cols[i % 2]
            with col:
                st.image(file, use_column_width=True, caption=file.name)
                # 預設編號
                p_no = i + 1
                # 輸入欄位
                c1, c2 = st.columns([1, 3])
                no_input = c1.number_input(f"編號 #{i+1}", value=p_no, min_value=1, key=f"no_{i}")
                
                # 預設說明文字 (智慧預填：如果是拆除工程，預填常見詞)
                default_desc = "依施工計畫執行"
                default_result = "與計畫相符"
                if "拆除" in check_item:
                    default_desc = "室裝材分類拆除集中"
                    default_result = "室裝材分類拆除集中"
                
                desc_input = st.text_input(f"說明 #{i+1}", value=default_desc, key=f"desc_{i}")
                result_input = st.text_input(f"實測/結果 #{i+1}", value=default_result, key=f"res_{i}")
                
                photo_data.append({
                    "file": file,
                    "no": no_input,
                    "desc": desc_input,
                    "result": result_input
                })
                st.markdown("---")
        
        submit_btn = st.form_submit_button("✅ 確認資料並生成報表", use_container_width=True)

    if submit_btn:
        # 彙整資訊
        project_info = {
            "project_name": project_name,
            "contractor": contractor,
            "agency": agency,
            "sub_agency": sub_agency,
            "designer": designer,
            "date": check_date,
            "location": location,
            "check_item": check_item,
            "check_content": check_content
        }
        
        with st.spinner("正在生成 Word 文件中..."):
            doc = generate_docx(project_info, photo_data)
            
            # 儲存到記憶體
            bio = io.BytesIO()
            doc.save(bio)
            
            # 下載按鈕
            file_name = f"{get_roc_date(check_date)}{check_item.split('(')[0]}_{location}.docx"
            st.success("🎉 報表生成成功！")
            st.download_button(
                label="📥 下載 Word 報表 (.docx)",
                data=bio.getvalue(),
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

else:
    st.info("👋 請從左側確認專案資料，並在上方上傳照片以開始使用。")
