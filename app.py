import streamlit as st
import pandas as pd
import io
import zipfile
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from datetime import datetime
import pytz

# --- 0. 設定頁面 ---
st.set_page_config(page_title="工地自檢表回報系統", page_icon="🏗️")

# --- 1. 設定收件人名單 (請依需求修改這裡) ---
# 格式： "顯示名稱": "Email地址"
# 這裡可以用中文名稱，對應到實際的 Email
RECIPIENTS = {
    "總公司工務部": "office_main@example.com", 
    "專案經理": "manager@example.com",
    "測試用 (寄給自己)": st.secrets["email"]["account"] # 這會讀取您的寄件帳號
}

st.title("🏗️ 施工自檢表回報系統")
st.info("💡 手機端建議使用「寄送」；電腦端可使用「下載」。")

# --- 2. 輸入介面 ---
with st.expander("📝 1. 填寫檢查內容", expanded=True):
    # 選擇收件人
    selected_name = st.selectbox("📬 請選擇收件單位", list(RECIPIENTS.keys()))
    target_email = RECIPIENTS[selected_name]
    
    col1, col2 = st.columns(2)
    with col1:
        project_name = st.text_input("專案名稱", value="A棟排樁工程")
    with col2:
        inspector = st.text_input("檢查人員", value="王小明")
    
    # 時間設定 (台灣時間)
    tw_timezone = pytz.timezone('Asia/Taipei')
    today = datetime.now(tw_timezone).strftime("%Y-%m-%d")
    st.caption(f"📅 檢查日期：{today}")
    
    # 檢查項目 (模擬排樁/預壘樁工程)
    st.write("---")
    check_1 = st.checkbox("1. 樁位放樣點位確認", value=True)
    check_2 = st.checkbox("2. 鋼筋籠長度及保護層檢查", value=True)
    check_3 = st.checkbox("3. 特密管位置及深度確認")
    check_4 = st.checkbox("4. 混凝土澆置紀錄完整")
    
    note = st.text_area("現場備註事項", "今日施工進度正常。")
    
    uploaded_photos = st.file_uploader("📸 2. 現場照片上傳 (可多選)", accept_multiple_files=True, type=['jpg', 'png', 'jpeg'])

# --- 3. 核心功能：生成與寄信 ---
def create_zip_file():
    """生成 ZIP 檔案並回傳 BytesIO 物件 (不落地)"""
    # A. 製作 Excel 數據
    data = {
        "檢查項目": ["樁位放樣", "鋼筋籠檢查", "特密管確認", "混凝土澆置", "現場備註"],
        "檢查結果": ["合格" if check_1 else "不合格", 
                   "合格" if check_2 else "不合格", 
                   "合格" if check_3 else "不合格", 
                   "合格" if check_4 else "不合格",
                   note],
        "檢查日期": [today] * 5,
        "檢查人員": [inspector] * 5
    }
    df = pd.DataFrame(data)
    
    # B. 打包 ZIP (在記憶體中)
    zip_mem = io.BytesIO()
    with zipfile.ZipFile(zip_mem, "w", zipfile.ZIP_DEFLATED) as zf:
        # 1. 寫入 Excel
        with io.BytesIO() as excel_buffer:
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='自檢表')
            zf.writestr(f"{project_name}_自檢表_{today}.xlsx", excel_buffer.getvalue())
        
        # 2. 寫入照片
        if uploaded_photos:
            for photo in uploaded_photos:
                zf.writestr(f"現場照片/{photo.name}", photo.getvalue())
    
    zip_mem.seek(0)
    return zip_mem

def send_email(zip_data, recipient_email, recipient_name):
    """寄信功能"""
    try:
        # 讀取 Secrets (雲端設定)
        gmail_user = st.secrets["email"]["account"]
        gmail_password = st.secrets["email"]["password"]
        
        msg = MIMEMultipart()
        msg['Subject'] = f'【工地回報】{project_name} - {today}'
        msg['From'] = gmail_user
        msg['To'] = recipient_email
        
        body = f"""
        收件單位：{recipient_name}
        專案名稱：{project_name}
        檢查人員：{inspector}
        回報時間：{datetime.now(tw_timezone).strftime("%Y-%m-%d %H:%M")}
        
        ※ 系統自動發送，附件包含 Excel 自檢表與現場照片。
        """
        msg.attach(MIMEText(body, 'plain'))

        # 夾帶 ZIP
        part = MIMEApplication(zip_data.getvalue(), Name="SiteReport.zip")
        part['Content-Disposition'] = f'attachment; filename="{project_name}_{today}_回報.zip"'
        msg.attach(part)

        # 發送
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(gmail_user, gmail_password)
        server.send_message(msg)
        server.quit()
        return True, "發送成功"
    except Exception as e:
        return False, str(e)

# --- 4. 操作按鈕區 ---
st.divider()
st.subheader("🚀 執行操作")

# 初始化 Session State
if 'generated_zip' not in st.session_state:
    st.session_state.generated_zip = None

# 第一步：生成資料
if st.button("步驟 1：生成報表資料", type="primary"):
    if not uploaded_photos and not note:
        st.warning("⚠️ 請至少填寫備註或上傳照片。")
    else:
        with st.spinner("📦 資料打包中..."):
            st.session_state.generated_zip = create_zip_file()
            st.success("✅ 資料已準備就緒！請選擇下一步。")

# 第二步：選擇動作 (只有生成後才會出現)
if st.session_state.generated_zip is not None:
    col_a, col_b = st.columns(2)
    
    # 左邊：寄信
    with col_a:
        if st.button(f"📧 寄送給：{selected_name}"):
            with st.spinner("📨 正在傳送至辦公室..."):
                success, msg = send_email(st.session_state.generated_zip, target_email, selected_name)
                if success:
                    st.success(f"✅ 已成功寄出至 {target_email}")
                else:
                    st.error(f"❌ 寄送失敗：{msg}")

    # 右邊：下載
    with col_b:
        st.download_button(
            label="💾 下載 ZIP 檔案",
            data=st.session_state.generated_zip,
            file_name=f"{project_name}_{today}_回報.zip",
            mime="application/zip"
        )
