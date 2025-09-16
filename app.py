#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook Timeline - Streamlit 網頁應用程式
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import json
from outlook_timeline import OutlookTimeline, EmailInfo
import base64

# 設定頁面配置
st.set_page_config(
    page_title="Outlook Timeline",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自訂 CSS
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 10px;
        margin-bottom: 2rem;
        color: white;
        text-align: center;
    }
    
    .metric-card {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #667eea;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin-bottom: 1rem;
    }
    
    .email-card {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #28a745;
        margin-bottom: 1rem;
    }
    
    .keyword-tag {
        background: #ff6b6b;
        color: white;
        padding: 0.2rem 0.5rem;
        border-radius: 12px;
        font-size: 0.8rem;
        margin: 0.2rem;
        display: inline-block;
    }
    
    .stButton > button {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.5rem 2rem;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# 主標題
st.markdown("""
<div class="main-header">
    <h1>📧 Outlook Timeline</h1>
    <p>M365 郵件關鍵字搜尋與時間軸分析工具</p>
</div>
""", unsafe_allow_html=True)

# 初始化 session state
if 'analyzer' not in st.session_state:
    st.session_state.analyzer = None
if 'emails' not in st.session_state:
    st.session_state.emails = []
if 'connected' not in st.session_state:
    st.session_state.connected = False

# 側邊欄 - 連接設定
st.sidebar.header("🔐 帳號設定")

with st.sidebar.expander("M365 連接設定", expanded=not st.session_state.connected):
    # 嘗試從 Streamlit secrets 讀取預設值
    default_username = ""
    default_server = "outlook.office365.com"
    default_port = 993
    
    try:
        if "outlook" in st.secrets:
            default_username = st.secrets["outlook"].get("M365_USERNAME", "")
            default_server = st.secrets["outlook"].get("IMAP_SERVER", "outlook.office365.com")
            default_port = int(st.secrets["outlook"].get("IMAP_PORT", 993))
    except:
        pass
    
    username = st.text_input("M365 帳號", value=default_username)
    
    # 密碼類型選擇
    auth_type = st.radio(
        "認證方式",
        ["應用程式密碼 (推薦)", "帳號密碼 + 兩段驗證"],
        help="建議使用應用程式密碼以避免兩段驗證問題"
    )
    
    if auth_type == "應用程式密碼 (推薦)":
        password = st.text_input("應用程式密碼", type="password", 
                                help="在 Microsoft 帳戶安全性頁面生成的應用程式專用密碼")
        mfa_code = None
    else:
        password = st.text_input("帳號密碼", type="password")
        mfa_code = st.text_input("兩段驗證碼 (6位數字)", max_chars=6,
                                help="請輸入您手機上顯示的 6 位數驗證碼")
    
    server = st.text_input("IMAP 伺服器", value=default_server)
    port = st.number_input("端口", value=default_port, min_value=1, max_value=65535)
    
    if st.button("🔗 連接", key="connect_btn"):
        if username and password:
            # 檢查兩段驗證
            if auth_type == "帳號密碼 + 兩段驗證" and not mfa_code:
                st.warning("請輸入兩段驗證碼")
            else:
                try:
                    with st.spinner("正在連接到 M365..."):
                        # 如果使用兩段驗證，將驗證碼附加到密碼
                        final_password = password
                        if auth_type == "帳號密碼 + 兩段驗證" and mfa_code:
                            final_password = f"{password}{mfa_code}"
                        
                        analyzer = OutlookTimeline(username, final_password)
                        if analyzer.connect():
                            st.session_state.analyzer = analyzer
                            st.session_state.connected = True
                            st.success("✓ 連接成功！")
                            
                            # 取得資料夾清單
                            folders = analyzer.get_folders()
                            st.session_state.folders = folders
                            st.info(f"找到 {len(folders)} 個資料夾")
                        else:
                            if auth_type == "帳號密碼 + 兩段驗證":
                                st.error("✗ 連接失敗 - 請檢查密碼和驗證碼")
                            else:
                                st.error("✗ 連接失敗 - 請檢查應用程式密碼")
                except Exception as e:
                    st.error(f"連接錯誤: {e}")
                    if auth_type == "帳號密碼 + 兩段驗證":
                        st.info("💡 建議：使用應用程式密碼可避免兩段驗證問題")
        else:
            st.warning("請輸入帳號和密碼")

# 連接狀態顯示
if st.session_state.connected:
    st.sidebar.success("🟢 已連接")
    if st.sidebar.button("🔌 中斷連接"):
        if st.session_state.analyzer:
            st.session_state.analyzer.disconnect()
        st.session_state.connected = False
        st.session_state.analyzer = None
        st.session_state.emails = []
        st.rerun()
else:
    st.sidebar.error("🔴 未連接")

# 側邊欄 - 搜尋設定
st.sidebar.header("🔍 搜尋設定")

# 關鍵字預設
keyword_presets = {
    "危機管理": ["緊急", "危機", "問題", "故障", "異常", "事故", "投訴", "客訴"],
    "專案追蹤": ["專案", "計畫", "進度", "里程碑", "截止日期", "交付", "完成"],
    "安全事件": ["安全", "資安", "入侵", "病毒", "漏洞", "威脅", "防護"],
    "自訂": []
}

preset_choice = st.sidebar.selectbox("選擇關鍵字預設", list(keyword_presets.keys()))

if preset_choice == "自訂":
    keywords_input = st.sidebar.text_input("輸入關鍵字 (以逗號分隔)")
    keywords = [k.strip() for k in keywords_input.split(",") if k.strip()] if keywords_input else []
else:
    keywords = keyword_presets[preset_choice]
    st.sidebar.write("預設關鍵字:")
    for keyword in keywords:
        st.sidebar.write(f"• {keyword}")

# 搜尋參數
days_back = st.sidebar.slider("搜尋天數", 1, 365, 30)

# 資料夾選擇
if st.session_state.connected and 'folders' in st.session_state:
    default_folders = ["INBOX", "Sent Items"]
    available_folders = [f for f in st.session_state.folders if f in default_folders]
    selected_folders = st.sidebar.multiselect("選擇資料夾", st.session_state.folders, default=available_folders)
else:
    selected_folders = ["INBOX", "Sent Items"]

include_sent = st.sidebar.checkbox("包含寄件備份", value=True)

# 搜尋按鈕
search_button = st.sidebar.button("🔍 開始搜尋", disabled=not st.session_state.connected or not keywords)

# 主要內容區域
if st.session_state.connected and search_button:
    if keywords:
        with st.spinner("正在搜尋郵件..."):
            try:
                emails = st.session_state.analyzer.search_emails(
                    keywords=keywords,
                    folders=selected_folders if selected_folders else None,
                    days_back=days_back,
                    include_sent=include_sent
                )
                st.session_state.emails = emails
                st.success(f"找到 {len(emails)} 封相關郵件")
            except Exception as e:
                st.error(f"搜尋錯誤: {e}")
    else:
        st.warning("請輸入搜尋關鍵字")

# 顯示結果
if st.session_state.emails:
    emails = st.session_state.emails
    
    # 統計資訊
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("總計郵件", len(emails))
    
    with col2:
        if emails:
            date_range = (emails[-1].date - emails[0].date).days + 1
            st.metric("天數範圍", date_range)
    
    with col3:
        all_keywords = set()
        for email in emails:
            all_keywords.update(email.keywords_found)
        st.metric("找到關鍵字", len(all_keywords))
    
    with col4:
        folders_count = len(set(email.folder for email in emails))
        st.metric("涉及資料夾", folders_count)
    
    # 時間分布圖表
    st.subheader("📊 時間分布分析")
    
    # 準備圖表資料
    df_emails = pd.DataFrame([
        {
            'date': email.date,
            'subject': email.subject,
            'sender': email.sender,
            'folder': email.folder,
            'keywords': ', '.join(email.keywords_found)
        }
        for email in emails
    ])
    
    # 按日期分組統計
    df_daily = df_emails.groupby(df_emails['date'].dt.date).size().reset_index()
    df_daily.columns = ['date', 'count']
    
    # 繪製時間分布圖
    fig = px.line(df_daily, x='date', y='count', 
                  title='每日郵件數量趨勢',
                  labels={'date': '日期', 'count': '郵件數量'})
    fig.update_layout(height=400)
    st.plotly_chart(fig, use_container_width=True)
    
    # 關鍵字分布圖
    keyword_counts = {}
    for email in emails:
        for keyword in email.keywords_found:
            keyword_counts[keyword] = keyword_counts.get(keyword, 0) + 1
    
    if keyword_counts:
        st.subheader("🏷️ 關鍵字分布")
        df_keywords = pd.DataFrame(list(keyword_counts.items()), columns=['關鍵字', '出現次數'])
        fig_keywords = px.bar(df_keywords, x='關鍵字', y='出現次數', 
                             title='關鍵字出現頻率')
        fig_keywords.update_layout(height=400)
        st.plotly_chart(fig_keywords, use_container_width=True)
    
    # 資料夾分布圖
    folder_counts = df_emails['folder'].value_counts()
    if len(folder_counts) > 1:
        st.subheader("📁 資料夾分布")
        fig_folders = px.pie(values=folder_counts.values, names=folder_counts.index,
                            title='郵件資料夾分布')
        fig_folders.update_layout(height=400)
        st.plotly_chart(fig_folders, use_container_width=True)
    
    # 詳細郵件清單
    st.subheader("📧 詳細郵件清單")
    
    # 排序選項
    sort_options = {
        "時間 (新到舊)": lambda x: -x.date.timestamp(),
        "時間 (舊到新)": lambda x: x.date.timestamp(),
        "寄件者": lambda x: x.sender,
        "主旨": lambda x: x.subject
    }
    
    sort_choice = st.selectbox("排序方式", list(sort_options.keys()))
    sorted_emails = sorted(emails, key=sort_options[sort_choice])
    
    # 分頁顯示
    emails_per_page = 10
    total_pages = (len(sorted_emails) + emails_per_page - 1) // emails_per_page
    
    if total_pages > 1:
        page = st.selectbox("頁數", range(1, total_pages + 1))
        start_idx = (page - 1) * emails_per_page
        end_idx = start_idx + emails_per_page
        page_emails = sorted_emails[start_idx:end_idx]
    else:
        page_emails = sorted_emails
    
    # 顯示郵件
    for i, email in enumerate(page_emails):
        with st.expander(f"📧 {email.subject} - {email.date.strftime('%Y-%m-%d %H:%M')}"):
            col1, col2 = st.columns([2, 1])
            
            with col1:
                st.write(f"**寄件者:** {email.sender}")
                st.write(f"**收件者:** {email.receiver}")
                st.write(f"**日期:** {email.date.strftime('%Y-%m-%d %H:%M:%S')}")
                st.write(f"**資料夾:** {email.folder}")
            
            with col2:
                st.write("**找到的關鍵字:**")
                keywords_html = ""
                for keyword in email.keywords_found:
                    keywords_html += f'<span class="keyword-tag">{keyword}</span> '
                st.markdown(keywords_html, unsafe_allow_html=True)
            
            st.write("**內容預覽:**")
            preview = email.body[:500] + "..." if len(email.body) > 500 else email.body
            st.text_area("", preview, height=100, key=f"preview_{i}", disabled=True)
    
    # 匯出功能
    st.subheader("💾 匯出報告")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("📄 匯出 CSV"):
            analyzer = OutlookTimeline()
            analyzer.emails = emails
            csv_data = analyzer._generate_csv_report()
            st.download_button(
                label="下載 CSV 檔案",
                data=csv_data,
                file_name=f"outlook_timeline_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
    
    with col2:
        if st.button("📊 匯出 JSON"):
            analyzer = OutlookTimeline()
            analyzer.emails = emails
            json_data = analyzer.generate_timeline_report("json")
            st.download_button(
                label="下載 JSON 檔案",
                data=json_data,
                file_name=f"outlook_timeline_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                mime="application/json"
            )
    
    with col3:
        if st.button("🌐 匯出 HTML"):
            analyzer = OutlookTimeline()
            analyzer.emails = emails
            html_data = analyzer._generate_html_report()
            st.download_button(
                label="下載 HTML 檔案",
                data=html_data,
                file_name=f"outlook_timeline_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html",
                mime="text/html"
            )

else:
    # 歡迎頁面
    if not st.session_state.connected:
        st.info("👈 請先在左側設定您的 M365 帳號資訊並連接")
        
        st.markdown("""
        ## 🚀 使用說明
        
        ### 1. 連接設定
        
        #### 方式一：應用程式密碼 (推薦) 🔐
        - 輸入您的 M365 帳號
        - 使用應用程式密碼登入
        - **優點**：無需每次輸入驗證碼，更安全穩定
        
        #### 方式二：帳號密碼 + 兩段驗證 📱
        - 輸入您的 M365 帳號和原始密碼
        - 輸入手機上顯示的 6 位數驗證碼
        - **注意**：驗證碼有時效性，需要快速輸入
        
        ### 2. 搜尋郵件
        - 選擇預設關鍵字組合或自訂關鍵字
        - 設定搜尋時間範圍
        - 選擇要搜尋的資料夾
        
        ### 3. 分析結果
        - 檢視時間軸分布圖表
        - 分析關鍵字出現頻率
        - 瀏覽詳細郵件內容
        - 匯出報告 (CSV/JSON/HTML)
        
        ### 💡 如何設定應用程式密碼
        1. 登入 [Microsoft 帳戶安全性](https://account.microsoft.com/security)
        2. 選擇「進階安全性選項」
        3. 在「應用程式密碼」區塊點擊「建立新的應用程式密碼」
        4. 輸入應用程式名稱（例如：Outlook Timeline）
        5. 複製產生的密碼並貼到左側的應用程式密碼欄位
        
        ### ⚠️ 常見問題
        - **連接失敗**：確認已在 Outlook 設定中啟用 IMAP
        - **兩段驗證失敗**：驗證碼有時效性，請盡快輸入
        - **企業帳號問題**：可能需要聯繫 IT 管理員啟用 IMAP
        """)
    else:
        st.info("👈 請在左側設定搜尋參數並開始搜尋")

# 頁尾
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <p>© 2025 Outlook Timeline - M365 郵件關鍵字搜尋與時間軸分析工具</p>
    <p>適用於危機事件追蹤、專案管理和安全分析</p>
</div>
""", unsafe_allow_html=True)