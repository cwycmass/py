import streamlit as st
import pandas as pd
import requests
import time
import io
import re
from bs4 import BeautifulSoup

# --- BILINGUAL DICTIONARY SYSTEM ---
LANG_DICT = {
    "EN": {
        "title": "📚 HKPL School Library Catalog Matcher",
        "lang_select": "🌐 Select Language / 選擇語言",
        "upload_label": "Step 1: Upload School Library Excel sheet (.xlsx)",
        "upload_help": "Maximum allowance: 1,000 rows per session.",
        "select_isbn": "Step 2: Select the column that contains your ISBNs:",
        "btn_start": "🚀 Start Verification",
        "btn_resume": "▶️ Resume Verification",
        "btn_stop": "⏸️ Stop / Pause Safely",
        "btn_reset": "🔄 Clear & Upload New File",
        "btn_retry_errors": "🔄 Retry Timeout Errors",
        "preview_title": "📊 Real-Time Verification Progress Dashboard",
        "err_limit": "❌ Error: The file contains more than 1,000 entries. Please reduce the size and re-upload.",
        "m_total": "Total Rows",
        "m_found": "Found Titles",
        "m_notfound": "Not Found",
        "m_unfinished": "Unfinished Rows",
        "m_error": "Timeout Errors",
        "download_lbl": "📥 Download Updated Excel Spreadsheet",
        "status_idle": "Status: System Ready.",
        "status_running": "Status: Processing records... You can click 'Stop' anytime.",
        "status_paused": "Status: Interrupted by user. Unfinished rows are preserved.",
        "status_completed": "Status: 🎉 Task complete! All rows processed.",
        "col_status": "In HKPL Catalog",
        "col_title": "HKPL Extracted Title",
        "lbl_unfinished": "Unfinished",
        "lbl_na": "N/A"
    },
    "ZH": {
        "title": "📚 香港公共圖書館自動批次圖書比對系統",
        "lang_select": "🌐 Select Language / 選擇語言",
        "upload_label": "第一步：上傳學校圖書館 Excel 檔案 (.xlsx)",
        "upload_help": "系統處理上限：每次最多 1,000 行數據。",
        "select_isbn": "第二步：選擇包含 ISBN 碼的數據欄位：",
        "btn_start": "🚀 開始自動比對",
        "btn_resume": "▶️ 繼續未完成的比對任務",
        "btn_stop": "⏸️ 安全停止 / 暫停",
        "btn_reset": "🔄 清除數據並重新上傳",
        "btn_retry_errors": "🔄 重新測試連線逾時項目",
        "preview_title": "📊 圖書館數據即時驗證進度儀表板",
        "err_limit": "❌ 錯誤：檔案數據超過 1,000 行上限。請縮減行數後重新上傳。",
        "m_total": "總數據量",
        "m_found": "已尋獲書籍種類",
        "m_notfound": "無館藏",
        "m_unfinished": "未完成行數",
        "m_error": "連線逾時錯誤",
        "download_lbl": "📥 下載更新後的 Excel 數據報告",
        "status_idle": "狀態：系統準備就緒。",
        "status_running": "狀態：正在進行線上比對... 您可以隨時點擊「暫停」。",
        "status_paused": "狀態：已被用戶手動暫停。未完成的行數已妥善保存。",
        "status_completed": "狀態：🎉 比對工作全部圓滿完成！",
        "col_status": "香港公共圖書館館藏",
        "col_title": "圖書館登記書籍名稱",
        "lbl_unfinished": "未完成",
        "lbl_na": "無數據"
    }
}

# --- INITIALIZATION & CONFIGURATION ---
st.set_page_config(page_title="HKPL Matcher PRO", page_icon="📚", layout="wide")

# Persistent state initialization
if "current_index" not in st.session_state:
    st.session_state.current_index = 0
if "run_state" not in st.session_state:
    st.session_state.run_state = "IDLE"  # IDLE, RUNNING, PAUSED, COMPLETED
if "df_data" not in st.session_state:
    st.session_state.df_data = None
if "isbn_col_name" not in st.session_state:
    st.session_state.isbn_col_name = ""

# Language Toggle Widget (Sidebar) - Defaulted to Traditional Chinese
selected_lang = st.sidebar.radio("Language / 語言", ["繁體中文", "English"], index=0)
L = LANG_DICT["ZH"] if selected_lang == "繁體中文" else LANG_DICT["EN"]

st.title(L["title"])
st.markdown("---")

# --- UTILITY SCRAPING FUNCTIONS WITH AUTO-RETRY ---
def parse_hkpl_title(html_content):
    """Parses standard HKPL WebCat HTML strings to extract titles if present."""
    soup = BeautifulSoup(html_content, 'html.parser')
    for link in soup.find_all('a'):
        href = link.get('href', '')
        if 'search/detail' in href or 'query?term' in href:
            txt = link.get_text(strip=True)
            if txt and len(txt) > 2:
                return txt
    title_element = soup.find(class_="title") or soup.find(class_="bookTitle")
    if title_element:
        return title_element.get_text(strip=True)
    return "Found (Title text unparsed)"

def check_isbn_on_hkpl(isbn):
    """Executes requests with an automated loop fallback handling connection time-outs."""
    url = f"https://webcat.hkpl.gov.hk/search/query?term_1={isbn}&field_1=isbn&theme=WEB"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    
    max_retries = 3
    for attempt in range(max_retries):
        try:
            response = requests.get(url, headers=headers, timeout=7)
            if response.status_code == 200:
                html = response.text
                
                # Check for direct "Not Found" messages AND HKPL's auto-suggest fallback warnings
                not_found_triggers = [
                    "No record found",
                    "沒有找到符合的紀錄",
                    "系統建議使用",  # HKPL Chinese auto-suggest trigger
                    "System recommends", # HKPL English auto-suggest trigger
                    f"沒有符合{isbn}的檢索結果",
                    f"No results match {isbn}"
                ]
                
                if any(trigger in html for trigger in not_found_triggers):
                    return "No", L["lbl_na"]
                else:
                    extracted_title = parse_hkpl_title(html)
                    return "Yes", extracted_title
                    
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError):
            if attempt < max_retries - 1:
                time.sleep(2)  # Wait before triggering retry sequence
                continue
            else:
                return "Error (Timeout)", L["lbl_na"]
                
    return "Error", L["lbl_na"]

# --- CORE FILE MANAGEMENT LAYERS ---
uploaded_file = st.file_uploader(L["upload_label"], type=["xlsx"], help=L["upload_help"])

if uploaded_file is not None and st.session_state.df_data is None:
    raw_df = pd.read_excel(uploaded_file)
    
    # 1,000 Rows Strict Enforcer Rule
    if len(raw_df) > 1000:
        st.error(L["err_limit"])
    else:
        # Core data tracking keys remain fixed and unchanging in background memory
        raw_df["hkpl_status"] = "Unfinished"
        raw_df["hkpl_title"] = "N/A"
        
        st.session_state.df_data = raw_df
        st.session_state.current_index = 0
        st.session_state.run_state = "IDLE"

# Reset App State Action Handler
if st.session_state.df_data is not None:
    if st.sidebar.button(L["btn_reset"]):
        st.session_state.df_data = None
        st.session_state.current_index = 0
        st.session_state.run_state = "IDLE"
        st.rerun()

# --- RUN STATE ENGINE ---
if st.session_state.df_data is not None:
    df = st.session_state.df_data
    columns = [col for col in df.columns.tolist() if col not in ["hkpl_status", "hkpl_title"]]
    
    if st.session_state.isbn_col_name not in columns:
        st.session_state.isbn_col_name = columns[0]
        
    st.session_state.isbn_col_name = st.selectbox(L["select_isbn"], columns, index=columns.index(st.session_state.isbn_col_name))
    
    # Action Controller Buttons Layout Setup
    st.markdown("### Controls")
    c1, c2, c3 = st.columns(3)
    
    with c1:
        if st.session_state.run_state in ["IDLE", "PAUSED"]:
            btn_label = L["btn_start"] if st.session_state.current_index == 0 else L["btn_resume"]
            if st.button(btn_label, type="primary", use_container_width=True):
                st.session_state.run_state = "RUNNING"
                st.rerun()
    with c2:
        if st.session_state.run_state == "RUNNING":
            if st.button(L["btn_stop"], type="secondary", use_container_width=True):
                st.session_state.run_state = "PAUSED"
                st.rerun()
                
    # Display matching alert banners depending on operational modes
    if st.session_state.run_state == "IDLE":
        st.info(L["status_idle"])
    elif st.session_state.run_state == "RUNNING":
        st.warning(L["status_running"])
    elif st.session_state.run_state == "PAUSED":
        st.info(L["status_paused"])
    elif st.session_state.run_state == "COMPLETED":
        st.success(L["status_completed"])

    # --- PROGRESS CALCULATION ENGINE & METRICS DASHBOARD ---
    total_count = len(df)
    found_count = int((df["hkpl_status"] == "Yes").sum())
    notfound_count = int((df["hkpl_status"] == "No").sum())
    error_count = int(df["hkpl_status"].str.contains("Error", na=False).sum())
    unfinished_count = int((df["hkpl_status"] == "Unfinished").sum())
    
    st.markdown(f"### {L['preview_title']}")
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric(L["m_total"], total_count)
    m2.metric(L["m_found"], found_count)
    m3.metric(L["m_notfound"], notfound_count)
    m4.metric(L["m_error"], error_count)
    m5.metric(L["m_unfinished"], unfinished_count)
    
    # Visible Progress Tracking Bar
    prog_pct = float(st.session_state.current_index / total_count) if total_count > 0 else 0
    st.progress(min(prog_pct, 1.0))

    # Real-Time Data View: Rename background columns to active language headers on the fly
    display_df = df.rename(columns={
        "hkpl_status": L["col_status"],
        "hkpl_title": L["col_title"]
    })
    
    # Map raw value status items for clean rendering in the selected language
    display_df[L["col_status"]] = display_df[L["col_status"]].replace({"Unfinished": L["lbl_unfinished"]})
    
    st.dataframe(display_df, height=350, use_container_width=True)

    # --- HIGHLY CONTROLLABLE STEP LOOP PROCESSING SECTION ---
    if st.session_state.run_state == "RUNNING":
        idx = st.session_state.current_index
        
        if idx < total_count:
            # Skip rows that are already processed (useful for retrying errors)
            if df.at[idx, "hkpl_status"] != "Unfinished":
                st.session_state.current_index += 1
                st.rerun()
            else:
                raw_isbn_val = str(df.iloc[idx][st.session_state.isbn_col_name]).strip()
                clean_isbn = raw_isbn_val.split('.')[0] if '.' in raw_isbn_val else raw_isbn_val
                
                if not clean_isbn or clean_isbn.lower() == 'nan' or len(clean_isbn) < 8:
                    df.at[idx, "hkpl_status"] = "Invalid ISBN"
                    df.at[idx, "hkpl_title"] = L["lbl_na"]
                else:
                    # Execution of the live verification request script block
                    status_res, title_res = check_isbn_on_hkpl(clean_isbn)
                    df.at[idx, "hkpl_status"] = status_res
                    df.at[idx, "hkpl_title"] = title_res
                
                # Increment tracking counters
                st.session_state.current_index += 1
                st.session_state.df_data = df
                
                # Polite scraping delay pause to guard against server resource locks
                time.sleep(1.1)
                st.rerun()
        else:
            st.session_state.run_state = "COMPLETED"
            st.rerun()

    # --- DOWNSTREAM EXCEL EXPORT & RETRY WORKFLOW MANAGEMENT ---
    if st.session_state.run_state in ["PAUSED", "COMPLETED"]:
        st.markdown("---")
        
        col_act1, col_act2 = st.columns(2)
        
        # Action 1: Reconnect & Retry Errors (Only if run is completed and errors exist)
        with col_act1:
            if st.session_state.run_state == "COMPLETED" and error_count > 0:
                if st.button(L["btn_retry_errors"], type="primary"):
                    # Find all error rows and set them back to Unfinished
                    error_mask = df["hkpl_status"].str.contains("Error", na=False)
                    df.loc[error_mask, "hkpl_status"] = "Unfinished"
                    
                    # Reset the index loop back to the beginning so it sweeps through and catches them
                    st.session_state.df_data = df
                    st.session_state.current_index = 0
                    st.session_state.run_state = "RUNNING"
                    st.rerun()
        
        # Action 2: Download the final excel sheet
        with col_act2:
            export_df = df.rename(columns={
                "hkpl_status": L["col_status"],
                "hkpl_title": L["col_title"]
            })
            export_df[L["col_status"]] = export_df[L["col_status"]].replace({"Unfinished": L["lbl_unfinished"]})
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                export_df.to_excel(writer, index=False)
            processed_bytes = output.getvalue()
            
            st.download_button(
                label=L["download_lbl"],
                data=processed_bytes,
                file_name="hkpl_checked_library_inventory.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
