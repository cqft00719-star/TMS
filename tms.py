@st.cache_resource
def get_client():
    try:
        scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
        if "gcp_service_account" not in st.secrets:
            st.error("❌ Secrets 未設定 gcp_service_account，請去 Manage app -> Settings -> Secrets 設定")
            st.stop()
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"❌ Google認證失敗: {e}")
        st.info("請檢查 Streamlit Secrets 格式，特別係 private_key 要有 \\n")
        st.stop()

def get_sh():
    try:
        if "connections" not in st.secrets or "gsheets" not in st.secrets["connections"]:
            st.error("❌ Secrets 未設定 connections.gsheets.spreadsheet")
            st.stop()
        return get_client().open_by_key(st.secrets["connections"]["gsheets"]["spreadsheet"])
    except Exception as e:
        st.error(f"❌ 開唔到Google Sheet: {e}")
        st.stop()
