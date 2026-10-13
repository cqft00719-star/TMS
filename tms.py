import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime
from datetime import timezone, timedelta
import uuid
import traceback

HKT = timezone(timedelta(hours=8))
def now_hk():
    return datetime.datetime.now(HKT)

st.set_page_config(page_title="TMS 運輸系統", layout="wide")

@st.cache_resource
def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
    if "gcp_service_account" not in st.secrets:
        st.error("❌ Secrets 無 gcp_service_account")
        st.stop()
    try:
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"❌ Google認證失敗: {e}")
        st.code(traceback.format_exc())
        st.stop()

def get_sh():
    try:
        sheet_id = st.secrets["connections"]["gsheets"]["spreadsheet"]
        return get_client().open_by_key(sheet_id)
    except Exception as e:
        st.error(f"❌ 開唔到Google Sheet: {e} | ID: {st.secrets.get('connections',{}).get('gsheets',{}).get('spreadsheet','無ID')}")
        st.info(f"請去 https://docs.google.com/spreadsheets/d/{st.secrets['connections']['gsheets']['spreadsheet']} 按 Share -> 加入 {st.secrets['gcp_service_account']['client_email']} 做 Editor")
        st.code(traceback.format_exc())
        st.stop()

@st.cache_data(ttl=5)
def load_df(tab):
    try:
        ws = get_sh().worksheet(tab)
        data = ws.get_all_records()
        return pd.DataFrame(data) if data else pd.DataFrame()
    except:
        return pd.DataFrame()

def append_row(tab, row):
    get_sh().worksheet(tab).append_row(row, value_input_option='USER_ENTERED')
    st.cache_data.clear()

def ensure_sheets():
    sh = get_sh()
    titles = [ws.title for ws in sh.worksheets()]
    if "Users" not in titles:
        sh.add_worksheet(title="Users", rows=100, cols=10).update([["username","password","name","role"], ["admin","admin123","Manager","admin"], ["John","John123","John","driver"], ["op01","op01","OP01","op"]])
    if "Jobs" not in titles:
        sh.add_worksheet(title="Jobs", rows=1000, cols=20).update([["JobID","Date","Client","From","To","Driver","Status","QTY","Notes","CreatedBy","CreatedAt_HKT"]])

if "user" not in st.session_state:
    st.session_state.user = None

if not st.session_state.user:
    st.title("TMS 多人運輸系統 - 登入")
    st.caption(f"香港時間: {now_hk().strftime('%Y-%m-%d %H:%M:%S HKT')}")
    df_u = load_df("Users")
    if df_u.empty or "username" not in df_u.columns:
        df_u = pd.DataFrame([{"username":"admin","password":"admin123","name":"Manager","role":"admin"}, {"username":"John","password":"John123","name":"John","role":"driver"}, {"username":"op01","password":"op01","name":"OP01","role":"op"}])
    df_u["username"]=df_u["username"].astype(str).str.strip()
    df_u["password"]=df_u["password"].astype(str).str.strip()
    u=st.text_input("用戶名")
    p=st.text_input("密碼", type="password")
    if st.button("登入", type="primary", use_container_width=True):
        m=df_u[(df_u["username"]==str(u).strip()) & (df_u["password"]==str(p).strip())]
        if not m.empty:
            st.session_state.user=m.iloc[0].to_dict()
            st.rerun()
        else:
            st.error(f"帳號錯誤，現有: {', '.join(df_u['username'].tolist())}")
    st.stop()

ensure_sheets()
user=st.session_state.user
role=str(user.get("role","driver")).lower()
operator=user.get("username","")
st.sidebar.write(f"👤 {operator} | {role} | {now_hk().strftime('%H:%M HKT')}")
if st.sidebar.button("登出"):
    st.session_state.user=None; st.cache_data.clear(); st.rerun()
st.title(f"TMS - {role.upper()}")
df_jobs=load_df("Jobs")
df_users=load_df("Users")
driver_list=df_users[df_users["role"].astype(str).str.lower()=="driver"]["username"].astype(str).tolist() if not df_users.empty and "username" in df_users.columns else ["John"]
if role in ["admin","op"]:
    t1,t2=st.tabs(["📋 任務","➕ 建立"])
    with t1:
        st.dataframe(df_jobs.sort_index(ascending=False) if not df_jobs.empty else df_jobs, use_container_width=True)
    with t2:
        with st.form("create", clear_on_submit=True):
            client=st.text_input("客戶 *")
            c1,c2=st.columns(2)
            f=c1.text_input("起點"); t=c2.text_input("終點")
            d=c1.selectbox("司機",[""]+driver_list); q=c2.number_input("數量",0.0,1.0)
            notes=st.text_area("備註")
            if st.form_submit_button("建立"):
                if not client: st.error("客戶必填")
                else:
                    jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                    append_row("Jobs",[jid,str(now_hk().date()),client,f,t,d,"Assigned" if d else "Pending",str(q),notes,operator,now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")])
                    st.success(f"{jid} 已建立"); st.rerun()
else:
    my=df_jobs[df_jobs["Driver"].astype(str)==operator] if not df_jobs.empty and "Driver" in df_jobs.columns else pd.DataFrame()
    st.dataframe(my, use_container_width=True)
