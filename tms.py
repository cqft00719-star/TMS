import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime
from datetime import timezone, timedelta
import uuid

# --- 香港時間設定 (無需pytz) ---
HKT = timezone(timedelta(hours=8))
def now_hk():
    return datetime.datetime.now(HKT)
def today_hk():
    return now_hk().date()

st.set_page_config(page_title="TMS 運輸系統", layout="wide")

@st.cache_resource
def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    return gspread.authorize(creds)

def get_sh():
    return get_client().open_by_key(st.secrets["connections"]["gsheets"]["spreadsheet"])

@st.cache_data(ttl=5)
def load_df(tab):
    try:
        ws = get_sh().worksheet(tab)
        data = ws.get_all_records()
        if not data:
            return pd.DataFrame()
        return pd.DataFrame(data)
    except Exception as e:
        return pd.DataFrame()

def append_row(tab, row):
    get_sh().worksheet(tab).append_row(row, value_input_option='USER_ENTERED')
    st.cache_data.clear()

def ensure_sheets():
    sh = get_sh()
    titles = [ws.title for ws in sh.worksheets()]
    if "Users" not in titles:
        sh.add_worksheet(title="Users", rows=100, cols=10)
        ws = sh.worksheet("Users")
        ws.update([["username","password","name","role"], ["admin","admin123","Manager","admin"], ["John","John123","John","driver"], ["op01","op01","OP01","op"]])
    if "Jobs" not in titles:
        sh.add_worksheet(title="Jobs", rows=1000, cols=20)
        ws = sh.worksheet("Jobs")
        ws.update([["JobID","Date","Client","From","To","Driver","Status","QTY","Notes","CreatedBy","CreatedAt_HKT"]])

# --- LOGIN - 修復版 ---
if "user" not in st.session_state:
    st.session_state.user = None

if not st.session_state.user:
    st.title("TMS 多人運輸系統 - 登入")
    st.caption(f"香港時間: {now_hk().strftime('%Y-%m-%d %H:%M:%S HKT')}")
    df_u = load_df("Users")
    # 防呆：如果Users空表或欄名錯
    if df_u.empty or "username" not in df_u.columns or "password" not in df_u.columns:
        df_u = pd.DataFrame([
            {"username":"admin","password":"admin123","name":"Manager","role":"admin"},
            {"username":"John","password":"John123","name":"John","role":"driver"},
            {"username":"op01","password":"op01","name":"OP01","role":"op"},
            {"username":"driver01","password":"driver01","name":"Driver01","role":"driver"},
        ])
    df_u["username"] = df_u["username"].astype(str).str.strip()
    df_u["password"] = df_u["password"].astype(str).str.strip()

    u = st.text_input("用戶名 (admin / John / op01)", key="tms_user")
    p = st.text_input("密碼", type="password", key="tms_pass")
    if st.button("登入", type="primary", use_container_width=True):
        if not u or not p:
            st.error("請輸入用戶名同密碼")
        else:
            m = df_u[(df_u["username"]==str(u).strip()) & (df_u["password"]==str(p).strip())]
            if not m.empty:
                st.session_state.user = m.iloc[0].to_dict()
                st.success(f"歡迎 {m.iloc[0]['username']}")
                st.rerun()
            else:
                st.error(f"帳號或密碼錯誤。現有帳號: {', '.join(df_u['username'].tolist())}")
    st.stop()

# --- 登入後 ---
ensure_sheets()
user = st.session_state.user
role = str(user.get("role","driver")).lower()
operator = user.get("username","unknown")

st.sidebar.write(f"👤 {user.get('name', operator)} | {role}")
st.sidebar.write(f"🕐 {now_hk().strftime('%H:%M HKT')}")
if st.sidebar.button("登出"):
    st.session_state.user = None
    st.cache_data.clear()
    st.rerun()

st.title(f"TMS - {role.upper()} 模式")

df_jobs = load_df("Jobs")
df_users = load_df("Users")
driver_list = df_users[df_users["role"].astype(str).str.lower()=="driver"]["username"].astype(str).tolist() if not df_users.empty and "username" in df_users.columns else ["John","driver01"]

if role == "admin" or role == "op":
    tab1, tab2, tab3 = st.tabs(["📋 任務總覽", "➕ 建立新任務", "👥 用戶管理"])

    with tab1:
        st.subheader(f"所有運輸任務 ({len(df_jobs)} 筆)")
        if df_jobs.empty:
            st.info("未有任務")
        else:
            # 篩選
            status_filter = st.selectbox("按狀態篩選", ["全部"] + sorted(df_jobs["Status"].astype(str).unique().tolist()) if "Status" in df_jobs.columns else ["全部"])
            view = df_jobs.copy()
            if status_filter!= "全部" and "Status" in view.columns:
                view = view[view["Status"].astype(str)==status_filter]
            # 司機篩選
            if role == "admin":
                driver_filter = st.selectbox("按司機篩選", ["全部"] + driver_list)
                if driver_filter!= "全部" and "Driver" in view.columns:
                    view = view[view["Driver"].astype(str)==driver_filter]
            st.dataframe(view.sort_index(ascending=False), use_container_width=True, height=500)

            # 更新狀態
            if not view.empty:
                st.markdown("---")
                st.markdown("**快速更新狀態**")
                with st.form("update_status"):
                    job_id = st.selectbox("選擇 JobID", view["JobID"].astype(str).tolist() if "JobID" in view.columns else [])
                    new_status = st.selectbox("新狀態", ["Pending","Assigned","InTransit","Delivered","Completed"])
                    new_driver = st.selectbox("指派司機 (可選)", ["不更改"] + driver_list)
                    if st.form_submit_button("更新"):
                        try:
                            ws = get_sh().worksheet("Jobs")
                            cell = ws.find(job_id)
                            if cell:
                                # Status 係第7欄 (G), Driver係第6欄 (F)
                                ws.update_cell(cell.row, 7, new_status)
                                if new_driver!= "不更改":
                                    ws.update_cell(cell.row, 6, new_driver)
                                st.cache_data.clear()
                                st.success(f"{job_id} 已更新為 {new_status}")
                                st.rerun()
                        except Exception as e:
                            st.error(f"更新失敗: {e}")

    with tab2:
        st.subheader("建立新運輸任務")
        with st.form("create_job", clear_on_submit=True):
            c1, c2 = st.columns(2)
            date = c1.date_input("日期", value=today_hk())
            client = c1.text_input("客戶 *必填")
            from_loc = c1.text_input("起點 From")
            to_loc = c2.text_input("終點 To")
            driver = c2.selectbox("指派司機", [""] + driver_list)
            qty = c2.number_input("數量", min_value=0.0, step=1.0)
            notes = st.text_area("備註")
            if st.form_submit_button("建立任務"):
                if not client:
                    st.error("客戶必填")
                else:
                    job_id = f"JOB-{now_hk().strftime('%Y%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                    status = "Assigned" if driver else "Pending"
                    created_at = now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")
                    append_row("Jobs", [job_id, str(date), client, from_loc, to_loc, driver, status, str(qty), notes, operator, created_at])
                    st.success(f"任務 {job_id} 已建立")
                    st.balloons()
                    st.rerun()

    with tab3:
        st.dataframe(df_users, use_container_width=True)
        with st.form("add_user", clear_on_submit=True):
            c1,c2,c3,c4 = st.columns(4)
            nu = c1.text_input("username *")
            pw = c2.text_input("password *")
            nm = c3.text_input("name")
            rl = c4.selectbox("role", ["admin","op","driver"])
            if st.form_submit_button("新增用戶"):
                if nu and pw:
                    append_row("Users", [nu, pw, nm, rl])
                    st.success(f"{nu} 已新增")
                    st.rerun()

else: # driver 模式
    tab1, tab2 = st.tabs(["🚚 我的任務", "✅ 更新狀態"])
    my_jobs = df_jobs[df_jobs["Driver"].astype(str)==operator] if not df_jobs.empty and "Driver" in df_jobs.columns else pd.DataFrame()

    with tab1:
        st.subheader(f"{operator} 的任務 ({len(my_jobs)} 筆)")
        if my_jobs.empty:
            st.info("未有指派任務")
        else:
            st.dataframe(my_jobs.sort_index(ascending=False), use_container_width=True, height=500)

    with tab2:
        if my_jobs.empty:
            st.info("未有任務可更新")
        else:
            with st.form("driver_update"):
                job_id = st.selectbox("選擇我的 JobID", my_jobs["JobID"].astype(str).tolist() if "JobID" in my_jobs.columns else [])
                new_status = st.selectbox("更新狀態", ["InTransit","Delivered","Completed"])
                if st.form_submit_button("確認更新"):
                    try:
                        ws = get_sh().worksheet("Jobs")
                        cell = ws.find(job_id)
                        if cell:
                            ws.update_cell(cell.row, 7, new_status)
                            st.cache_data.clear()
                            st.success(f"{job_id} -> {new_status} 更新成功 {now_hk().strftime('%H:%M HKT')}")
                            st.rerun()
                    except Exception as e:
                        st.error(f"失敗: {e}")
