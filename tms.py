import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import datetime
from datetime import timezone, timedelta
import uuid
import io

HKT = timezone(timedelta(hours=8))
def now_hk(): return datetime.datetime.now(HKT)
def today_hk(): return now_hk().date()

st.set_page_config(page_title="TMS 運輸系統", layout="wide", page_icon="🚚")

@st.cache_resource
def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    return gspread.authorize(creds), creds

def get_sh():
    client, _ = get_client()
    return client.open_by_key(st.secrets["connections"]["gsheets"]["spreadsheet"])

def get_drive_service():
    _, creds = get_client()
    return build('drive', 'v3', credentials=creds)

@st.cache_data(ttl=10)
def load_df(tab):
    try:
        ws = get_sh().worksheet(tab)
        return pd.DataFrame(ws.get_all_records())
    except:
        return pd.DataFrame()

def append_rows(tab, rows):
    get_sh().worksheet(tab).append_rows(rows, value_input_option='USER_ENTERED')
    st.cache_data.clear()

def append_row(tab, row):
    append_rows(tab, [row])

def ensure_sheets():
    sh = get_sh()
    existing = [ws.title for ws in sh.worksheets()]

    def create_if_not_exist(title, headers, defaults=None):
        if title not in existing:
            ws = sh.add_worksheet(title=title, rows=200, cols=20)
            ws.update('A1', [headers])
            if defaults:
                ws.append_rows(defaults, value_input_option='USER_ENTERED')
            existing.append(title)

    create_if_not_exist("Users", ["username","password","name","role"], [["admin","admin123","Admin","admin"], ["John","John123","John","driver"], ["op01","op01","OP01","op"]])
    create_if_not_exist("Jobs", ["JobID","Date","Client","From","To","Driver","Status","QTY","Notes","CreatedBy","CreatedAt_HKT","POD_Link","PodTime_HKT"])
    create_if_not_exist("Expenses", ["ExpenseID","Date","JobID","Driver","Type","Amount","Remark","CreatedAt_HKT"])
    create_if_not_exist("Config", ["Key","Value"], [["POD_Folder_ID",""]])

# --- 登入 ---
if "user" not in st.session_state: st.session_state.user=None
if not st.session_state.user:
    st.title("TMS 多人運輸系統 - 登入")
    st.caption(f"香港時間: {now_hk().strftime('%Y-%m-%d %H:%M HKT')}")
    df_u=load_df("Users")
    if df_u.empty:
        df_u=pd.DataFrame([{"username":"admin","password":"admin123","name":"Admin","role":"admin"}, {"username":"John","password":"John123","name":"John","role":"driver"}])
    u=st.text_input("用戶名"); p=st.text_input("密碼",type="password")
    if st.button("登入",type="primary",use_container_width=True):
        m=df_u[(df_u["username"].astype(str).str.strip()==str(u).strip()) & (df_u["password"].astype(str).str.strip()==str(p).strip())]
        if not m.empty:
            st.session_state.user=m.iloc[0].to_dict(); st.rerun()
        else:
            st.error(f"帳號密碼錯，現有: {', '.join(df_u['username'].astype(str).tolist())}")
    st.stop()

ensure_sheets()
user=st.session_state.user
role=str(user.get("role","driver")).lower()
operator=user.get("username","")
st.sidebar.title(f"👤 {user.get('name')} ({role})")
st.sidebar.caption(f"{now_hk().strftime('%Y-%m-%d %H:%M HKT')}")
if st.sidebar.button("登出"): st.session_state.user=None; st.cache_data.clear(); st.rerun()

# --- 功能分頁 ---
if role in ["admin","op"]:
    tabs = st.tabs(["📊 Dashboard", "📋 任務管理", "📤 Excel批量導入", "💰 雜費", "👥 用戶"])
else:
    tabs = st.tabs(["🚚 我的任務", "📸 上傳POD", "💰 我的雜費"])

df_jobs=load_df("Jobs")
df_exp=load_df("Expenses")

# 1. Dashboard
if role in ["admin","op"]:
    with tabs[0]:
        st.subheader("Dashboard")
        if df_jobs.empty:
            st.info("暫無數據，請去Excel導入或手動建立")
        else:
            c1,c2,c3,c4=st.columns(4)
            c1.metric("總任務", len(df_jobs))
            c2.metric("待派", len(df_jobs[df_jobs["Status"]=="Pending"]) if "Status" in df_jobs.columns else 0)
            c3.metric("運輸中", len(df_jobs[df_jobs["Status"]=="InTransit"]) if "Status" in df_jobs.columns else 0)
            c4.metric("已完成", len(df_jobs[df_jobs["Status"]=="Completed"]) if "Status" in df_jobs.columns else 0)
            if "Status" in df_jobs.columns:
                st.bar_chart(df_jobs["Status"].value_counts())
            if "Driver" in df_jobs.columns and not df_jobs["Driver"].empty:
                st.write("司機任務分佈")
                st.bar_chart(df_jobs["Driver"].value_counts())
            st.dataframe(df_jobs.sort_values("Date", ascending=False) if "Date" in df_jobs.columns else df_jobs, use_container_width=True)

# 2. 任務管理
with tabs[1 if role in ["admin","op"] else 0]:
    st.subheader("任務列表")
    st.dataframe(df_jobs, use_container_width=True, height=400)
    if role in ["admin","op"] and not df_jobs.empty:
        with st.expander("更新狀態 / 指派司機"):
            job_id=st.selectbox("JobID", df_jobs["JobID"].astype(str).tolist(), key="upd_job")
            df_users=load_df("Users")
            driver_list=df_users[df_users["role"].astype(str).str.lower()=="driver"]["username"].tolist() if not df_users.empty and "username" in df_users.columns else ["John"]
            ns=st.selectbox("狀態", ["Pending","Assigned","InTransit","Delivered","Completed"], key="upd_stat")
            nd=st.selectbox("司機", ["不更改"]+driver_list, key="upd_driver")
            if st.button("確認更新"):
                ws=get_sh().worksheet("Jobs")
                cell=ws.find(job_id)
                if cell:
                    ws.update_cell(cell.row, 7, ns)
                    if nd!="不更改": ws.update_cell(cell.row, 6, nd)
                    st.cache_data.clear(); st.success("已更新"); st.rerun()
    if role in ["admin","op"]:
        with st.form("create_job", clear_on_submit=True):
            st.write("**手動建立任務**")
            c1,c2=st.columns(2)
            client=c1.text_input("客戶*"); date=c1.date_input("日期", value=today_hk())
            from_loc=c1.text_input("起點"); to_loc=c2.text_input("終點")
            df_users=load_df("Users")
            driver_list=df_users[df_users["role"].astype(str).str.lower()=="driver"]["username"].tolist() if not df_users.empty and "username" in df_users.columns else ["John"]
            driver=c2.selectbox("司機", [""]+driver_list); qty=c2.number_input("數量",0.0,1.0)
            notes=st.text_area("備註")
            if st.form_submit_button("建立"):
                if not client: st.error("客戶必填")
                else:
                    jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4]}"
                    append_row("Jobs",[jid,str(date),client,from_loc,to_loc,driver,"Assigned" if driver else "Pending",str(qty),notes,operator,now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"),"",""])
                    st.success(jid); st.rerun()

# 3. Excel批量導入
if role in ["admin","op"]:
    with tabs[2]:
        st.subheader("📤 Excel批量導入Order")
        st.caption("支援欄位: Client, From, To, QTY, Date, Driver, Notes | 第一行做表頭")
        up=st.file_uploader("上傳Excel (.xlsx)", type=["xlsx"])
        if up:
            df=pd.read_excel(up)
            st.dataframe(df.head(), use_container_width=True)
            mapping={}
            cols=["(不導入)"]+df.columns.tolist()
            c1,c2=st.columns(2)
            for k in ["Client","From","To","QTY","Date","Driver","Notes"]:
                mapping[k]=c1.selectbox(f"{k} 對應Excel欄", cols, key=f"map_{k}") if k in ["Client","From","To"] else c2.selectbox(f"{k} 對應Excel欄", cols, key=f"map_{k}")
            if st.button("開始導入", type="primary"):
                rows=[]
                for _, r in df.iterrows():
                    def gv(k): return str(r[mapping[k]]) if mapping[k]!="(不導入)" and mapping[k] in df.columns else ""
                    if not gv("Client"): continue
                    jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4]}"
                    rows.append([jid, gv("Date") or str(today_hk()), gv("Client"), gv("From"), gv("To"), gv("Driver"), "Assigned" if gv("Driver") else "Pending", gv("QTY"), gv("Notes"), operator, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"), "", ""])
                if rows:
                    append_rows("Jobs", rows)
                    st.success(f"已導入 {len(rows)} 行"); st.balloons()
                else:
                    st.error("無有效行，請檢查Client欄有無對應")

# 4. POD上傳
with tabs[2 if role in ["admin","op"] else 1]:
    st.subheader("📸 上傳POD簽收")
    my_jobs = df_jobs if role in ["admin","op"] else df_jobs[df_jobs["Driver"].astype(str)==operator] if not df_jobs.empty and "Driver" in df_jobs.columns else pd.DataFrame()
    if my_jobs.empty:
        st.info("無任務")
    else:
        jid=st.selectbox("選擇JobID", my_jobs["JobID"].astype(str).tolist(), key="pod_job")
        pod_file=st.file_uploader("影相或上傳POD", type=["jpg","png","pdf"], key="pod_file")
        if pod_file and st.button("上傳POD", type="primary"):
            try:
                drive=get_drive_service()
                media=MediaIoBaseUpload(io.BytesIO(pod_file.read()), mimetype=pod_file.type, resumable=True)
                file_meta={"name": f"{jid}_POD_{now_hk().strftime('%Y%m%d%H%M')}_{pod_file.name}"}
                cfg=load_df("Config")
                folder_id=""
                if not cfg.empty and "Value" in cfg.columns:
                    try: folder_id=cfg[cfg["Key"]=="POD_Folder_ID"]["Value"].values[0]
                    except: pass
                if folder_id and len(str(folder_id))>10: file_meta["parents"]=[str(folder_id).strip()]
                uploaded=drive.files().create(body=file_meta, media_body=media, fields='id, webViewLink').execute()
                link=uploaded.get("webViewLink","")
                ws=get_sh().worksheet("Jobs")
                cell=ws.find(jid)
                if cell:
                    ws.update_cell(cell.row, 12, link)
                    ws.update_cell(cell.row, 13, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"))
                    ws.update_cell(cell.row, 7, "Delivered")
                st.cache_data.clear()
                st.success(f"上傳成功")
                st.link_button("查看POD", link)
            except Exception as e:
                st.error(f"上傳失敗: {e}")

# 5. 雜費
with tabs[3 if role in ["admin","op"] else 2]:
    st.subheader("💰 雜費錄入")
    st.dataframe(df_exp, use_container_width=True)
    with st.form("exp_form", clear_on_submit=True):
        c1,c2,c3=st.columns(3)
        jid=c1.selectbox("關聯JobID", [""]+ (df_jobs["JobID"].astype(str).tolist() if not df_jobs.empty else []))
        etype=c2.selectbox("類型", ["隧道費","泊車費","油費","雜費","其他"])
        amt=c3.number_input("金額 $", min_value=0.0, step=1.0)
        remark=st.text_input("備註")
        if st.form_submit_button("提交雜費"):
            eid=f"EXP-{now_hk().strftime('%m%d%H%M')}"
            append_row("Expenses", [eid, str(today_hk()), jid, operator, etype, amt, remark, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")])
            st.success("已提交"); st.rerun()

# 用戶管理
if role in ["admin","op"]:
    with tabs[4]:
        st.subheader("用戶管理")
        st.dataframe(load_df("Users"), use_container_width=True)
        with st.form("add_user", clear_on_submit=True):
            nu=st.text_input("新用戶名"); np=st.text_input("密碼"); nn=st.text_input("姓名"); nr=st.selectbox("角色",["driver","op","admin"])
            if st.form_submit_button("新增用戶"):
                append_row("Users", [nu,np,nn,nr]); st.success("已新增"); st.rerun()
