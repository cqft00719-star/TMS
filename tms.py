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

def append_row(tab, row): append_rows(tab, [row])

def ensure_sheets():
    sh = get_sh()
    existing = [ws.title for ws in sh.worksheets()]
    def create_if_not_exist(title, headers, defaults=None):
        if title not in existing:
            ws = sh.add_worksheet(title=title, rows=300, cols=20)
            ws.update('A1', [headers])
            if defaults:
                ws.append_rows(defaults, value_input_option='USER_ENTERED')
            existing.append(title)
    create_if_not_exist("Users", ["username","password","name","role"], [["admin","admin123","Admin","admin"], ["John","John123","John","driver"], ["op01","op01","OP01","op"]])
    create_if_not_exist("Jobs", ["JobID","Date","Client","From","To","Driver","Status","QTY","Notes","CreatedBy","CreatedAt_HKT","POD_Link","PodTime_HKT"])
    create_if_not_exist("Expenses", ["ExpenseID","Date","JobID","Driver","Type","Amount","Remark","CreatedAt_HKT"])
    create_if_not_exist("Config", ["Key","Value"], [["POD_Folder_ID",""]])

if "user" not in st.session_state: st.session_state.user=None
if not st.session_state.user:
    st.title("TMS 多人運輸系統 - 登入")
    df_u=load_df("Users")
    if df_u.empty: df_u=pd.DataFrame([{"username":"admin","password":"admin123","name":"Admin","role":"admin"}])
    u=st.text_input("用戶名"); p=st.text_input("密碼",type="password")
    if st.button("登入",type="primary", use_container_width=True):
        m=df_u[(df_u["username"].astype(str).str.strip()==str(u).strip()) & (df_u["password"].astype(str).str.strip()==str(p).strip())]
        if not m.empty:
            st.session_state.user=m.iloc[0].to_dict(); st.rerun()
        else: st.error("帳號錯")
    st.stop()

ensure_sheets()
user=st.session_state.user
role=str(user.get("role","driver")).lower()
operator=user.get("username","")

st.sidebar.title(f"👤 {user.get('name')} ({role})")
st.sidebar.caption(now_hk().strftime('%Y-%m-%d %H:%M HKT'))
if st.sidebar.button("登出"): st.session_state.user=None; st.cache_data.clear(); st.rerun()

# 分頁定義
if role in ["admin"]:
    tabs = st.tabs(["📊 Dashboard", "📋 任務管理", "📤 上傳Order", "💰 雜費", "👥 用戶"])
elif role == "op":
    tabs = st.tabs(["📋 任務管理", "📤 上傳Order", "💰 雜費"])
else: # driver
    tabs = st.tabs(["🟢 可接任務池", "🚚 我的任務", "📸 上傳POD", "💰 我的雜費"])

df_jobs=load_df("Jobs")
df_exp=load_df("Expenses")

# ---- ADMIN DASHBOARD ----
if role=="admin":
    with tabs[0]:
        st.subheader("Dashboard")
        if df_jobs.empty: st.info("暫無數據")
        else:
            c1,c2,c3,c4=st.columns(4)
            c1.metric("總任務", len(df_jobs))
            c2.metric("待搶", len(df_jobs[(df_jobs["Status"]=="Pending") | (df_jobs["Driver"]=="")]) if "Status" in df_jobs.columns else 0)
            c3.metric("已接單", len(df_jobs[df_jobs["Status"]=="Accepted"]) if "Status" in df_jobs.columns else 0)
            c4.metric("已完成", len(df_jobs[df_jobs["Status"]=="Completed"]) if "Status" in df_jobs.columns else 0)
            st.bar_chart(df_jobs["Status"].value_counts() if "Status" in df_jobs.columns else pd.Series())
            st.dataframe(df_jobs, use_container_width=True)

# ---- 任務管理 (Admin / OP) ----
if role in ["admin","op"]:
    idx = 1 if role=="admin" else 0
    with tabs[idx]:
        st.subheader("所有任務")
        st.dataframe(df_jobs, use_container_width=True, height=400)
        # 建立
        with st.form("create_job", clear_on_submit=True):
            st.write("**手動建立Order**")
            c1,c2=st.columns(2)
            client=c1.text_input("客戶*"); date=c1.date_input("日期", value=today_hk())
            from_loc=c1.text_input("起點"); to_loc=c2.text_input("終點")
            qty=c2.number_input("數量",0.0,1.0); notes=st.text_area("備註")
            if st.form_submit_button("建立Order (待搶)"):
                if not client: st.error("客戶必填")
                else:
                    jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4]}"
                    append_row("Jobs",[jid,str(date),client,from_loc,to_loc,"","Pending",str(qty),notes,operator,now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"),"",""])
                    st.success(f"{jid} 已建立，司機可搶單"); st.rerun()

# ---- 上傳Order (Admin + Operator) ----
if role in ["admin","op"]:
    idx = 2 if role=="admin" else 1
    with tabs[idx]:
        st.subheader("📤 Operator 上傳Order (Excel)")
        st.info("Operator可批量上傳，狀態會設為Pending，司機自行搶單")
        up=st.file_uploader("上傳Excel.xlsx", type=["xlsx"], key="op_upload")
        if up:
            df=pd.read_excel(up)
            st.dataframe(df.head(), use_container_width=True)
            cols=["(不導入)"]+df.columns.tolist()
            mapping={}
            for k in ["Client","From","To","QTY","Date","Notes"]:
                mapping[k]=st.selectbox(f"{k} 對應", cols, key=f"map_{k}")
            if st.button("開始導入 - 設為待搶", type="primary"):
                rows=[]
                for _, r in df.iterrows():
                    def gv(k): return str(r[mapping[k]]) if mapping[k]!="(不導入)" and mapping[k] in df.columns else ""
                    if not gv("Client"): continue
                    jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4]}"
                    rows.append([jid, gv("Date") or str(today_hk()), gv("Client"), gv("From"), gv("To"), "", "Pending", gv("QTY"), gv("Notes"), operator, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"), "", ""])
                if rows:
                    append_rows("Jobs", rows)
                    st.success(f"已導入 {len(rows)} 張Order，司機可搶單"); st.balloons()

# ---- DRIVER: 可接任務池 (搶單) ----
if role=="driver":
    with tabs[0]:
        st.subheader("🟢 可接任務池 - 點擊搶單")
        if df_jobs.empty:
            st.info("暫無可接任務")
        else:
            # Pending 且無司機 或 Driver為空
            avail = df_jobs[(df_jobs["Status"].astype(str)=="Pending") | (df_jobs["Driver"].astype(str)=="")]
            if avail.empty:
                st.success("暫無待搶任務，全部已被接")
            else:
                st.dataframe(avail, use_container_width=True)
                for i, row in avail.iterrows():
                    jid = row["JobID"]
                    with st.container(border=True):
                        c1,c2,c3=st.columns([3,1,1])
                        c1.write(f"**{jid}** | {row.get('Client','')} | {row.get('From','')} -> {row.get('To','')} | QTY:{row.get('QTY','')}")
                        if c3.button(f"搶此單", key=f"claim_{jid}", type="primary"):
                            ws=get_sh().worksheet("Jobs")
                            try:
                                cell=ws.find(jid)
                                # 再查一次是否已被搶
                                current_driver = ws.cell(cell.row, 6).value
                                if current_driver and str(current_driver).strip()!="":
                                    st.error("已被其他司機搶走")
                                else:
                                    ws.update_cell(cell.row, 6, operator) # Driver
                                    ws.update_cell(cell.row, 7, "Accepted") # 已接單
                                    st.cache_data.clear()
                                    st.success(f"{jid} 搶單成功！已自動變更為 已接單")
                                    st.rerun()
                            except Exception as e:
                                st.error(f"搶單失敗: {e}")

# ---- DRIVER: 我的任務 ----
with tabs[1 if role!="driver" else 1]:
    if role=="driver":
        st.subheader("🚚 我的任務 (已接單)")
        my_jobs = df_jobs[df_jobs["Driver"].astype(str)==operator] if not df_jobs.empty and "Driver" in df_jobs.columns else pd.DataFrame()
        if my_jobs.empty:
            st.info("你暫未接單，請去可接任務池搶單")
        else:
            st.dataframe(my_jobs, use_container_width=True)
            # 允許司機自己改狀態為 InTransit / Delivered
            sel=st.selectbox("選擇我的Job改狀態", my_jobs["JobID"].tolist(), key="driver_status_job")
            ns=st.selectbox("更新狀態", ["Accepted-已接單","InTransit-運輸中","Delivered-已送達","Completed-已完成"], key="driver_ns")
            if st.button("更新我的任務狀態"):
                ws=get_sh().worksheet("Jobs")
                cell=ws.find(sel)
                if cell:
                    status_map={"Accepted-已接單":"Accepted","InTransit-運輸中":"InTransit","Delivered-已送達":"Delivered","Completed-已完成":"Completed"}
                    ws.update_cell(cell.row, 7, status_map[ns])
                    st.cache_data.clear(); st.success("已更新"); st.rerun()

# ---- POD上傳 ----
with tabs[2 if role!="driver" else 2]:
    st.subheader("📸 上傳POD")
    # 司機只能上傳自己的，Admin/OP可上傳全部
    if role=="driver":
        my_jobs = df_jobs[df_jobs["Driver"].astype(str)==operator] if not df_jobs.empty else pd.DataFrame()
        # 只顯示已接單的
        my_jobs = my_jobs[my_jobs["Status"].isin(["Accepted","InTransit","Delivered"])] if not my_jobs.empty and "Status" in my_jobs.columns else my_jobs
    else:
        my_jobs = df_jobs

    if my_jobs.empty:
        st.info("無可上傳POD的任務")
    else:
        jid=st.selectbox("選擇JobID", my_jobs["JobID"].astype(str).tolist(), key="pod_job2")
        pod_file=st.file_uploader("拍照或上傳", type=["jpg","png","pdf","jpeg"], key="pod_file2")
        if pod_file and st.button("上傳POD", type="primary", key="btn_pod"):
            try:
                drive=get_drive_service()
                media=MediaIoBaseUpload(io.BytesIO(pod_file.read()), mimetype=pod_file.type, resumable=True)
                file_meta={"name": f"{jid}_POD_{now_hk().strftime('%Y%m%d%H%M')}_{pod_file.name}"}
                cfg=load_df("Config")
                folder_id=""
                if not cfg.empty:
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
                st.success("POD上傳成功，已自動改為已送達")
                st.link_button("查看POD", link)
            except Exception as e:
                st.error(f"上傳失敗: {e}")

# ---- 雜費 ----
with tabs[3 if role in ["admin","op"] else 3]:
    st.subheader("💰 雜費錄入 - 司機可自行輸入")
    st.dataframe(df_exp[df_exp["Driver"].astype(str)==operator] if role=="driver" and not df_exp.empty else df_exp, use_container_width=True)
    with st.form("exp_form2", clear_on_submit=True):
        c1,c2,c3=st.columns(3)
        if role=="driver":
            my_jobs = df_jobs[df_jobs["Driver"].astype(str)==operator] if not df_jobs.empty else pd.DataFrame()
            jid_list=[""]+ (my_jobs["JobID"].tolist() if not my_jobs.empty else [])
        else:
            jid_list=[""]+ (df_jobs["JobID"].tolist() if not df_jobs.empty else [])
        jid=c1.selectbox("關聯JobID", jid_list)
        etype=c2.selectbox("類型", ["隧道費","泊車費","油費","雜費","其他"])
        amt=c3.number_input("金額 $", 0.0, step=1.0)
        remark=st.text_input("備註")
        if st.form_submit_button("提交雜費"):
            eid=f"EXP-{now_hk().strftime('%m%d%H%M')}"
            append_row("Expenses", [eid, str(today_hk()), jid, operator, etype, amt, remark, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")])
            st.success("雜費已提交"); st.rerun()

# 用戶管理 (Admin only)
if role=="admin":
    with tabs[4]:
        st.subheader("用戶管理")
        st.dataframe(load_df("Users"), use_container_width=True)
        with st.form("add_user", clear_on_submit=True):
            nu=st.text_input("用戶名"); np=st.text_input("密碼"); nn=st.text_input("姓名"); nr=st.selectbox("角色",["driver","op","admin"])
            if st.form_submit_button("新增"):
                append_row("Users", [nu,np,nn,nr]); st.success("已新增"); st.rerun()
