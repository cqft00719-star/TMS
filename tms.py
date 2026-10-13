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

def get_template_excel():
    df_template = pd.DataFrame([
        {"Client":"ABC公司","From":"葵涌","To":"中環","QTY":2,"Date":str(today_hk()),"Notes":"易碎"},
        {"Client":"XYZ公司","From":"觀塘","To":"沙田","QTY":1,"Date":str(today_hk()),"Notes":""},
    ])
    df_template = df_template[["Client","From","To","QTY","Date","Notes"]]
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_template.to_excel(writer, index=False, sheet_name='OrderTemplate')
        pd.DataFrame([{"欄位":"Client","必填":"是","說明":"客戶名"}, {"欄位":"From","必填":"否","說明":"起點"}, {"欄位":"To","必填":"否","說明":"終點"}, {"欄位":"QTY","必填":"否","說明":"數量"}, {"欄位":"Date","必填":"否","說明":"YYYY-MM-DD"}, {"欄位":"Notes","必填":"否","說明":"備註"}]).to_excel(writer, index=False, sheet_name='說明')
    output.seek(0)
    return output.getvalue()

def extract_folder_id_from_url(s):
    if not s: return ""
    s=str(s).strip()
    if "folders/" in s:
        try: return s.split("folders/")[1].split("?")[0].split("/")[0].split("&")[0]
        except: pass
    if "/d/" in s and "google.com" in s:
        try: return s.split("/d/")[1].split("/")[0]
        except: pass
    return s

def ensure_pod_folder():
    drive = get_drive_service()
    cfg = load_df("Config")
    folder_id = ""
    if not cfg.empty and "Value" in cfg.columns:
        try:
            v = cfg[cfg["Key"]=="POD_Folder_ID"]["Value"].values
            if len(v)>0: folder_id = extract_folder_id_from_url(v[0])
        except: pass
    if folder_id and len(folder_id)>10:
        try:
            drive.files().get(fileId=folder_id, fields='id', supportsAllDrives=True).execute()
            return folder_id
        except:
            folder_id=""
    try:
        res = drive.files().list(q="name='TMS_POD' and mimeType='application/vnd.google-apps.folder' and trashed=false", fields="files(id,name)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
        if res.get("files"):
            folder_id = res["files"][0]["id"]
        else:
            file_meta = {"name": "TMS_POD", "mimeType": "application/vnd.google-apps.folder"}
            folder = drive.files().create(body=file_meta, fields='id', supportsAllDrives=True).execute()
            folder_id = folder.get("id")
            try: drive.permissions().create(fileId=folder_id, body={"role":"reader","type":"anyone"}, supportsAllDrives=True).execute()
            except: pass
        try:
            ws = get_sh().worksheet("Config")
            cell = ws.find("POD_Folder_ID")
            if cell: ws.update_cell(cell.row, 2, folder_id)
            else: ws.update_cell(2, 2, folder_id)
        except: pass
        return folder_id
    except Exception as e:
        st.warning(f"自動建Folder失敗: {e}")
        return ""

if "user" not in st.session_state: st.session_state.user=None
if not st.session_state.user:
    st.title("TMS 多人運輸系統 - 登入")
    st.caption(f"香港時間: {now_hk().strftime('%Y-%m-%d %H:%M HKT')}")
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

if role in ["admin"]:
    tabs = st.tabs(["📊 Dashboard", "📋 任務管理", "📤 上傳Order", "📸 上傳POD", "💰 雜費", "👥 用戶"])
elif role == "op":
    tabs = st.tabs(["📋 任務管理", "📤 上傳Order", "📸 上傳POD", "💰 雜費"])
else:
    tabs = st.tabs(["🟢 可接任務池", "🚚 我的任務", "📸 上傳POD", "💰 我的雜費"])

df_jobs=load_df("Jobs")
df_exp=load_df("Expenses")

if role=="admin":
    with tabs[0]:
        st.subheader("Dashboard")
        if df_jobs.empty: st.info("暫無數據")
        else:
            c1,c2,c3,c4=st.columns(4)
            c1.metric("總任務", len(df_jobs))
            c2.metric("待搶", len(df_jobs[df_jobs["Status"]=="Pending"]) if "Status" in df_jobs.columns else 0)
            c3.metric("已接單", len(df_jobs[df_jobs["Status"]=="Accepted"]) if "Status" in df_jobs.columns else 0)
            c4.metric("已完成", len(df_jobs[df_jobs["Status"]=="Completed"]) if "Status" in df_jobs.columns else 0)
            st.bar_chart(df_jobs["Status"].value_counts() if "Status" in df_jobs.columns else pd.Series())
            st.dataframe(df_jobs, use_container_width=True)

if role in ["admin","op"]:
    idx = 1 if role=="admin" else 0
    with tabs[idx]:
        st.subheader("所有任務")
        st.dataframe(df_jobs, use_container_width=True, height=350)

if role in ["admin","op"]:
    idx = 2 if role=="admin" else 1
    with tabs[idx]:
        st.subheader("📤 上傳Order - 手動 + Excel + Template")
        c1,c2=st.columns([1,2])
        with c1:
            st.download_button(label="📥 下載Order模板.xlsx", data=get_template_excel(), file_name=f"TMS_Order_Template_{today_hk()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
        with c2: st.info("欄位：Client(必填), From, To, QTY, Date, Notes")
        st.divider()
        st.write("**手動加入Order**")
        with st.form("manual_order", clear_on_submit=True):
            c1,c2,c3=st.columns(3)
            client=c1.text_input("客戶*"); date=c1.date_input("日期", value=today_hk())
            from_loc=c2.text_input("起點"); to_loc=c2.text_input("終點")
            qty=c3.number_input("數量", 0.0, 100.0, 1.0); notes=c3.text_input("備註")
            if st.form_submit_button("✍️ 手動建立", use_container_width=True):
                if not client.strip(): st.error("客戶必填")
                else:
                    jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4]}"
                    append_row("Jobs",[jid,str(date),client.strip(),from_loc,to_loc,"","Pending",str(qty),notes,operator,now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"),"",""])
                    st.success(f"已建立 {jid}"); st.rerun()
        st.divider()
        st.write("**Excel批量上傳**")
        up=st.file_uploader("上傳Excel", type=["xlsx"], key="op_upload")
        if up:
            df=pd.read_excel(up)
            st.dataframe(df.head(), use_container_width=True)
            cols=["(不導入)"]+df.columns.tolist()
            mapping={}
            cc1,cc2,cc3=st.columns(3)
            mapping["Client"]=cc1.selectbox("Client*", cols, key="map_Client")
            mapping["From"]=cc1.selectbox("From", cols, key="map_From")
            mapping["To"]=cc2.selectbox("To", cols, key="map_To")
            mapping["QTY"]=cc2.selectbox("QTY", cols, key="map_QTY")
            mapping["Date"]=cc3.selectbox("Date", cols, key="map_Date")
            mapping["Notes"]=cc3.selectbox("Notes", cols, key="map_Notes")
            if st.button("🚀 開始導入", type="primary", use_container_width=True):
                rows=[]
                for _, r in df.iterrows():
                    def gv(k):
                        v = r[mapping[k]] if mapping[k]!="(不導入)" and mapping[k] in df.columns else ""
                        return str(v).strip() if pd.notna(v) else ""
                    if not gv("Client"): continue
                    jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4]}"
                    rows.append([jid, gv("Date") or str(today_hk()), gv("Client"), gv("From"), gv("To"), "", "Pending", gv("QTY"), gv("Notes"), operator, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"), "", ""])
                if rows:
                    append_rows("Jobs", rows); st.success(f"成功導入 {len(rows)} 張"); st.balloons(); st.rerun()

if role=="driver":
    with tabs[0]:
        st.subheader("🟢 可接任務池")
        if df_jobs.empty: st.info("暫無可接任務")
        else:
            avail = df_jobs[(df_jobs["Status"].astype(str)=="Pending") | (df_jobs["Driver"].astype(str)=="")]
            if avail.empty: st.success("暫無待搶")
            else:
                st.dataframe(avail, use_container_width=True)
                for _, row in avail.iterrows():
                    jid=row["JobID"]
                    with st.container(border=True):
                        c1,c2=st.columns([4,1])
                        c1.write(f"**{jid}** | {row.get('Client','')} | {row.get('From','')}->{row.get('To','')}")
                        if c2.button("搶此單", key=f"claim_{jid}", type="primary"):
                            ws=get_sh().worksheet("Jobs"); cell=ws.find(jid)
                            if ws.cell(cell.row, 6).value not in [None,""," "]: st.error("已被搶")
                            else:
                                ws.update_cell(cell.row, 6, operator); ws.update_cell(cell.row, 7, "Accepted")
                                st.cache_data.clear(); st.success(f"{jid} 搶單成功"); st.rerun()

with tabs[1 if role!="driver" else 1]:
    if role=="driver":
        st.subheader("🚚 我的任務")
        my_jobs = df_jobs[df_jobs["Driver"].astype(str)==operator] if not df_jobs.empty and "Driver" in df_jobs.columns else pd.DataFrame()
        if my_jobs.empty: st.info("未接單")
        else:
            st.dataframe(my_jobs, use_container_width=True)
            sel=st.selectbox("選擇我的Job改狀態", my_jobs["JobID"].tolist(), key="driver_status_job")
            ns=st.selectbox("更新狀態", ["Accepted-已接單","InTransit-運輸中","Delivered-已送達","Completed-已完成"], key="driver_ns")
            if st.button("更新狀態"):
                ws=get_sh().worksheet("Jobs"); cell=ws.find(sel)
                if cell:
                    mp={"Accepted-已接單":"Accepted","InTransit-運輸中":"InTransit","Delivered-已送達":"Delivered","Completed-已完成":"Completed"}
                    ws.update_cell(cell.row, 7, mp[ns]); st.cache_data.clear(); st.success("已更新"); st.rerun()

with tabs[3 if role=="admin" else 2 if role=="op" else 2]:
    st.subheader("📸 上傳POD - 自動入Google Drive")
    my_jobs_pool = df_jobs if role in ["admin","op"] else df_jobs[df_jobs["Driver"].astype(str)==operator] if not df_jobs.empty else pd.DataFrame()
    if role=="driver" and not my_jobs_pool.empty and "Status" in my_jobs_pool.columns:
        my_jobs_pool = my_jobs_pool[my_jobs_pool["Status"].isin(["Accepted","InTransit","Delivered"])]
    if my_jobs_pool.empty: st.info("無任務可上傳")
    else:
        jid=st.selectbox("選擇JobID", my_jobs_pool["JobID"].astype(str).tolist(), key="pod_job_final")
        pod_file=st.file_uploader("拍照或上傳", type=["jpg","png","pdf","jpeg"], key="pod_file_final")
        if pod_file: st.image(pod_file, caption="預覽", width=350)
        if pod_file and st.button("上傳POD", type="primary", key="btn_pod_final"):
            try:
                with st.spinner("上傳到Google Drive中..."):
                    drive=get_drive_service()
                    folder_id = ensure_pod_folder()
                    if not folder_id or len(folder_id)<10:
                        st.error(f"搵唔到TMS_POD Folder，請確認已Share俾 {st.secrets['gcp_service_account']['client_email']}")
                        st.stop()
                    media=MediaIoBaseUpload(io.BytesIO(pod_file.getvalue()), mimetype=pod_file.type, resumable=True)
                    file_meta={"name": f"{jid}_POD_{now_hk().strftime('%Y%m%d%H%M')}_{pod_file.name}", "parents":[folder_id]}
                    uploaded=drive.files().create(body=file_meta, media_body=media, fields='id, webViewLink', supportsAllDrives=True).execute()
                    try:
                        drive.permissions().create(fileId=uploaded["id"], body={"role":"reader","type":"anyone"}, supportsAllDrives=True).execute()
                    except: pass
                    link=uploaded.get("webViewLink","")
                    ws=get_sh().worksheet("Jobs"); cell=ws.find(jid)
                    if cell:
                        ws.update_cell(cell.row, 12, link)
                        ws.update_cell(cell.row, 13, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"))
                        ws.update_cell(cell.row, 7, "Delivered")
                    st.cache_data.clear()
                    st.success("✅ 上傳成功，已自動入TMS_POD")
                    st.link_button("🔗 查看相片", link)
                    st.link_button("📁 打開Folder", f"https://drive.google.com/drive/folders/{folder_id}")
            except Exception as e:
                st.error(f"上傳失敗: {e}")
                import traceback; st.code(traceback.format_exc())

with tabs[4 if role=="admin" else 3 if role=="op" else 3]:
    st.subheader("💰 雜費錄入")
    st.dataframe(df_exp[df_exp["Driver"].astype(str)==operator] if role=="driver" and not df_exp.empty else df_exp, use_container_width=True)
    with st.form("exp_form2", clear_on_submit=True):
        c1,c2,c3=st.columns(3)
        jid_list=[""]+ (df_jobs[df_jobs["Driver"].astype(str)==operator]["JobID"].tolist() if role=="driver" and not df_jobs.empty else df_jobs["JobID"].tolist() if not df_jobs.empty else [])
        jid=c1.selectbox("關聯JobID", jid_list)
        etype=c2.selectbox("類型", ["隧道費","泊車費","油費","雜費","其他"])
        amt=c3.number_input("金額 $", 0.0, step=1.0); remark=st.text_input("備註")
        if st.form_submit_button("提交雜費"):
            eid=f"EXP-{now_hk().strftime('%m%d%H%M')}"
            append_row("Expenses", [eid, str(today_hk()), jid, operator, etype, amt, remark, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")])
            st.success("已提交"); st.rerun()

if role=="admin":
    with tabs[5]:
        st.subheader("用戶管理"); st.dataframe(load_df("Users"), use_container_width=True)
        with st.form("add_user", clear_on_submit=True):
            nu=st.text_input("用戶名"); np=st.text_input("密碼"); nn=st.text_input("姓名"); nr=st.selectbox("角色",["driver","op","admin"])
            if st.form_submit_button("新增"):
                append_row("Users", [nu,np,nn,nr]); st.success("已新增"); st.rerun()
