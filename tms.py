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
import base64

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
    # 加多個POD_Image欄位俾Base64備份
    create_if_not_exist("Jobs", ["JobID","Date","Client","From","To","Driver","Status","QTY","Notes","CreatedBy","CreatedAt_HKT","POD_Link","PodTime_HKT","POD_Image"])
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
        if not df_jobs.empty:
            c1,c2,c3,c4=st.columns(4)
            c1.metric("總任務", len(df_jobs)); c2.metric("待搶", len(df_jobs[df_jobs["Status"]=="Pending"]) if "Status" in df_jobs.columns else 0)
            c3.metric("已接單", len(df_jobs[df_jobs["Status"]=="Accepted"]) if "Status" in df_jobs.columns else 0); c4.metric("已完成", len(df_jobs[df_jobs["Status"]=="Completed"]) if "Status" in df_jobs.columns else 0)
            st.dataframe(df_jobs, use_container_width=True)

if role in ["admin","op"]:
    idx = 1 if role=="admin" else 0
    with tabs[idx]:
        st.subheader("所有任務")
        st.dataframe(df_jobs, use_container_width=True, height=350)
        # 顯示POD相
        if not df_jobs.empty and "POD_Image" in df_jobs.columns:
            for _, r in df_jobs.iterrows():
                if str(r.get("POD_Image","")).startswith("data:image"):
                    with st.expander(f"{r['JobID']} POD相"):
                        st.image(r["POD_Image"])

if role in ["admin","op"]:
    idx = 2 if role=="admin" else 1
    with tabs[idx]:
        st.subheader("📤 上傳Order")
        st.download_button("📥 下載模板", data=get_template_excel(), file_name="TMS_Template.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with st.form("manual_order", clear_on_submit=True):
            c1,c2,c3=st.columns(3)
            client=c1.text_input("客戶*"); date=c1.date_input("日期", value=today_hk())
            from_loc=c2.text_input("起點"); to_loc=c2.text_input("終點")
            qty=c3.number_input("數量", 0.0, 100.0, 1.0); notes=c3.text_input("備註")
            if st.form_submit_button("✍️ 建立"):
                if client.strip():
                    jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4]}"
                    append_row("Jobs",[jid,str(date),client.strip(),from_loc,to_loc,"","Pending",str(qty),notes,operator,now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"),"","",""])
                    st.success(f"已建立 {jid}"); st.rerun()
        up=st.file_uploader("Excel批量上傳", type=["xlsx"])
        if up:
            df=pd.read_excel(up); st.dataframe(df.head())
            if st.button("🚀 導入"):
                rows=[]
                for _, r in df.iterrows():
                    if str(r.iloc[0]).strip()=="": continue
                    jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4]}"
                    rows.append([jid, str(today_hk()), str(r.iloc[0]), str(r.iloc[1]) if len(r)>1 else "", str(r.iloc[2]) if len(r)>2 else "", "", "Pending", "1", "", operator, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"), "", "", ""])
                if rows: append_rows("Jobs", rows); st.success(f"導入{len(rows)}"); st.rerun()

if role=="driver":
    with tabs[0]:
        st.subheader("🟢 可接任務池")
        avail = df_jobs[(df_jobs["Status"].astype(str)=="Pending") | (df_jobs["Driver"].astype(str)=="")] if not df_jobs.empty else pd.DataFrame()
        if avail.empty: st.info("暫無待搶")
        else:
            st.dataframe(avail, use_container_width=True)
            for _, row in avail.iterrows():
                jid=row["JobID"]
                if st.button(f"搶 {jid}", key=f"claim_{jid}", type="primary"):
                    ws=get_sh().worksheet("Jobs"); cell=ws.find(jid)
                    if ws.cell(cell.row, 6).value not in [None,""," "]: st.error("已被搶")
                    else:
                        ws.update_cell(cell.row, 6, operator); ws.update_cell(cell.row, 7, "Accepted")
                        st.cache_data.clear(); st.rerun()

with tabs[1 if role!="driver" else 1]:
    if role=="driver":
        st.subheader("🚚 我的任務")
        my_jobs = df_jobs[df_jobs["Driver"].astype(str)==operator] if not df_jobs.empty else pd.DataFrame()
        st.dataframe(my_jobs, use_container_width=True)

# ===== 核心修復：POD上傳 =====
with tabs[3 if role=="admin" else 2 if role=="op" else 2]:
    st.subheader("📸 上傳POD")
    my_jobs_pool = df_jobs if role in ["admin","op"] else df_jobs[df_jobs["Driver"].astype(str)==operator] if not df_jobs.empty else pd.DataFrame()
    if my_jobs_pool.empty: st.info("無任務")
    else:
        jid=st.selectbox("選擇JobID", my_jobs_pool["JobID"].astype(str).tolist(), key="pod_job_v11")
        pod_file=st.file_uploader("拍照或上傳", type=["jpg","png","pdf","jpeg"], key="pod_file_v11")
        if pod_file: st.image(pod_file, width=350)
        if pod_file and st.button("上傳POD", type="primary"):
            try:
                ws=get_sh().worksheet("Jobs"); cell=ws.find(jid)
                # 1. 先試Drive
                try:
                    cfg = load_df("Config")
                    fid = ""
                    if not cfg.empty:
                        try: fid = extract_folder_id_from_url(cfg[cfg["Key"]=="POD_Folder_ID"]["Value"].values[0])
                        except: pass
                    drive=get_drive_service()
                    media=MediaIoBaseUpload(io.BytesIO(pod_file.getvalue()), mimetype=pod_file.type, resumable=True)
                    meta={"name": f"{jid}_POD_{now_hk().strftime('%m%d%H%M')}_{pod_file.name}"}
                    if fid: meta["parents"]=[fid]
                    uploaded=drive.files().create(body=meta, media_body=media, fields='id, webViewLink', supportsAllDrives=True).execute()
                    try: drive.permissions().create(fileId=uploaded["id"], body={"role":"reader","type":"anyone"}, supportsAllDrives=True).execute()
                    except: pass
                    link=uploaded.get("webViewLink","")
                    if cell:
                        ws.update_cell(cell.row, 12, link)
                        ws.update_cell(cell.row, 13, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"))
                        ws.update_cell(cell.row, 7, "Delivered")
                    st.cache_data.clear()
                    st.success("✅ 上傳到Google Drive成功")
                    st.link_button("🔗 查看相片", link)
                    st.stop()
                except Exception as e:
                    err=str(e)
                    if "storage quota" in err or "403" in err:
                        st.warning("Drive免費帳號quota限制，自動改為存入Google Sheet，App一樣睇到相")
                        # 2. Fallback: 存Base64入Sheet
                        b64 = base64.b64encode(pod_file.getvalue()).decode()
                        data_url = f"data:{pod_file.type};base64,{b64}"
                        # Base64太長，分段存？Sheet cell上限50000字符，壓縮圖片
                        # 先壓縮
                        try:
                            from PIL import Image
                            img = Image.open(io.BytesIO(pod_file.getvalue()))
                            img.thumbnail((800,800))
                            buf=io.BytesIO(); img.save(buf, format="JPEG", quality=70)
                            b64 = base64.b64encode(buf.getvalue()).decode()
                            data_url = f"data:image/jpeg;base64,{b64}"
                        except: pass
                        if cell:
                            ws.update_cell(cell.row, 12, "BASE64_SAVED")
                            ws.update_cell(cell.row, 13, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"))
                            ws.update_cell(cell.row, 7, "Delivered")
                            try: ws.update_cell(cell.row, 14, data_url[:49000]) # 14 = N欄 POD_Image
                            except: ws.update_cell(cell.row, 14, data_url[:49000])
                        st.cache_data.clear()
                        st.success("✅ POD已保存到Google Sheet，唔使Drive都睇到")
                        st.image(data_url, caption="已保存POD")
                    else:
                        raise e
            except Exception as e:
                st.error(f"上傳失敗: {e}")
                import traceback; st.code(traceback.format_exc())

with tabs[4 if role=="admin" else 3 if role=="op" else 3]:
    st.subheader("💰 雜費")
    st.dataframe(df_exp, use_container_width=True)

if role=="admin":
    with tabs[5]:
        st.subheader("用戶管理"); st.dataframe(load_df("Users"), use_container_width=True)
