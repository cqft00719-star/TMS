import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import datetime
from datetime import timezone, timedelta
import uuid, io, base64

HKT = timezone(timedelta(hours=8))
def now_hk(): return datetime.datetime.now(HKT)
def today_hk(): return now_hk().date()
st.set_page_config(page_title="TMS V23 Drive", layout="wide", page_icon="🚚")

@st.cache_resource
def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    return gspread.authorize(creds), creds, Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)

def get_sh():
    client,_,_=get_client()
    return client.open_by_key(st.secrets["connections"]["gsheets"]["spreadsheet"])

def get_drive():
    _,_,creds=get_client()
    return build('drive','v3',credentials=creds)

@st.cache_data(ttl=30)
def load_df(tab):
    ws=get_sh().worksheet(tab)
    vals=ws.get_all_values()
    if len(vals)<1: return pd.DataFrame()
    headers=[h.strip() for h in vals[0]]
    clean=[]
    for i,h in enumerate(headers):
        if h=="": h=f"COL_{i}"
        if clean.count(h)>0: h=f"{h}_{i}"
        clean.append(h)
    rows=[]
    for r in vals[1:]:
        if any(str(c).strip()!="" for c in r):
            if len(r)<len(clean): r=r+[""]*(len(clean)-len(r))
            rows.append(r[:len(clean)])
    df=pd.DataFrame(rows, columns=clean)
    return df

def append_row(tab,row):
    get_sh().worksheet(tab).append_rows([row], value_input_option='USER_ENTERED')
    st.cache_data.clear()

def ensure_sheets():
    sh=get_sh()
    existing=[ws.title for ws in sh.worksheets()]
    def add(t,h,d=None):
        if t not in existing:
            ws=sh.add_worksheet(title=t, rows=1000, cols=20)
            ws.update('A1',[h])
            if d: ws.append_rows(d, value_input_option='USER_ENTERED')
    add("Users",["username","password","name","role"],[["admin","admin123","Admin","admin"],["John","John123","John","driver"],["Mo","Mo123","Mo","driver"],["op01","op01","OP01","op"]])
    add("Jobs",["JobID","Date","Client","From","To","Driver","Status","QTY","Notes","CreatedBy","CreatedAt_HKT","POD_Link","PodTime_HKT","POD_Image"])
    add("Expenses",["ExpenseID","Date","JobID","Driver","Type","Amount","Remark","CreatedAt_HKT"])
    add("Config",["Key","Value"],[["POD_Folder_ID",""]])

def get_pod_folder_id():
    try:
        df=load_df("Config")
        if not df.empty and "Key" in df.columns:
            row=df[df["Key"]=="POD_Folder_ID"]
            if not row.empty:
                v=str(row.iloc[0]["Value"]).strip()
                if "folders/" in v: v=v.split("folders/")[1].split("/")[0]
                if len(v)>10: return v
    except: pass
    return None

def upload_to_drive(file_bytes, filename, folder_id=None):
    drive=get_drive()
    file_metadata={'name':filename}
    if folder_id: file_metadata['parents']=[folder_id]
    media=MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype='image/jpeg', resumable=True)
    file=drive.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
    # 公開可讀
    try:
        drive.permissions().create(fileId=file['id'], body={'type':'anyone','role':'reader'}).execute()
    except: pass
    return file.get('webViewLink'), file.get('id')

def find_pod_col(df):
    for col in df.columns:
        if df[col].astype(str).str.contains("data:image", na=False).any(): return col
        if "IMAGE" in col.upper(): return col
    return None

def show_table(df):
    if df.empty:
        st.info("無資料"); return
    pod_col=find_pod_col(df)
    df_show=df.drop(columns=[pod_col], errors="ignore") if pod_col else df
    if "JobID" in df_show.columns:
        try: df_show=df_show.sort_values(by="JobID", ascending=False)
        except: pass
    html=df_show.head(200).to_html(index=False, border=0)
    st.markdown(f"<div style='overflow-x:auto'>{html}</div>", unsafe_allow_html=True)

if "user" not in st.session_state: st.session_state.user=None
if not st.session_state.user:
    st.title("TMS V23 登入")
    df_u=load_df("Users")
    u=st.text_input("用戶名"); p=st.text_input("密碼",type="password")
    if st.button("登入",type="primary", use_container_width=True):
        m=df_u[(df_u["username"].astype(str).str.strip()==str(u).strip()) & (df_u["password"].astype(str).str.strip()==str(p).strip())]
        if not m.empty: st.session_state.user=m.iloc[0].to_dict(); st.rerun()
        else: st.error("帳號錯")
    st.stop()

ensure_sheets()
user=st.session_state.user
role=str(user.get("role","driver")).lower()
operator=user.get("username","")
st.sidebar.title(f"👤 {user.get('name')} ({role})")
if st.sidebar.button("登出"): st.session_state.user=None; st.cache_data.clear(); st.rerun()
if st.sidebar.button("🔄 清除Cache刷新"): st.cache_data.clear(); st.rerun()
folder_id=get_pod_folder_id()
if folder_id: st.sidebar.caption(f"Drive Folder: {folder_id[:15]}...")
else: st.sidebar.warning("未設定POD_Folder_ID，去Config填Drive Folder ID")

if role=="admin": tabs=st.tabs(["📊 Dashboard","📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費","👥 用戶"])
elif role=="op": tabs=st.tabs(["📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費"])
else: tabs=st.tabs(["🟢 可接任務池","🚚 我的任務","📸 上傳POD","💰 我的雜費"])

df_jobs=load_df("Jobs")

if role in ["admin","op"]:
    idx=1 if role=="admin" else 0
    with tabs[idx]:
        st.subheader("📋 所有任務")
        st.write(f"共 {len(df_jobs)} 張 - 已按Job#排序")
        show_table(df_jobs)
        st.divider()
        st.subheader("📸 POD下載區 - Job#_POD_時間.jpg (Drive版，永久有效)")
        df_show=df_jobs.sort_values(by="JobID", ascending=False) if not df_jobs.empty and "JobID" in df_jobs.columns else df_jobs
        found=False
        for _, r in df_show.iterrows():
            jid=str(r.get("JobID","")).strip()
            if not jid: continue
            link=str(r.get("POD_Link","") or r.get("COL_11","") or "").strip()
            pod_time=str(r.get("PodTime_HKT","") or r.get("COL_12","") or "").strip()
            # 判斷是否Drive Link
            if link.startswith("http"):
                found=True
                with st.container(border=True):
                    c1,c2=st.columns([3,1])
                    c1.markdown(f"**{jid}** | {r.get('Client','')} | {r.get('Status','')} | {pod_time}")
                    safe_time=pod_time.replace(":","-").replace(" ","_").replace("/","-") if pod_time else now_hk().strftime("%m%d_%H%M")
                    fname=f"{jid}_POD_{safe_time}.jpg"
                    c1.write(f"Link: {link[:80]}...")
                    c2.link_button(f"⬇️ 下載 {fname}", link, type="primary", use_container_width=True)
            else:
                # 舊Base64爛圖提示重新上傳
                pod_col=find_pod_col(df_jobs)
                b64=str(r.get(pod_col,"")) if pod_col else ""
                if len(b64)>1000:
                    with st.container(border=True):
                        c1,c2=st.columns([3,1])
                        c1.markdown(f"**{jid}** | 舊圖已損壞(被切斷) | {pod_time}")
                        c1.error("舊Base64已損壞，請去『上傳POD』重新上傳一次，之後會自動存Drive")
                        c2.write("⬇️ 需要重上傳")
        if not found:
            st.info("暫無Drive POD，請去『上傳POD』上傳，新圖會自動出下載掣")

with tabs[3 if role=="admin" else 2 if role=="op" else 2]:
    st.subheader("📸 上傳POD - V23 Drive版")
    if not folder_id:
        st.error("請先去Google Drive開一個Folder叫 TMS_POD，右鍵 -> 分享 -> 複製連結，Link入面folders/後面嗰段ID，貼去Sheets的Config表 POD_Folder_ID Value格，再按清除Cache刷新")
        st.write("例如 https://drive.google.com/drive/folders/1AbCdeFgHiJkLmNoPqRsTuVwXyZ -> 1AbCdeFgHiJkLmNoPqRsTuVwXyZ")
    my_pool=df_jobs if role in ["admin","op"] else df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else df_jobs
    my_pool=my_pool.sort_values(by="JobID", ascending=False) if not my_pool.empty else my_pool
    if my_pool.empty: st.info("無任務")
    else:
        jid=st.selectbox("選擇JobID", my_pool["JobID"].astype(str).tolist())
        pod_file=st.file_uploader("拍照 / 選圖", type=["jpg","png","jpeg"])
        if pod_file: st.image(pod_file, width=300, caption=f"{jid}_POD")
        if pod_file and st.button("上傳到Drive", type="primary", disabled=not folder_id):
            try:
                from PIL import Image
                img=Image.open(io.BytesIO(pod_file.getvalue()))
                img.thumbnail((1600,1600))
                buf=io.BytesIO(); img.save(buf, format="JPEG", quality=85, optimize=True)
                fb=buf.getvalue()
                safe_time=now_hk().strftime("%Y-%m-%d_%H%M%S")
                fname=f"{jid}_POD_{safe_time}.jpg"
                with st.spinner(f"上傳 {fname} 到Drive..."):
                    link,fid=upload_to_drive(fb, fname, folder_id)
                ws=get_sh().worksheet("Jobs"); cell=ws.find(jid)
                if cell:
                    ws.update_cell(cell.row, 12, link)
                    ws.update_cell(cell.row, 13, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"))
                    ws.update_cell(cell.row, 7, "Delivered")
                    ws.update_cell(cell.row, 14, "DRIVE_SAVED")
                st.cache_data.clear()
                st.success(f"✅ 已上傳到Drive: {fname}")
                st.link_button(f"⬇️ 立即下載 {fname}", link, type="primary")
                st.write(link)
            except Exception as e:
                st.error(f"上傳失敗 {e}")
                import traceback; st.code(traceback.format_exc())

# 其他Tab簡化
if role=="admin":
    with tabs[0]: st.subheader("Dashboard"); show_table(df_jobs)
if role=="driver":
    with tabs[0]:
        st.subheader("🟢 可接任務池")
        if not df_jobs.empty:
            def is_pending(s): return str(s).strip().lower()=="pending"
            def is_empty(d): return str(d).strip()==""
            mask_p=df_jobs["Status"].apply(is_pending) if "Status" in df_jobs.columns else pd.Series([False]*len(df_jobs))
            mask_e=df_jobs["Driver"].apply(is_empty) if "Driver" in df_jobs.columns else pd.Series([False]*len(df_jobs))
            avail=df_jobs[mask_p | mask_e].copy()
            show_table(avail)
            for _, row in avail.iterrows():
                jid=row["JobID"]
                if st.button(f"搶 {jid}", key=f"claim_{jid}"):
                    ws=get_sh().worksheet("Jobs"); cell=ws.find(jid)
                    if cell: ws.update_cell(cell.row, 6, operator); ws.update_cell(cell.row, 7, "Accepted"); st.cache_data.clear(); st.rerun()
    with tabs[1]: st.subheader("🚚 我的任務"); my=df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else pd.DataFrame(); show_table(my)
