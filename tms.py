import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime
from datetime import timezone, timedelta
import uuid, io, base64

HKT = timezone(timedelta(hours=8))
def now_hk(): return datetime.datetime.now(HKT)
def today_hk(): return now_hk().date()
st.set_page_config(page_title="TMS V18", layout="wide", page_icon="🚚")

@st.cache_resource
def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    return gspread.authorize(creds), creds

def get_sh():
    client,_=get_client()
    return client.open_by_key(st.secrets["connections"]["gsheets"]["spreadsheet"])

@st.cache_data(ttl=30)
def load_df(tab):
    ws=get_sh().worksheet(tab)
    vals=ws.get_all_values()
    if len(vals)<1: return pd.DataFrame()
    headers=[h.strip() for h in vals[0]]
    clean=[]
    for i,h in enumerate(headers):
        if h=="": h=f"COL_{i}"
        clean.append(h)
    rows=[]
    for r in vals[1:]:
        if any(str(c).strip()!="" for c in r):
            if len(r)<len(clean): r=r+[""]*(len(clean)-len(r))
            rows.append(r[:len(clean)])
    df=pd.DataFrame(rows, columns=clean)
    df=df.replace({"nan":"","None":""})
    return df

def append_rows(tab, rows):
    get_sh().worksheet(tab).append_rows(rows, value_input_option='USER_ENTERED')
    st.cache_data.clear()
def append_row(tab,row): append_rows(tab,[row])

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

def show_table(df, max_rows=100):
    if df.empty:
        st.info("無資料"); return
    df_show=df.drop(columns=["POD_Image"], errors="ignore")
    if "JobID" in df_show.columns:
        try: df_show=df_show.sort_values(by="JobID", ascending=False)
        except: pass
    df_show=df_show.head(max_rows)
    html=df_show.to_html(index=False, border=0)
    st.markdown(f"<div style='overflow-x:auto'>{html}</div>", unsafe_allow_html=True)

if "user" not in st.session_state: st.session_state.user=None
if not st.session_state.user:
    st.title("TMS V18 登入")
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

if role=="admin": tabs=st.tabs(["📊 Dashboard","📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費","👥 用戶"])
elif role=="op": tabs=st.tabs(["📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費"])
else: tabs=st.tabs(["🟢 可接任務池","🚚 我的任務","📸 上傳POD","💰 我的雜費"])

df_jobs=load_df("Jobs")

if role=="admin":
    with tabs[0]:
        st.subheader("Dashboard")
        st.write(f"Jobs表讀到 {len(df_jobs)} 行")
        show_table(df_jobs)

if role in ["admin","op"]:
    idx=1 if role=="admin" else 0
    with tabs[idx]:
        st.subheader("📋 所有任務 - Job#排序 + 可下載")
        st.write(f"共 {len(df_jobs)} 張")
        show_table(df_jobs, 200)
        st.divider()
        st.subheader("📸 POD下載區 - 文件名 Job#_POD_時間.jpg")
        if df_jobs.empty:
            st.info("暫無任務")
        else:
            df_show=df_jobs.sort_values(by="JobID", ascending=False) if "JobID" in df_jobs.columns else df_jobs
            found=False
            for _, r in df_show.iterrows():
                jid=str(r.get("JobID","")).strip()
                if not jid: continue
                b64img=str(r.get("POD_Image","")).strip()
                link=str(r.get("POD_Link","")).strip()
                pod_time=str(r.get("PodTime_HKT","")).strip()
                # V18 強制：只要POD_Image有野或POD_Link有野就顯示
                has_image = len(b64img) > 100 # 有Base64
                has_link = link.startswith("http")
                if has_image or has_link:
                    found=True
                    with st.container(border=True):
                        c1,c2=st.columns([3,1])
                        c1.markdown(f"**{jid}** | Client:{r.get('Client','')} | Status:{r.get('Status','')} | {pod_time}")
                        safe_time=pod_time.replace(':','-').replace(' ','_').replace('/','-') if pod_time else now_hk().strftime('%m%d_%H%M')
                        fname=f"{jid}_POD_{safe_time}.jpg"
                        if has_image:
                            try:
                                # 兼容有無data:頭
                                if "," in b64img:
                                    bdata=b64img.split(",",1)[1]
                                else:
                                    bdata=b64img
                                fb=base64.b64decode(bdata)
                                c1.image(fb, width=300, caption=fname)
                                c2.download_button(f"⬇️ 下載 {fname}", data=fb, file_name=fname, mime="image/jpeg", key=f"dl_{jid}_{safe_time}_v18", type="primary", use_container_width=True)
                            except Exception as e:
                                c1.error(f"圖片解碼失敗 {e} 長度:{len(b64img)}")
                                # 顯示頭100字debug
                                c1.code(b64img[:200])
                        elif has_link:
                            c2.link_button("🔗 查看Drive", link, use_container_width=True)
            if not found:
                st.warning("暫時掃唔到POD，原因：Sheet入面 POD_Image 欄係空。請去『上傳POD』再上傳一次，我會自動按Job#命名")
                # Debug 顯示每行POD欄長度
                for _, r in df_show.iterrows():
                    st.write(f"{r.get('JobID')}: POD_Image長度={len(str(r.get('POD_Image','')))} POD_Link={str(r.get('POD_Link',''))[:50]}")

if role in ["admin","op"]:
    idx=2 if role=="admin" else 1
    with tabs[idx]:
        st.subheader("📤 上傳Order")
        with st.form("manual", clear_on_submit=True):
            client=st.text_input("客戶*"); date=st.date_input("日期", value=today_hk())
            from_loc=st.text_input("起點"); to_loc=st.text_input("終點")
            qty=st.number_input("數量", 1.0); notes=st.text_input("備註")
            if st.form_submit_button("建立"):
                if client.strip():
                    jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                    append_row("Jobs",[jid,str(date),client,from_loc,to_loc,"","Pending",str(qty),notes,operator,now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"),"","",""])
                    st.success(jid); st.rerun()

if role=="driver":
    with tabs[0]:
        st.subheader("🟢 可接任務池")
        if df_jobs.empty: st.info("暫無")
        else:
            def is_pending(s): return str(s).strip().lower()=="pending"
            def is_empty(d): return str(d).strip()=="" or str(d).strip().lower() in ["nan","none"]
            mask_p=df_jobs["Status"].apply(is_pending) if "Status" in df_jobs.columns else pd.Series([False]*len(df_jobs))
            mask_e=df_jobs["Driver"].apply(is_empty) if "Driver" in df_jobs.columns else pd.Series([False]*len(df_jobs))
            avail=df_jobs[mask_p | mask_e].copy()
            avail=avail[~avail["Status"].astype(str).str.lower().isin(["delivered","completed"])] if not avail.empty else avail
            st.success(f"可接：{len(avail)} 張")
            show_table(avail, 200)
            for _, row in avail.iterrows():
                jid=row["JobID"]
                with st.container(border=True):
                    c1,c2=st.columns([4,1])
                    c1.markdown(f"**{jid}** | {row.get('Client','')} | {row.get('From','')}->{row.get('To','')}")
                    if c2.button("搶此單", key=f"claim_{jid}", type="primary", use_container_width=True):
                        ws=get_sh().worksheet("Jobs"); cell=ws.find(jid)
                        if cell:
                            curr=str(ws.cell(cell.row, 6).value or "").strip()
                            if curr!="": st.error(f"已被 {curr} 搶")
                            else: ws.update_cell(cell.row, 6, operator); ws.update_cell(cell.row, 7, "Accepted"); st.cache_data.clear(); st.rerun()
    with tabs[1]:
        st.subheader("🚚 我的任務")
        my_jobs=df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else pd.DataFrame()
        show_table(my_jobs, 200)

with tabs[3 if role=="admin" else 2 if role=="op" else 2]:
    st.subheader("📸 上傳POD - Job# Prefix 自動下載")
    my_pool=df_jobs if role in ["admin","op"] else df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else pd.DataFrame()
    my_pool=my_pool.sort_values(by="JobID", ascending=False) if not my_pool.empty else my_pool
    if my_pool.empty: st.info("無任務")
    else:
        jid=st.selectbox("選擇JobID (已按Job#排序)", my_pool["JobID"].astype(str).tolist())
        pod_file=st.file_uploader("拍照", type=["jpg","png","jpeg"])
        if pod_file: st.image(pod_file, width=300, caption=f"{jid}_POD_{pod_file.name}")
        if pod_file and st.button("上傳POD", type="primary"):
            from PIL import Image
            img=Image.open(io.BytesIO(pod_file.getvalue())); img.thumbnail((1200,1200))
            buf=io.BytesIO(); img.save(buf, format="JPEG", quality=80)
            fb=buf.getvalue(); b64=base64.b64encode(fb).decode(); data_url=f"data:image/jpeg;base64,{b64}"
            save_name=f"{jid}_POD_{now_hk().strftime('%Y%m%d_%H%M')}.jpg"
            ws=get_sh().worksheet("Jobs"); cell=ws.find(jid)
            if cell:
                ws.update_cell(cell.row, 12, "BASE64_SAVED"); ws.update_cell(cell.row, 13, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")); ws.update_cell(cell.row, 7, "Delivered"); ws.update_cell(cell.row, 14, data_url[:49000])
            st.cache_data.clear(); st.success(f"✅ {save_name} 已保存"); st.download_button(f"⬇️ 立即下載 {save_name}", fb, file_name=save_name, mime="image/jpeg", type="primary")

with tabs[4 if role=="admin" else 3 if role=="op" else 3]:
    st.subheader("💰 雜費"); show_table(load_df("Expenses"))
if role=="admin":
    with tabs[5]: st.subheader("用戶"); show_table(load_df("Users"))
