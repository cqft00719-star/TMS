import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import datetime
from datetime import timezone, timedelta
import uuid, io, base64

HKT = timezone(timedelta(hours=8))
def now_hk(): return datetime.datetime.now(HKT)
def today_hk(): return now_hk().date()
st.set_page_config(page_title="TMS V15", layout="wide", page_icon="🚚")

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
    try:
        ws=get_sh().worksheet(tab)
        vals=ws.get_all_values()
        if len(vals)<1: return pd.DataFrame()
        headers=[h.strip() for h in vals[0]]
        rows=[]
        for r in vals[1:]:
            if any(str(c).strip()!="" for c in r):
                if len(r)<len(headers): r=r+[""]*(len(headers)-len(r))
                rows.append(r[:len(headers)])
        df=pd.DataFrame(rows, columns=headers)
        for col in ["JobID","Driver","Status","Client"]:
            if col in df.columns: df[col]=df[col].astype(str).str.strip()
        df=df.replace({"nan":"","None":""})
        return df
    except Exception as e:
        st.error(f"讀 {tab} 失敗: {e}"); return pd.DataFrame()

def append_rows(tab, rows):
    get_sh().worksheet(tab).append_rows(rows, value_input_option='USER_ENTERED')
    st.cache_data.clear()
def append_row(tab,row): append_rows(tab,[row])

def ensure_sheets():
    sh=get_sh()
    existing=[ws.title for ws in sh.worksheets()]
    def add(title,headers,defaults=None):
        if title not in existing:
            ws=sh.add_worksheet(title=title, rows=1000, cols=20)
            ws.update('A1',[headers])
            if defaults: ws.append_rows(defaults, value_input_option='USER_ENTERED')
    add("Users",["username","password","name","role"],[["admin","admin123","Admin","admin"],["John","John123","John","driver"],["Mo","Mo123","Mo","driver"],["op01","op01","OP01","op"]])
    add("Jobs",["JobID","Date","Client","From","To","Driver","Status","QTY","Notes","CreatedBy","CreatedAt_HKT","POD_Link","PodTime_HKT","POD_Image"])
    add("Expenses",["ExpenseID","Date","JobID","Driver","Type","Amount","Remark","CreatedAt_HKT"])
    add("Config",["Key","Value"],[["POD_Folder_ID",""]])

def safe_display_df(df):
    # 唔顯示POD_Image，避開PyArrow爆
    if df.empty: return df
    return df.drop(columns=["POD_Image"], errors="ignore")

if "user" not in st.session_state: st.session_state.user=None
if not st.session_state.user:
    st.title("TMS V15 登入")
    df_u=load_df("Users")
    if df_u.empty: df_u=pd.DataFrame([{"username":"admin","password":"admin123","name":"Admin","role":"admin"}])
    u=st.text_input("用戶名"); p=st.text_input("密碼",type="password")
    if st.button("登入",type="primary", use_container_width=True):
        m=df_u[(df_u["username"]==str(u).strip()) & (df_u["password"]==str(p).strip())]
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
df_exp=load_df("Expenses")

if role=="admin":
    with tabs[0]:
        st.subheader("Dashboard")
        st.write(f"Jobs表讀到 {len(df_jobs)} 行")
        if not df_jobs.empty:
            try: st.dataframe(safe_display_df(df_jobs).sort_values(by="JobID", ascending=False), use_container_width=True)
            except Exception as e: st.dataframe(safe_display_df(df_jobs), use_container_width=True)

if role in ["admin","op"]:
    idx=1 if role=="admin" else 0
    with tabs[idx]:
        st.subheader("📋 所有任務 - Job#排序 + 下載POD")
        st.write(f"共 {len(df_jobs)} 張")
        if not df_jobs.empty:
            df_show=df_jobs.sort_values(by="JobID", ascending=False) if "JobID" in df_jobs.columns else df_jobs
            st.dataframe(safe_display_df(df_show), use_container_width=True, height=350)
            st.divider(); st.write("**📸 POD下載區 - Job# Prefix**")
            for _, r in df_show.iterrows():
                jid=str(r.get("JobID","")).strip()
                if not jid: continue
                b64img=str(r.get("POD_Image","")) if "POD_Image" in df_show.columns else ""
                link=str(r.get("POD_Link",""))
                if b64img.startswith("data:") or link.startswith("http"):
                    with st.container(border=True):
                        c1,c2=st.columns([3,1])
                        c1.markdown(f"**{jid}** | {r.get('Client','')} | {r.get('Status','')}")
                        if b64img.startswith("data:"):
                            try:
                                _, bdata=b64img.split(",",1); fb=base64.b64decode(bdata)
                                c1.image(fb, width=250)
                                c2.download_button(f"⬇️ 下載 {jid}_POD.jpg", fb, file_name=f"{jid}_POD.jpg", mime="image/jpeg", key=f"dl_{jid}", type="primary", use_container_width=True)
                            except: pass
                        elif link.startswith("http"):
                            c2.link_button("🔗 查看", link, use_container_width=True)
        else:
            st.info("暫無任務")

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
        st.subheader("🟢 可接任務池 - V15修復")
        if df_jobs.empty: st.info("暫無")
        else:
            def is_pending(s): return str(s).strip().lower()=="pending"
            def is_empty(d): return str(d).strip()=="" or str(d).strip().lower() in ["nan","none"]
            mask_p=df_jobs["Status"].apply(is_pending) if "Status" in df_jobs.columns else pd.Series([False]*len(df_jobs))
            mask_e=df_jobs["Driver"].apply(is_empty) if "Driver" in df_jobs.columns else pd.Series([False]*len(df_jobs))
            avail=df_jobs[mask_p | mask_e].copy()
            avail=avail[~avail["Status"].astype(str).str.lower().isin(["delivered","completed"])] if not avail.empty else avail
            avail=avail.sort_values(by="JobID", ascending=False) if not avail.empty else avail
            st.success(f"可接：{len(avail)} 張 - 已按Job#排序")
            st.dataframe(safe_display_df(avail), use_container_width=True)
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
                            else:
                                ws.update_cell(cell.row, 6, operator); ws.update_cell(cell.row, 7, "Accepted")
                                st.cache_data.clear(); st.rerun()
    with tabs[1]:
        st.subheader("🚚 我的任務")
        my_jobs=df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else pd.DataFrame()
        my_jobs=my_jobs.sort_values(by="JobID", ascending=False) if not my_jobs.empty else my_jobs
        st.dataframe(safe_display_df(my_jobs), use_container_width=True)

with tabs[3 if role=="admin" else 2 if role=="op" else 2]:
    st.subheader("📸 上傳POD - Job# Prefix")
    my_pool=df_jobs if role in ["admin","op"] else df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else pd.DataFrame()
    my_pool=my_pool.sort_values(by="JobID", ascending=False) if not my_pool.empty else my_pool
    if my_pool.empty: st.info("無任務")
    else:
        jid=st.selectbox("選擇JobID (已按Job#排序)", my_pool["JobID"].astype(str).tolist())
        pod_file=st.file_uploader("拍照", type=["jpg","png","jpeg"])
        if pod_file: st.image(pod_file, width=300, caption=f"{jid}_POD_{pod_file.name}")
        if pod_file and st.button("上傳POD", type="primary"):
            try:
                from PIL import Image
                img=Image.open(io.BytesIO(pod_file.getvalue())); img.thumbnail((1200,1200))
                buf=io.BytesIO(); img.save(buf, format="JPEG", quality=80)
                fb=buf.getvalue(); b64=base64.b64encode(fb).decode(); data_url=f"data:image/jpeg;base64,{b64}"
                save_name=f"{jid}_POD_{now_hk().strftime('%Y%m%d_%H%M')}.jpg"
                ws=get_sh().worksheet("Jobs"); cell=ws.find(jid)
                if cell:
                    ws.update_cell(cell.row, 12, "BASE64_SAVED"); ws.update_cell(cell.row, 13, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")); ws.update_cell(cell.row, 7, "Delivered"); ws.update_cell(cell.row, 14, data_url[:49000])
                st.cache_data.clear()
                st.success(f"✅ {save_name} 已保存")
                st.download_button(f"⬇️ 下載 {save_name}", fb, file_name=save_name, mime="image/jpeg", type="primary")
            except Exception as e: st.error(str(e))

with tabs[4 if role=="admin" else 3 if role=="op" else 3]:
    st.subheader("💰 雜費"); st.dataframe(safe_display_df(df_exp), use_container_width=True)

if role=="admin":
    with tabs[5]:
        st.subheader("用戶"); st.dataframe(load_df("Users"), use_container_width=True)
