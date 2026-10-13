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
st.set_page_config(page_title="TMS V24 Final", layout="wide", page_icon="🚚")

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
            ws=sh.add_worksheet(title=t, rows=1000, cols=25)
            ws.update('A1',[h])
            if d: ws.append_rows(d, value_input_option='USER_ENTERED')
    add("Users",["username","password","name","role"],[["admin","admin123","Admin","admin"],["John","John123","John","driver"],["Mo","Mo123","Mo","driver"],["op01","op01","OP01","op"]])
    add("Jobs",["JobID","Date","Client","From","To","Driver","Status","QTY","Notes","CreatedBy","CreatedAt_HKT","POD_Link","PodTime_HKT","POD_Image"])
    add("Expenses",["ExpenseID","Date","JobID","Driver","Type","Amount","Remark","CreatedAt_HKT"])
    add("Config",["Key","Value"],[["POD_Folder_ID",""]])

def find_pod_col(df):
    for col in df.columns:
        if df[col].astype(str).str.contains("data:image", na=False).any(): return col
    for c in df.columns:
        if "IMAGE" in c.upper(): return c
    return None

def safe_b64_decode(s):
    try:
        if "," in s: s=s.split(",",1)[1]
        s=s.strip().replace("\n","").replace("\r","")
        s=s[:len(s)//4*4]
        return base64.b64decode(s)
    except: return None

def show_table(df):
    if df.empty:
        st.info("無資料"); return
    pod=find_pod_col(df)
    df2=df.drop(columns=[pod], errors="ignore") if pod else df
    if "JobID" in df2.columns:
        try: df2=df2.sort_values(by="JobID", ascending=False)
        except: pass
    st.markdown(f"<div style='overflow-x:auto'>{df2.head(200).to_html(index=False, border=0)}</div>", unsafe_allow_html=True)

if "user" not in st.session_state: st.session_state.user=None
if not st.session_state.user:
    st.title("TMS V24 登入")
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

if role in ["admin","op"]:
    idx=1 if role=="admin" else 0
    with tabs[idx]:
        st.subheader("📋 所有任務 - 已按Job#排序")
        show_table(df_jobs)
        st.divider()
        st.subheader("📸 POD下載區 - 文件名 Job#_POD_時間.jpg (已按Job#排序)")
        pod_col=find_pod_col(df_jobs)
        df_show=df_jobs.sort_values(by="JobID", ascending=False) if "JobID" in df_jobs.columns else df_jobs
        found=False
        for _, r in df_show.iterrows():
            jid=str(r.get("JobID","")).strip()
            if not jid: continue
            b64=str(r.get(pod_col,"")) if pod_col else ""
            if len(b64)<100:
                for v in r.values:
                    if str(v).startswith("data:image") and len(str(v))>100: b64=str(v); break
            link=str(r.get("POD_Link","") or r.get("COL_11","") or "").strip()
            pod_time=str(r.get("PodTime_HKT","") or r.get("COL_12","") or "").strip()
            fb=None
            if len(b64)>500:
                fb=safe_b64_decode(b64)
            # Drive link都支援
            if fb is not None:
                found=True
                with st.container(border=True):
                    c1,c2=st.columns([3,1])
                    c1.markdown(f"**{jid}** | {r.get('Client','')} | {r.get('Status','')} | {pod_time} | 大小:{len(b64)}")
                    safe_time=pod_time.replace(":","-").replace(" ","_").replace("/","-") if pod_time else now_hk().strftime("%m%d_%H%M")
                    fname=f"{jid}_POD_{safe_time}.jpg"
                    c1.image(fb, width=350, caption=fname)
                    c2.download_button(f"⬇️ 下載 {fname}", data=fb, file_name=fname, mime="image/jpeg", key=f"dl_{jid}_{safe_time}_{uuid.uuid4().hex[:4]}", type="primary", use_container_width=True)
            elif link.startswith("http"):
                found=True
                with st.container(border=True):
                    c1,c2=st.columns([3,1])
                    c1.markdown(f"**{jid}** | {pod_time}")
                    c2.link_button(f"⬇️ 下載 {jid}_POD.jpg", link, type="primary", use_container_width=True)
        if not found:
            st.info("暫無有效POD，請去上傳POD頁重新上傳")

with tabs[3 if role=="admin" else 2 if role=="op" else 2]:
    st.subheader("📸 上傳POD - V24 無需Drive，自動壓縮 <35k")
    st.caption("V23的403係因為Service Account無quota，V24改為超級壓縮，直接存Sheet，唔經Drive，一定得")
    my_pool=df_jobs if role in ["admin","op"] else df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else df_jobs
    my_pool=my_pool.sort_values(by="JobID", ascending=False) if not my_pool.empty else my_pool
    if my_pool.empty: st.info("無任務")
    else:
        jid=st.selectbox("選擇JobID (已按Job#排序)", my_pool["JobID"].astype(str).tolist())
        pod_file=st.file_uploader("拍照 / 選圖", type=["jpg","png","jpeg"])
        if pod_file: st.image(pod_file, width=300, caption=f"{jid}_POD")
        if pod_file and st.button("上傳POD", type="primary"):
            from PIL import Image
            try:
                img=Image.open(io.BytesIO(pod_file.getvalue())).convert("RGB")
                # V24 核心：壓到500px，質量40，保證<35k
                img.thumbnail((500,500))
                buf=io.BytesIO()
                img.save(buf, format="JPEG", quality=40, optimize=True)
                fb=buf.getvalue()
                # 如果仲大過35k，再壓到400px 質量30
                if len(fb)>25000:
                    img.thumbnail((400,400))
                    buf=io.BytesIO()
                    img.save(buf, format="JPEG", quality=30, optimize=True)
                    fb=buf.getvalue()
                b64=base64.b64encode(fb).decode()
                data_url=f"data:image/jpeg;base64,{b64}"
                st.write(f"壓縮後大小: {len(fb)} bytes, Base64長度: {len(data_url)} (必須<49000)")
                if len(data_url)>=49000:
                    st.error("仲係太大，請影近啲或揀張細啲嘅圖")
                else:
                    ws=get_sh().worksheet("Jobs")
                    cell=ws.find(jid)
                    if cell:
                        ws.update_cell(cell.row, 12, "COMPRESSED_V24")
                        ws.update_cell(cell.row, 13, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"))
                        ws.update_cell(cell.row, 7, "Delivered")
                        ws.update_cell(cell.row, 14, data_url)
                    st.cache_data.clear()
                    st.success(f"✅ {jid}_POD_{now_hk().strftime('%Y%m%d_%H%M')}.jpg 已保存")
                    st.download_button(f"⬇️ 立即下載 {jid}_POD.jpg", data=fb, file_name=f"{jid}_POD_{now_hk().strftime('%Y%m%d_%H%M')}.jpg", mime="image/jpeg", type="primary")
                    st.info("返去『任務管理』最底就見到下載區有圖同下載掣")
            except Exception as e:
                st.error(f"失敗 {e}")
                import traceback; st.code(traceback.format_exc())

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
        st.subheader("🟢 可接任務池 V24")
        if not df_jobs.empty:
            mask=df_jobs["Driver"].astype(str).str.strip()==""
            avail=df_jobs[mask | (df_jobs["Status"].astype(str).str.lower()=="pending")]
            show_table(avail)
            for _, row in avail.iterrows():
                jid=row["JobID"]
                if st.button(f"搶 {jid}", key=f"claim_{jid}", type="primary"):
                    ws=get_sh().worksheet("Jobs"); cell=ws.find(jid)
                    if cell: ws.update_cell(cell.row, 6, operator); ws.update_cell(cell.row, 7, "Accepted"); st.cache_data.clear(); st.rerun()
    with tabs[1]:
        st.subheader("🚚 我的任務"); my=df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else pd.DataFrame(); show_table(my)
