import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime
from datetime import timezone, timedelta
import uuid, io, base64, zipfile

HKT = timezone(timedelta(hours=8))
def now_hk(): return datetime.datetime.now(HKT)
def today_hk(): return now_hk().date()
st.set_page_config(page_title="TMS V25 Multi-POD ZIP", layout="wide", page_icon="🚚")

@st.cache_resource
def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    return gspread.authorize(creds)

def get_sh():
    client=get_client()
    return client.open_by_key(st.secrets["connections"]["gsheets"]["spreadsheet"])

@st.cache_data(ttl=30)
def load_df(tab):
    try:
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
    except: return pd.DataFrame()

def append_rows(tab, rows):
    get_sh().worksheet(tab).append_rows(rows, value_input_option='USER_ENTERED')
    st.cache_data.clear()
def append_row(tab,row): append_rows(tab,[row])

def ensure_sheets():
    sh=get_sh()
    existing=[ws.title for ws in sh.worksheets()]
    def add(t,h,d=None):
        if t not in existing:
            ws=sh.add_worksheet(title=t, rows=2000, cols=25)
            ws.update('A1',[h])
            if d: ws.append_rows(d, value_input_option='USER_ENTERED')
    add("Users",["username","password","name","role"],[["admin","admin123","Admin","admin"],["John","John123","John","driver"],["Mo","Mo123","Mo","driver"],["op01","op01","OP01","op"]])
    add("Jobs",["JobID","Date","Client","From","To","Driver","Status","QTY","Notes","CreatedBy","CreatedAt_HKT","POD_Link","PodTime_HKT","POD_Image"])
    add("PODs",["POD_ID","JobID","FileName","Size","Base64_Image","CreatedBy","CreatedAt_HKT","PodTime_HKT"])
    add("Expenses",["ExpenseID","Date","JobID","Driver","Type","Amount","Remark","CreatedAt_HKT"])
    add("Config",["Key","Value"],[["POD_Folder_ID",""]])

def safe_b64_decode(s):
    try:
        if not s or len(str(s))<100: return None
        s=str(s).strip()
        if "," in s: s=s.split(",",1)[1]
        s=s.replace("\n","").replace("\r","").replace(" ","")
        s=s[:len(s)//4*4]
        if len(s)==0: return None
        return base64.b64decode(s)
    except: return None

def find_pod_col(df):
    for col in df.columns:
        if df[col].astype(str).str.contains("data:image", na=False).any(): return col
    for c in df.columns:
        if "IMAGE" in c.upper(): return c
    return None

def show_table(df):
    if df.empty: st.info("無資料"); return
    pod=find_pod_col(df)
    df2=df.drop(columns=[pod], errors="ignore") if pod else df
    if "JobID" in df2.columns:
        try: df2=df2.sort_values(by="JobID", ascending=False)
        except: pass
    st.markdown(f"<div style='overflow-x:auto'>{df2.head(200).to_html(index=False, border=0)}</div>", unsafe_allow_html=True)

# ---------- Auth ----------
if "user" not in st.session_state: st.session_state.user=None
if not st.session_state.user:
    st.title("TMS V25 登入")
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

if role=="admin": tabs=st.tabs(["📊 Dashboard","📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費","👥 用戶","🧹 維護"])
elif role=="op": tabs=st.tabs(["📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費","🧹 維護"])
else: tabs=st.tabs(["🟢 可接任務池","🚚 我的任務","📸 上傳POD"])

df_jobs=load_df("Jobs")
df_pods=load_df("PODs")

# ---------- 任務管理 + POD下載 + ZIP ----------
if role in ["admin","op"]:
    idx=1 if role=="admin" else 0
    with tabs[idx]:
        st.subheader("📋 所有任務 - 已按Job#排序")
        st.caption(f"Jobs: {len(df_jobs)} | PODs表(多張): {len(df_pods)}")
        show_table(df_jobs)
        st.divider()
        st.subheader("📸 POD下載區 - 支持一個Job多張相 + 一鍵ZIP")

        # 合併舊Jobs表的單張 + 新PODs表的多張
        all_pods=[]
        # 舊表
        pod_col=find_pod_col(df_jobs)
        if pod_col:
            for _, r in df_jobs.iterrows():
                jid=str(r.get("JobID","")).strip()
                b64=str(r.get(pod_col,""))
                if len(b64)>500:
                    fb=safe_b64_decode(b64)
                    if fb:
                        pod_time=str(r.get("PodTime_HKT","") or r.get("COL_12","") or "")
                        all_pods.append({"JobID":jid, "FileName":f"{jid}_POD_{pod_time.replace(':','-').replace(' ','_')}.jpg" if pod_time else f"{jid}_POD_OLD.jpg", "bytes":fb, "time":pod_time, "source":"Jobs"})
        # 新表多張
        if not df_pods.empty:
            for _, r in df_pods.iterrows():
                jid=str(r.get("JobID","")).strip()
                b64=str(r.get("Base64_Image","") or r.get("COL_4",""))
                fb=safe_b64_decode(b64)
                if fb and jid:
                    all_pods.append({"JobID":jid, "FileName":str(r.get("FileName","") or f"{jid}_POD_{r.get('POD_ID','')}.jpg"), "bytes":fb, "time":str(r.get("PodTime_HKT","")), "source":"PODs"})

        if not all_pods:
            st.info("暫無有效POD，請去上傳POD")
        else:
            df_all=pd.DataFrame(all_pods).sort_values(by="JobID", ascending=False)
            # 全局ZIP
            st.markdown(f"**共 {len(df_all)} 張POD，{df_all['JobID'].nunique()} 個Job**")
            zip_all_buf=io.BytesIO()
            with zipfile.ZipFile(zip_all_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                for p in all_pods:
                    zf.writestr(p["FileName"], p["bytes"])
            zip_all_buf.seek(0)
            st.download_button(f"📦 打包下載全部POD ZIP ({len(all_pods)}張)", data=zip_all_buf, file_name=f"ALL_PODs_{now_hk().strftime('%Y%m%d_%H%M')}.zip", mime="application/zip", type="primary", key="zip_all")

            st.divider()
            # 按Job分組顯示
            for jid, group in df_all.groupby("JobID", sort=False):
                with st.container(border=True):
                    c_head,c_zip=st.columns([3,1])
                    c_head.markdown(f"### {jid} - {len(group)}張相")
                    # 此Job的ZIP
                    zip_buf=io.BytesIO()
                    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                        for _, p in group.iterrows():
                            zf.writestr(p["FileName"], p["bytes"])
                    zip_buf.seek(0)
                    c_zip.download_button(f"📦 打包此Job ({len(group)}張)", data=zip_buf, file_name=f"{jid}_PODs_{now_hk().strftime('%Y%m%d')}.zip", mime="application/zip", key=f"zip_{jid}_{uuid.uuid4().hex[:4]}", use_container_width=True)

                    # 顯示每張
                    cols=st.columns(3)
                    for i, (_, p) in enumerate(group.iterrows()):
                        with cols[i%3]:
                            st.image(p["bytes"], caption=p["FileName"], use_container_width=True)
                            st.download_button(f"⬇️ 下載 {p['FileName'][:25]}", data=p["bytes"], file_name=p["FileName"], mime="image/jpeg", key=f"dl_{jid}_{i}_{uuid.uuid4().hex[:4]}", use_container_width=True)

# ---------- 上傳POD 多張 ----------
with tabs[3 if role=="admin" else 2 if role=="op" else 2]:
    st.subheader("📸 上傳POD - V25 支持一個Job上傳多張 + 自動壓縮")
    my_pool=df_jobs if role in ["admin","op"] else df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else df_jobs
    my_pool=my_pool.sort_values(by="JobID", ascending=False) if not my_pool.empty else my_pool
    if my_pool.empty: st.info("無任務")
    else:
        jid=st.selectbox("選擇JobID (已按Job#排序)", my_pool["JobID"].astype(str).tolist(), key="sel_jid_v25")
        # 多張上傳
        pod_files=st.file_uploader("拍照 / 選圖 (可一次選多張)", type=["jpg","png","jpeg"], accept_multiple_files=True)
        if pod_files:
            for f in pod_files: st.image(f, width=200, caption=f.name)
            st.write(f"已選 {len(pod_files)} 張")
        if pod_files and st.button(f"上傳 {len(pod_files)} 張到 {jid}", type="primary"):
            from PIL import Image
            progress=st.progress(0)
            success=0
            for idx_file, pod_file in enumerate(pod_files):
                try:
                    img=Image.open(io.BytesIO(pod_file.getvalue())).convert("RGB")
                    img.thumbnail((500,500))
                    buf=io.BytesIO(); img.save(buf, format="JPEG", quality=40, optimize=True)
                    fb=buf.getvalue()
                    if len(fb)>25000:
                        img.thumbnail((400,400))
                        buf=io.BytesIO(); img.save(buf, format="JPEG", quality=30, optimize=True)
                        fb=buf.getvalue()
                    b64=base64.b64encode(fb).decode()
                    data_url=f"data:image/jpeg;base64,{b64}"
                    if len(data_url)>=49000:
                        st.warning(f"{pod_file.name} 壓縮後仲太大，跳過")
                        continue
                    fname=f"{jid}_POD_{now_hk().strftime('%Y%m%d_%H%M%S')}_{idx_file+1}.jpg"
                    pod_id=f"POD-{now_hk().strftime('%m%d%H%M%S')}-{uuid.uuid4().hex[:4]}"
                    # 寫入新PODs表
                    append_row("PODs",[pod_id, jid, fname, str(len(fb)), data_url, operator, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"), now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")])
                    # 同步更新Jobs表狀態為Delivered + 最新圖覆蓋方便睇
                    ws=get_sh().worksheet("Jobs"); cell=ws.find(jid)
                    if cell:
                        ws.update_cell(cell.row, 7, "Delivered")
                        ws.update_cell(cell.row, 13, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"))
                        # Jobs表只保留最後一張縮略圖，詳細去PODs表
                        ws.update_cell(cell.row, 14, data_url)
                        ws.update_cell(cell.row, 12, f"MULTI_{len(pod_files)}")
                    success+=1
                    progress.progress((idx_file+1)/len(pod_files))
                except Exception as e:
                    st.error(f"{pod_file.name} 失敗 {e}")
            st.cache_data.clear()
            st.success(f"✅ 成功上傳 {success}/{len(pod_files)} 張到 {jid}")
            st.info("去『任務管理』最底就見到多張圖 + 每個Job獨立ZIP + 全部ZIP")

# ---------- 維護：清舊爛Base64 ----------
if role in ["admin","op"]:
    idx_maint=6 if role=="admin" else 4
    if len(tabs)>idx_maint:
        with tabs[idx_maint]:
            st.subheader("🧹 維護 - 清舊爛Base64 (48977問題)")
            st.write("你之前V19-V22切斷嗰啲爛圖會導致下載失敗，呢度一鍵清走")
            if st.button("🔍 掃描爛POD", type="secondary"):
                pod_col=find_pod_col(df_jobs)
                broken=[]
                if pod_col:
                    for _, r in df_jobs.iterrows():
                        b64=str(r.get(pod_col,""))
                        if len(b64)>40000: # 超大就係被切斷
                            fb=safe_b64_decode(b64)
                            if fb is None:
                                broken.append(r.get("JobID"))
                if broken:
                    st.error(f"發現 {len(broken)} 個爛POD: {', '.join(broken)}")
                    st.session_state.broken=broken
                else:
                    st.success("無發現爛POD，全部正常")
            if st.button("🗑️ 一鍵清除爛POD的Base64欄 (保留Job，清空圖片)", type="primary"):
                ws=get_sh().worksheet("Jobs")
                pod_col_idx=None
                # 搵POD_Image欄index
                headers=ws.row_values(1)
                for i,h in enumerate(headers):
                    if "IMAGE" in h.upper() or "POD" in h.upper() and "LINK" not in h.upper():
                        if i>=13: pod_col_idx=i+1
                # 簡單：清空所有>40000的
                if not pod_col_idx: pod_col_idx=14
                all_vals=ws.get_all_values()
                cleaned=0
                for row_idx in range(1, len(all_vals)):
                    b64=all_vals[row_idx][pod_col_idx-1] if len(all_vals[row_idx])>=pod_col_idx else ""
                    if len(b64)>40000:
                        if safe_b64_decode(b64) is None:
                            ws.update_cell(row_idx+1, pod_col_idx, "")
                            cleaned+=1
                st.cache_data.clear()
                st.success(f"已清理 {cleaned} 個爛POD，請重新上傳多張版")

# ---------- 其他 ----------
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
        if not df_jobs.empty:
            mask=df_jobs["Driver"].astype(str).str.strip()==""
            avail=df_jobs[mask | (df_jobs["Status"].astype(str).str.lower()=="pending")]
            show_table(avail)
            for _, row in avail.iterrows():
                jid=row["JobID"]
                if st.button(f"搶 {jid}", key=f"claim_{jid}", type="primary"):
                    ws=get_sh().worksheet("Jobs"); cell=ws.find(jid)
                    if cell: ws.update_cell(cell.row, 6, operator); ws.update_cell(cell.row, 7, "Accepted"); st.cache_data.clear(); st.rerun()
    with tabs[1]: st.subheader("🚚 我的任務"); my=df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else pd.DataFrame(); show_table(my)
