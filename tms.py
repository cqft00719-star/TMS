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
st.set_page_config(page_title="TMS V26 Full", layout="wide", page_icon="🚚")

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

if "user" not in st.session_state: st.session_state.user=None
if not st.session_state.user:
    st.title("TMS V26 登入")
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
elif role=="op": tabs=st.tabs(["📊 Dashboard","📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費","🧹 維護"])
else: tabs=st.tabs(["🟢 可接任務池","🚚 我的任務","📸 上傳POD"])

df_jobs=load_df("Jobs")
df_pods=load_df("PODs")

# ==================== 1) Dashboard 補返 ====================
with tabs[0]:
    st.subheader("📊 Dashboard - 數據 + 圖表")
    if df_jobs.empty:
        st.info("暫無Jobs資料")
    else:
        # KPI
        total=len(df_jobs)
        pending=len(df_jobs[df_jobs["Status"].astype(str).str.lower()=="pending"]) if "Status" in df_jobs.columns else 0
        delivered=len(df_jobs[df_jobs["Status"].astype(str).str.lower()=="delivered"]) if "Status" in df_jobs.columns else 0
        accepted=len(df_jobs[df_jobs["Status"].astype(str).str.lower()=="accepted"]) if "Status" in df_jobs.columns else 0
        c1,c2,c3,c4=st.columns(4)
        c1.metric("總任務", total)
        c2.metric("Pending", pending)
        c3.metric("Accepted", accepted)
        c4.metric("Delivered", delivered)

        # 圖表
        colA,colB=st.columns(2)
        with colA:
            st.write("**按狀態分佈**")
            if "Status" in df_jobs.columns:
                status_counts=df_jobs["Status"].astype(str).value_counts()
                st.bar_chart(status_counts)
                st.dataframe(status_counts)
        with colB:
            st.write("**按客戶Top10**")
            if "Client" in df_jobs.columns:
                client_counts=df_jobs["Client"].astype(str).value_counts().head(10)
                st.bar_chart(client_counts)

        st.divider()
        st.write("**按日期趨勢 (近30日)**")
        try:
            if "Date" in df_jobs.columns:
                df_jobs["Date_parsed"]=pd.to_datetime(df_jobs["Date"], errors="coerce")
                trend=df_jobs.dropna(subset=["Date_parsed"]).groupby(df_jobs["Date_parsed"].dt.date).size()
                trend=trend.tail(30)
                st.line_chart(trend)
        except Exception as e:
            st.write(f"趨勢圖解析失敗: {e}")

        st.divider()
        st.write("**最新10張任務**")
        show_table(df_jobs.head(10) if "JobID" not in df_jobs.columns else df_jobs.sort_values(by="JobID", ascending=False).head(10))

        # POD統計
        st.divider()
        st.write(f"**POD統計**: PODs表多圖 {len(df_pods)} 張 | Jobs表舊圖 {df_jobs[find_pod_col(df_jobs)].astype(str).str.len().gt(500).sum() if find_pod_col(df_jobs) else 0} 張")

# ==================== 2) 任務管理 + ZIP (V25保留) ====================
task_tab_idx=1 if role=="admin" else 1 if role=="op" else None
if task_tab_idx is not None and role in ["admin","op"]:
    with tabs[task_tab_idx]:
        st.subheader("📋 所有任務 - 已按Job#排序")
        show_table(df_jobs)
        st.divider()
        st.subheader("📸 POD下載區 - 多圖 + 一鍵ZIP")
        all_pods=[]
        pod_col=find_pod_col(df_jobs)
        if pod_col:
            for _, r in df_jobs.iterrows():
                jid=str(r.get("JobID","")).strip()
                b64=str(r.get(pod_col,""))
                if len(b64)>500:
                    fb=safe_b64_decode(b64)
                    if fb:
                        pod_time=str(r.get("PodTime_HKT","") or "")
                        all_pods.append({"JobID":jid, "FileName":f"{jid}_POD_{pod_time.replace(':','-').replace(' ','_')}.jpg" if pod_time else f"{jid}_POD_OLD.jpg", "bytes":fb, "time":pod_time})
        if not df_pods.empty:
            for _, r in df_pods.iterrows():
                jid=str(r.get("JobID","")).strip()
                b64=str(r.get("Base64_Image","") or r.get("COL_4",""))
                fb=safe_b64_decode(b64)
                if fb and jid:
                    all_pods.append({"JobID":jid, "FileName":str(r.get("FileName","") or f"{jid}_POD_{r.get('POD_ID','')}.jpg"), "bytes":fb, "time":str(r.get("PodTime_HKT",""))})

        if not all_pods:
            st.info("暫無有效POD")
        else:
            df_all=pd.DataFrame(all_pods).sort_values(by="JobID", ascending=False)
            st.markdown(f"**共 {len(df_all)} 張POD，{df_all['JobID'].nunique()} 個Job**")
            zip_all_buf=io.BytesIO()
            with zipfile.ZipFile(zip_all_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                for p in all_pods: zf.writestr(p["FileName"], p["bytes"])
            zip_all_buf.seek(0)
            st.download_button(f"📦 打包下載全部POD ZIP ({len(all_pods)}張)", data=zip_all_buf, file_name=f"ALL_PODs_{now_hk().strftime('%Y%m%d_%H%M')}.zip", mime="application/zip", type="primary", key="zip_all_v26")
            st.divider()
            for jid, group in df_all.groupby("JobID", sort=False):
                with st.container(border=True):
                    c_head,c_zip=st.columns([3,1])
                    c_head.markdown(f"### {jid} - {len(group)}張相")
                    zip_buf=io.BytesIO()
                    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                        for _, p in group.iterrows(): zf.writestr(p["FileName"], p["bytes"])
                    zip_buf.seek(0)
                    c_zip.download_button(f"📦 打包此Job ({len(group)}張)", data=zip_buf, file_name=f"{jid}_PODs_{now_hk().strftime('%Y%m%d')}.zip", mime="application/zip", key=f"zip_{jid}_{uuid.uuid4().hex[:4]}_v26", use_container_width=True)
                    cols=st.columns(3)
                    for i, (_, p) in enumerate(group.iterrows()):
                        with cols[i%3]:
                            st.image(p["bytes"], caption=p["FileName"], use_container_width=True)
                            st.download_button(f"⬇️ 下載", data=p["bytes"], file_name=p["FileName"], mime="image/jpeg", key=f"dl_{jid}_{i}_{uuid.uuid4().hex[:4]}_v26", use_container_width=True)

# ==================== 3) 上傳Order 補返Excel批量 + Template ====================
order_tab_idx=2 if role=="admin" else 2 if role=="op" else None
if order_tab_idx is not None:
    with tabs[order_tab_idx]:
        st.subheader("📤 上傳Order - 手動 + Excel批量")
        tab_manual, tab_excel = st.tabs(["✍️ 手動建立", "📊 Excel批量上傳"])

        with tab_manual:
            with st.form("manual_v26", clear_on_submit=True):
                client=st.text_input("客戶*"); date=st.date_input("日期", value=today_hk())
                from_loc=st.text_input("起點"); to_loc=st.text_input("終點")
                qty=st.number_input("數量", 1.0); notes=st.text_input("備註")
                if st.form_submit_button("建立", type="primary"):
                    if client.strip():
                        jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                        append_row("Jobs",[jid,str(date),client,from_loc,to_loc,"","Pending",str(qty),notes,operator,now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"),"","",""])
                        st.success(jid); st.rerun()

        with tab_excel:
            st.write("**Template 下載 + 批量上傳 (支持多張Order)**")
            # Template生成
            template_df=pd.DataFrame(columns=["Date","Client","From","To","QTY","Notes"])
            template_df.loc[0]=[str(today_hk()),"Sandy","HK","KLN","1","Sample"]
            template_df.loc[1]=[str(today_hk()),"ABC Ltd","Central","TST","2","Urgent"]
            buf_template=io.BytesIO()
            template_df.to_excel(buf_template, index=False, engine='openpyxl')
            buf_template.seek(0)
            st.download_button("📥 下載Excel Template (範本)", data=buf_template, file_name="TMS_Order_Template.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", key="dl_template_v26")

            st.divider()
            excel_file=st.file_uploader("上傳已填好嘅Excel", type=["xlsx","xls"], key="excel_uploader_v26")
            if excel_file:
                try:
                    df_excel=pd.read_excel(excel_file)
                    st.write(f"讀到 {len(df_excel)} 行")
                    st.dataframe(df_excel.head(20))
                    # 欄位檢查
                    required=["Client"]
                    if "Client" not in df_excel.columns and "CLIENT" not in df_excel.columns:
                        st.error("Excel必須有 Client 欄 (Date, Client, From, To, QTY, Notes)")
                    else:
                        # 正規化欄名
                        df_excel.columns=[c.strip() for c in df_excel.columns]
                        # 兼容大小寫
                        col_map={c.lower():c for c in df_excel.columns}
                        def get_col(name): return df_excel[col_map.get(name.lower(), name)] if col_map.get(name.lower()) in df_excel.columns else ""

                        if st.button(f"✅ 確認批量建立 {len(df_excel)} 張Order", type="primary", key="batch_create_v26"):
                            rows_to_add=[]
                            for _, row in df_excel.iterrows():
                                if str(row.get(col_map.get('client','Client'),"")).strip()=="": continue
                                jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                                date_val=str(row.get(col_map.get('date','Date'), str(today_hk())))
                                client_val=str(row.get(col_map.get('client','Client'), ""))
                                from_val=str(row.get(col_map.get('from','From'), ""))
                                to_val=str(row.get(col_map.get('to','To'), ""))
                                qty_val=str(row.get(col_map.get('qty','QTY'), "1"))
                                notes_val=str(row.get(col_map.get('notes','Notes'), ""))
                                rows_to_add.append([jid,date_val,client_val,from_val,to_val,"","Pending",qty_val,notes_val,operator,now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"),"","",""])
                            if rows_to_add:
                                append_rows("Jobs", rows_to_add)
                                st.success(f"✅ 已批量建立 {len(rows_to_add)} 張Order")
                                st.rerun()
                except Exception as e:
                    st.error(f"Excel讀取失敗: {e}")

# ==================== 上傳POD多圖 ====================
pod_tab_idx=3 if role=="admin" else 3 if role=="op" else 2
with tabs[pod_tab_idx]:
    st.subheader("📸 上傳POD - 支持一個Job多張")
    my_pool=df_jobs if role in ["admin","op"] else df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else df_jobs
    my_pool=my_pool.sort_values(by="JobID", ascending=False) if not my_pool.empty else my_pool
    if my_pool.empty: st.info("無任務")
    else:
        jid=st.selectbox("選擇JobID", my_pool["JobID"].astype(str).tolist(), key="sel_jid_v26")
        pod_files=st.file_uploader("拍照 / 選圖 (可一次選多張)", type=["jpg","png","jpeg"], accept_multiple_files=True, key="multi_pod_v26")
        if pod_files:
            for f in pod_files: st.image(f, width=200, caption=f.name)
        if pod_files and st.button(f"上傳 {len(pod_files)} 張到 {jid}", type="primary", key="upload_multi_v26"):
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
                    if len(data_url)>=49000: continue
                    fname=f"{jid}_POD_{now_hk().strftime('%Y%m%d_%H%M%S')}_{idx_file+1}.jpg"
                    pod_id=f"POD-{now_hk().strftime('%m%d%H%M%S')}-{uuid.uuid4().hex[:4]}"
                    append_row("PODs",[pod_id, jid, fname, str(len(fb)), data_url, operator, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"), now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")])
                    ws=get_sh().worksheet("Jobs"); cell=ws.find(jid)
                    if cell:
                        ws.update_cell(cell.row, 7, "Delivered")
                        ws.update_cell(cell.row, 13, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"))
                        ws.update_cell(cell.row, 14, data_url)
                        ws.update_cell(cell.row, 12, f"MULTI_{len(pod_files)}")
                    success+=1
                    progress.progress((idx_file+1)/len(pod_files))
                except Exception as e: st.error(f"{pod_file.name} 失敗 {e}")
            st.cache_data.clear()
            st.success(f"✅ 成功上傳 {success}/{len(pod_files)} 張到 {jid}")

# 其他頁
if role=="driver":
    with tabs[0]:
        st.subheader("🟢 可接任務池")
        if not df_jobs.empty:
            mask=df_jobs["Driver"].astype(str).str.strip()==""
            avail=df_jobs[mask | (df_jobs["Status"].astype(str).str.lower()=="pending")]
            show_table(avail)
            for _, row in avail.iterrows():
                jid=row["JobID"]
                if st.button(f"搶 {jid}", key=f"claim_{jid}_v26", type="primary"):
                    ws=get_sh().worksheet("Jobs"); cell=ws.find(jid)
                    if cell: ws.update_cell(cell.row, 6, operator); ws.update_cell(cell.row, 7, "Accepted"); st.cache_data.clear(); st.rerun()
    with tabs[1]:
        st.subheader("🚚 我的任務"); my=df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else pd.DataFrame(); show_table(my)

# 維護頁
if role in ["admin","op"]:
    maint_idx=6 if role=="admin" else 5
    if len(tabs)>maint_idx:
        with tabs[maint_idx]:
            st.subheader("🧹 維護 - 清舊爛Base64")
            if st.button("🔍 掃描爛POD"):
                pod_col=find_pod_col(df_jobs)
                broken=[]
                if pod_col:
                    for _, r in df_jobs.iterrows():
                        b64=str(r.get(pod_col,""))
                        if len(b64)>40000 and safe_b64_decode(b64) is None:
                            broken.append(r.get("JobID"))
                if broken: st.error(f"發現 {len(broken)} 個爛POD: {', '.join(broken)}")
                else: st.success("無爛POD")
            if st.button("🗑️ 一鍵清除爛POD Base64", type="primary"):
                ws=get_sh().worksheet("Jobs")
                headers=ws.row_values(1)
                pod_col_idx=14
                for i,h in enumerate(headers):
                    if "IMAGE" in h.upper(): pod_col_idx=i+1
                all_vals=ws.get_all_values()
                cleaned=0
                for row_idx in range(1, len(all_vals)):
                    if len(all_vals[row_idx])>=pod_col_idx:
                        b64=all_vals[row_idx][pod_col_idx-1]
                        if len(b64)>40000 and safe_b64_decode(b64) is None:
                            ws.update_cell(row_idx+1, pod_col_idx, "")
                            cleaned+=1
                st.cache_data.clear()
                st.success(f"已清理 {cleaned} 個")
