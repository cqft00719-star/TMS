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

st.set_page_config(page_title="TMS 運輸系統 V12", layout="wide", page_icon="🚚")

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
    create_if_not_exist("Jobs", ["JobID","Date","Client","From","To","Driver","Status","QTY","Notes","CreatedBy","CreatedAt_HKT","POD_Link","PodTime_HKT","POD_Image"])
    create_if_not_exist("Expenses", ["ExpenseID","Date","JobID","Driver","Type","Amount","Remark","CreatedAt_HKT"])
    create_if_not_exist("Config", ["Key","Value"], [["POD_Folder_ID",""]])

def get_template_excel():
    df_template = pd.DataFrame([
        {"Client":"ABC公司","From":"葵涌","To":"中環","QTY":2,"Date":str(today_hk()),"Notes":"易碎"},
        {"Client":"XYZ公司","From":"觀塘","To":"沙田","QTY":1,"Date":str(today_hk()),"Notes":""},
    ])
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

# Dashboard
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
            df_show = df_jobs.sort_values(by="JobID", ascending=False)
            st.dataframe(df_show, use_container_width=True)

# 任務管理 - 可下載版
if role in ["admin","op"]:
    idx = 1 if role=="admin" else 0
    with tabs[idx]:
        st.subheader("📋 所有任務 - 按Job#排序 + 下載POD")
        if df_jobs.empty:
            st.info("暫無任務")
        else:
            df_show = df_jobs.sort_values(by="JobID", ascending=False) if "JobID" in df_jobs.columns else df_jobs
            st.dataframe(df_show[["JobID","Date","Client","From","To","Driver","Status","POD_Link","PodTime_HKT"]], use_container_width=True, height=350)

            st.divider()
            st.write("**📸 POD下載區 - 文件名已加Job# Prefix排序**")
            for _, r in df_show.iterrows():
                jid = str(r.get("JobID","")).strip()
                if not jid: continue
                link = str(r.get("POD_Link","")).strip()
                b64img = str(r.get("POD_Image","")).strip() if "POD_Image" in df_show.columns else ""
                has_pod = link not in ["","BASE64_SAVED"] or b64img.startswith("data:")

                if has_pod:
                    with st.container(border=True):
                        c1,c2,c3 = st.columns([3,2,2])
                        c1.markdown(f"**{jid}** | {r.get('Client','')} | {r.get('Date','')} | {r.get('Status','')}")
                        safe_time = str(r.get('PodTime_HKT','')).replace(':','-').replace(' ','_').replace('/','-')
                        file_name = f"{jid}_POD_{safe_time}.jpg"

                        # Base64圖片
                        if b64img.startswith("data:"):
                            try:
                                header, b64data = b64img.split(",",1)
                                file_bytes = base64.b64decode(b64data)
                                c1.image(file_bytes, width=250, caption=file_name)
                                c2.download_button(
                                    label=f"⬇️ 下載 {file_name}",
                                    data=file_bytes,
                                    file_name=file_name,
                                    mime="image/jpeg",
                                    key=f"dl_b64_{jid}_{safe_time}",
                                    type="primary",
                                    use_container_width=True
                                )
                            except Exception as ex:
                                c2.error(f"圖片讀取失敗 {ex}")
                        # Drive Link
                        elif link and link.startswith("http"):
                            c1.write(f"Drive: {link[:60]}...")
                            c2.link_button("🔗 查看原圖", link, use_container_width=True)
                            c3.link_button("📁 開Folder", f"https://drive.google.com/drive/folders/{extract_folder_id_from_url(load_df('Config')[load_df('Config')['Key']=='POD_Folder_ID']['Value'].values[0]) if not load_df('Config').empty else ''}", use_container_width=True)

# 上傳Order
if role in ["admin","op"]:
    idx = 2 if role=="admin" else 1
    with tabs[idx]:
        st.subheader("📤 上傳Order")
        st.download_button("📥 下載模板.xlsx", data=get_template_excel(), file_name=f"TMS_Template_{today_hk()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
        with st.form("manual_order", clear_on_submit=True):
            c1,c2,c3=st.columns(3)
            client=c1.text_input("客戶*"); date=c1.date_input("日期", value=today_hk())
            from_loc=c2.text_input("起點"); to_loc=c2.text_input("終點")
            qty=c3.number_input("數量", 0.0, 100.0, 1.0); notes=c3.text_input("備註")
            if st.form_submit_button("✍️ 建立Order", use_container_width=True):
                if client.strip():
                    jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                    append_row("Jobs",[jid,str(date),client.strip(),from_loc,to_loc,"","Pending",str(qty),notes,operator,now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"),"","",""])
                    st.success(f"已建立 {jid}"); st.rerun()
        up=st.file_uploader("Excel批量上傳", type=["xlsx"], key="up_excel")
        if up:
            df=pd.read_excel(up); st.dataframe(df.head())
            if st.button("🚀 開始導入", type="primary"):
                rows=[]
                for _, r in df.iterrows():
                    if pd.isna(r.iloc[0]) or str(r.iloc[0]).strip()=="": continue
                    jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                    rows.append([jid, str(today_hk()), str(r.iloc[0]), str(r.iloc[1]) if len(r)>1 else "", str(r.iloc[2]) if len(r)>2 else "", "", "Pending", "1", "", operator, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"), "", "", ""])
                if rows: append_rows("Jobs", rows); st.success(f"導入{len(rows)}張"); st.rerun()

# 司機搶單
if role=="driver":
    with tabs[0]:
        st.subheader("🟢 可接任務池 - 按Job#排序")
        avail = df_jobs[(df_jobs["Status"].astype(str)=="Pending") | (df_jobs["Driver"].astype(str)=="")] if not df_jobs.empty else pd.DataFrame()
        if avail.empty: st.success("暫無待搶")
        else:
            avail = avail.sort_values(by="JobID", ascending=False)
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
    with tabs[1]:
        st.subheader("🚚 我的任務")
        my_jobs = df_jobs[df_jobs["Driver"].astype(str)==operator].sort_values(by="JobID", ascending=False) if not df_jobs.empty else pd.DataFrame()
        st.dataframe(my_jobs, use_container_width=True)

# POD上傳 - V12可下載 + Job# Prefix
with tabs[3 if role=="admin" else 2 if role=="op" else 2]:
    st.subheader("📸 上傳POD - 自動Job# Prefix + 可下載")
    my_jobs_pool = df_jobs if role in ["admin","op"] else df_jobs[df_jobs["Driver"].astype(str)==operator] if not df_jobs.empty else pd.DataFrame()
    if my_jobs_pool.empty:
        st.info("無任務可上傳")
    else:
        my_jobs_pool = my_jobs_pool.sort_values(by="JobID", ascending=False)
        jid=st.selectbox("選擇JobID (已按Job#排序)", my_jobs_pool["JobID"].astype(str).tolist(), key="pod_job_v12")
        pod_file=st.file_uploader("拍照或上傳", type=["jpg","png","pdf","jpeg"], key="pod_file_v12")
        if pod_file:
            save_preview = f"{jid}_POD_{now_hk().strftime('%Y%m%d_%H%M')}_{pod_file.name}"
            st.image(pod_file, width=350, caption=f"預覽: {save_preview}")
            st.caption(f"儲存文件名: **{save_preview}**")
        if pod_file and st.button("上傳POD", type="primary", key="btn_pod_v12"):
            try:
                ws=get_sh().worksheet("Jobs"); cell=ws.find(jid)
                cfg = load_df("Config")
                fid = ""
                if not cfg.empty and "Value" in cfg.columns:
                    try: fid = extract_folder_id_from_url(cfg[cfg["Key"]=="POD_Folder_ID"]["Value"].values[0])
                    except: pass
                # 嘗試Drive
                drive_success=False
                try:
                    drive=get_drive_service()
                    media=MediaIoBaseUpload(io.BytesIO(pod_file.getvalue()), mimetype=pod_file.type, resumable=True)
                    save_name = f"{jid}_POD_{now_hk().strftime('%Y%m%d_%H%M')}_{pod_file.name}"
                    meta={"name": save_name}
                    if fid and len(fid)>10: meta["parents"]=[fid]
                    uploaded=drive.files().create(body=meta, media_body=media, fields='id, webViewLink', supportsAllDrives=True).execute()
                    try: drive.permissions().create(fileId=uploaded["id"], body={"role":"reader","type":"anyone"}, supportsAllDrives=True).execute()
                    except: pass
                    link=uploaded.get("webViewLink","")
                    if cell:
                        ws.update_cell(cell.row, 12, link)
                        ws.update_cell(cell.row, 13, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"))
                        ws.update_cell(cell.row, 7, "Delivered")
                    st.cache_data.clear()
                    st.success(f"✅ Drive上傳成功: {save_name}")
                    st.link_button("🔗 查看Drive相片", link)
                    drive_success=True
                except Exception as e:
                    if "quota" in str(e).lower() or "403" in str(e) or "storage" in str(e).lower():
                        st.warning("免費Gmail Drive quota限制，自動改為存入Google Sheet，照樣可下載")
                    else:
                        # 其他錯誤都 fallback
                        st.warning(f"Drive上傳失敗 ({e})，改存Sheet")

                if not drive_success:
                    try:
                        from PIL import Image
                        img = Image.open(io.BytesIO(pod_file.getvalue()))
                        img.thumbnail((1200,1200))
                        buf=io.BytesIO()
                        img.save(buf, format="JPEG", quality=80)
                        file_bytes = buf.getvalue()
                    except:
                        file_bytes = pod_file.getvalue()
                    b64 = base64.b64encode(file_bytes).decode()
                    data_url = f"data:image/jpeg;base64,{b64}"
                    save_name = f"{jid}_POD_{now_hk().strftime('%Y%m%d_%H%M')}.jpg"
                    if cell:
                        ws.update_cell(cell.row, 12, "BASE64_SAVED")
                        ws.update_cell(cell.row, 13, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"))
                        ws.update_cell(cell.row, 7, "Delivered")
                        # Sheet cell上限50000，截斷
                        ws.update_cell(cell.row, 14, data_url[:49000])
                    st.cache_data.clear()
                    st.success(f"✅ 已保存 {save_name} 到Sheet (Job# Prefix)")
                    st.download_button("⬇️ 立即下載此POD", data=file_bytes, file_name=save_name, mime="image/jpeg", type="primary", use_container_width=True)
                    st.image(file_bytes, width=300)
            except Exception as e:
                st.error(f"上傳失敗: {e}")
                import traceback; st.code(traceback.format_exc())

# 雜費
with tabs[4 if role=="admin" else 3 if role=="op" else 3]:
    st.subheader("💰 雜費")
    st.dataframe(df_exp.sort_values(by="ExpenseID", ascending=False) if not df_exp.empty else df_exp, use_container_width=True)
    with st.form("exp_form", clear_on_submit=True):
        c1,c2,c3=st.columns(3)
        jid_list=[""]+ (df_jobs["JobID"].astype(str).tolist() if not df_jobs.empty else [])
        jid=c1.selectbox("關聯JobID", jid_list)
        etype=c2.selectbox("類型", ["隧道費","泊車費","油費","雜費","其他"])
        amt=c3.number_input("金額 $", 0.0, step=1.0); remark=st.text_input("備註")
        if st.form_submit_button("提交雜費"):
            eid=f"EXP-{now_hk().strftime('%m%d%H%M')}"
            append_row("Expenses", [eid, str(today_hk()), jid, operator, etype, amt, remark, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")])
            st.success("已提交"); st.rerun()

# 用戶管理
if role=="admin":
    with tabs[5]:
        st.subheader("用戶管理"); st.dataframe(load_df("Users"), use_container_width=True)
        with st.form("add_user", clear_on_submit=True):
            nu=st.text_input("用戶名"); np=st.text_input("密碼"); nn=st.text_input("姓名"); nr=st.selectbox("角色",["driver","op","admin"])
            if st.form_submit_button("新增用戶"):
                append_row("Users", [nu,np,nn,nr]); st.success("已新增"); st.rerun()
