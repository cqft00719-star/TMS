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

st.set_page_config(page_title="TMS V13", layout="wide", page_icon="🚚")

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
        df = pd.DataFrame(ws.get_all_records())
        # 清理空格
        for c in df.columns:
            if df[c].dtype == object:
                df[c] = df[c].astype(str).str.strip()
                df[c] = df[c].replace({"nan":"","None":"","null":""})
        return df
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
            if defaults: ws.append_rows(defaults, value_input_option='USER_ENTERED')
            existing.append(title)
    create_if_not_exist("Users", ["username","password","name","role"], [["admin","admin123","Admin","admin"], ["John","John123","John","driver"], ["op01","op01","OP01","op"], ["Mo","Mo123","Mo","driver"]])
    create_if_not_exist("Jobs", ["JobID","Date","Client","From","To","Driver","Status","QTY","Notes","CreatedBy","CreatedAt_HKT","POD_Link","PodTime_HKT","POD_Image"])
    create_if_not_exist("Expenses", ["ExpenseID","Date","JobID","Driver","Type","Amount","Remark","CreatedAt_HKT"])
    create_if_not_exist("Config", ["Key","Value"], [["POD_Folder_ID",""]])

def get_template_excel():
    df_template = pd.DataFrame([{"Client":"Office","From":"21C","To":"18A1","QTY":3,"Date":str(today_hk()),"Notes":"送客3單"}])
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
    return s

if "user" not in st.session_state: st.session_state.user=None
if not st.session_state.user:
    st.title("TMS - 登入")
    df_u=load_df("Users")
    if df_u.empty: df_u=pd.DataFrame([{"username":"admin","password":"admin123","name":"Admin","role":"admin"}])
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
            st.metric("總任務", len(df_jobs))
            st.dataframe(df_jobs.sort_values(by="JobID", ascending=False), use_container_width=True)

if role in ["admin","op"]:
    idx = 1 if role=="admin" else 0
    with tabs[idx]:
        st.subheader("📋 所有任務 - Job#排序 + 下載POD")
        if st.button("🔄 刷新"): st.cache_data.clear(); st.rerun()
        if not df_jobs.empty:
            df_show = df_jobs.sort_values(by="JobID", ascending=False)
            st.dataframe(df_show, use_container_width=True, height=350)
            st.write("**📸 POD下載區**")
            for _, r in df_show.iterrows():
                jid=str(r.get("JobID",""))
                b64img=str(r.get("POD_Image","")) if "POD_Image" in r else ""
                link=str(r.get("POD_Link",""))
                if b64img.startswith("data:") or link.startswith("http"):
                    with st.container(border=True):
                        c1,c2=st.columns([3,1])
                        c1.write(f"**{jid}** {r.get('Client','')} {r.get('Status','')}")
                        if b64img.startswith("data:"):
                            try:
                                _, b64data = b64img.split(",",1)
                                fb = base64.b64decode(b64data)
                                c1.image(fb, width=200)
                                c2.download_button(f"⬇️ 下載 {jid}_POD.jpg", fb, file_name=f"{jid}_POD.jpg", mime="image/jpeg", key=f"dl_{jid}")
                            except: pass
                        elif link.startswith("http"):
                            c2.link_button("🔗 查看", link)

if role in ["admin","op"]:
    idx = 2 if role=="admin" else 1
    with tabs[idx]:
        st.subheader("📤 上傳Order")
        st.download_button("📥 下載模板", data=get_template_excel(), file_name="TMS_Template.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with st.form("manual", clear_on_submit=True):
            c1,c2,c3=st.columns(3)
            client=c1.text_input("客戶*"); date=c1.date_input("日期", value=today_hk())
            from_loc=c2.text_input("起點"); to_loc=c2.text_input("終點")
            qty=c3.number_input("數量", 1.0); notes=c3.text_input("備註")
            if st.form_submit_button("建立"):
                if client.strip():
                    jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                    append_row("Jobs",[jid,str(date),client,from_loc,to_loc,"","Pending",str(qty),notes,operator,now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"),"","",""])
                    st.success(jid); st.rerun()

# ===== 關鍵修復：Driver可接任務池 =====
if role=="driver":
    with tabs[0]:
        st.subheader("🟢 可接任務池 - 已修復")
        if st.button("🔄 刷新任務池"): st.cache_data.clear(); st.rerun()
        st.caption(f"Google Sheet有 {len(df_jobs)} 張 | 當前用戶 {operator}")
        if df_jobs.empty:
            st.info("暫無")
        else:
            # 強力filter：Status=Pending 或 Driver為空/nan
            def is_pending(s): return str(s).strip().lower()=="pending"
            def is_empty_driver(d): return str(d).strip()=="" or str(d).strip().lower() in ["nan","none","null"]
            # 只要 Pending 或 Driver空 就算可接
            mask_pending = df_jobs["Status"].apply(is_pending) if "Status" in df_jobs.columns else pd.Series([False]*len(df_jobs))
            mask_empty = df_jobs["Driver"].apply(is_empty_driver) if "Driver" in df_jobs.columns else pd.Series([False]*len(df_jobs))
            avail = df_jobs[mask_pending | mask_empty].copy()
            # 排除已Delivered/Completed但Driver空的異常 (可選)
            # avail = avail[~avail["Status"].str.lower().isin(["delivered","completed"]) | mask_empty]
            avail = avail.sort_values(by="JobID", ascending=False)
            st.write(f"可接：{len(avail)} 張")
            st.dataframe(avail, use_container_width=True)
            if avail.empty:
                st.warning("篩選後係0，檢查下Sheet Status係咪真係Pending")
                st.dataframe(df_jobs[["JobID","Driver","Status"]], use_container_width=True)
            for _, row in avail.iterrows():
                jid=row["JobID"]
                with st.container(border=True):
                    c1,c2=st.columns([4,1])
                    c1.markdown(f"**{jid}** | {row.get('Client','')} | {row.get('From','')} -> {row.get('To','')} | QTY:{row.get('QTY','')}")
                    if c2.button("搶此單", key=f"claim_{jid}", type="primary", use_container_width=True):
                        ws=get_sh().worksheet("Jobs"); cell=ws.find(jid)
                        if not cell: st.error("搵唔到Job")
                        else:
                            curr_driver = ws.cell(cell.row, 6).value
                            if str(curr_driver).strip() not in ["","None","nan"]:
                                st.error(f"已被 {curr_driver} 搶")
                            else:
                                ws.update_cell(cell.row, 6, operator)
                                ws.update_cell(cell.row, 7, "Accepted")
                                st.cache_data.clear()
                                st.success(f"{jid} 搶單成功！去『我的任務』睇"); st.rerun()
    with tabs[1]:
        st.subheader("🚚 我的任務")
        if st.button("🔄 刷新"): st.cache_data.clear(); st.rerun()
        my_jobs = df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator].sort_values(by="JobID", ascending=False) if not df_jobs.empty else pd.DataFrame()
        st.dataframe(my_jobs, use_container_width=True)
        if not my_jobs.empty:
            sel=st.selectbox("改狀態", my_jobs["JobID"].tolist())
            ns=st.selectbox("新狀態", ["Accepted","InTransit","Delivered","Completed"])
            if st.button("更新"):
                ws=get_sh().worksheet("Jobs"); cell=ws.find(sel)
                if cell: ws.update_cell(cell.row, 7, ns); st.cache_data.clear(); st.rerun()

# POD上傳
with tabs[3 if role=="admin" else 2 if role=="op" else 2]:
    st.subheader("📸 上傳POD")
    my_jobs_pool = df_jobs if role in ["admin","op"] else df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else pd.DataFrame()
    my_jobs_pool = my_jobs_pool.sort_values(by="JobID", ascending=False) if not my_jobs_pool.empty else my_jobs_pool
    if my_jobs_pool.empty: st.info("無任務")
    else:
        jid=st.selectbox("選擇JobID", my_jobs_pool["JobID"].astype(str).tolist(), key="pod_v13")
        pod_file=st.file_uploader("拍照", type=["jpg","png","jpeg"], key="pod_v13")
        if pod_file: st.image(pod_file, width=300, caption=f"{jid}_POD_{pod_file.name}")
        if pod_file and st.button("上傳POD", type="primary"):
            try:
                ws=get_sh().worksheet("Jobs"); cell=ws.find(jid)
                from PIL import Image
                img = Image.open(io.BytesIO(pod_file.getvalue())); img.thumbnail((1200,1200))
                buf=io.BytesIO(); img.save(buf, format="JPEG", quality=80)
                fb=buf.getvalue(); b64=base64.b64encode(fb).decode(); data_url=f"data:image/jpeg;base64,{b64}"
                save_name=f"{jid}_POD_{now_hk().strftime('%Y%m%d_%H%M')}.jpg"
                if cell:
                    ws.update_cell(cell.row, 12, "BASE64_SAVED"); ws.update_cell(cell.row, 13, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")); ws.update_cell(cell.row, 7, "Delivered"); ws.update_cell(cell.row, 14, data_url[:49000])
                st.cache_data.clear()
                st.success(f"✅ {save_name} 已保存")
                st.download_button(f"⬇️ 下載 {save_name}", fb, file_name=save_name, mime="image/jpeg", type="primary")
            except Exception as e:
                st.error(str(e)); import traceback; st.code(traceback.format_exc())

with tabs[4 if role=="admin" else 3 if role=="op" else 3]:
    st.subheader("💰 雜費"); st.dataframe(df_exp, use_container_width=True)

if role=="admin":
    with tabs[5]:
        st.subheader("用戶"); st.dataframe(load_df("Users"), use_container_width=True)
