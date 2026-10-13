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

st.set_page_config(page_title="TMS V14", layout="wide", page_icon="🚚")

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

# V14 核心修復：用get_all_values，避開Base64太長導致get_all_records死掉
@st.cache_data(ttl=30)
def load_df(tab):
    try:
        ws = get_sh().worksheet(tab)
        all_vals = ws.get_all_values()
        if len(all_vals) < 1:
            return pd.DataFrame()
        headers = [h.strip() for h in all_vals[0]]
        # 過濾全空行
        data_rows = []
        for r in all_vals[1:]:
            if any(str(c).strip()!="" for c in r):
                # 補齊欄位長度
                if len(r) < len(headers):
                    r = r + [""]*(len(headers)-len(r))
                data_rows.append(r[:len(headers)])
        df = pd.DataFrame(data_rows, columns=headers)
        # 輕量清洗，只洗關鍵欄
        for col in ["JobID","Driver","Status","Client"]:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
        df = df.replace({"nan":"","None":"","#N/A":""})
        return df
    except Exception as e:
        st.error(f"讀取 {tab} 失敗: {e}")
        import traceback; st.code(traceback.format_exc())
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
            ws = sh.add_worksheet(title=title, rows=1000, cols=20)
            ws.update('A1', [headers])
            if defaults: ws.append_rows(defaults, value_input_option='USER_ENTERED')
            existing.append(title)
    create_if_not_exist("Users", ["username","password","name","role"], [["admin","admin123","Admin","admin"], ["John","John123","John","driver"], ["op01","op01","OP01","op"], ["Mo","Mo123","Mo","driver"]])
    create_if_not_exist("Jobs", ["JobID","Date","Client","From","To","Driver","Status","QTY","Notes","CreatedBy","CreatedAt_HKT","POD_Link","PodTime_HKT","POD_Image"])
    create_if_not_exist("Expenses", ["ExpenseID","Date","JobID","Driver","Type","Amount","Remark","CreatedAt_HKT"])
    create_if_not_exist("Config", ["Key","Value"], [["POD_Folder_ID",""]])

def get_template_excel():
    df_template = pd.DataFrame([{"Client":"Sandy","From":"21C","To":"18A1","QTY":3,"Date":str(today_hk()),"Notes":"3號送客"}])
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_template.to_excel(writer, index=False)
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
    st.title("TMS V14 - 登入")
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
st.sidebar.write(f"Sheet ID:...{st.secrets['connections']['gsheets']['spreadsheet'][-6:]}")
if st.sidebar.button("登出"): st.session_state.user=None; st.cache_data.clear(); st.rerun()
if st.sidebar.button("🔄 清除Cache刷新"): st.cache_data.clear(); st.rerun()

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
        st.write(f"Jobs表讀到 {len(df_jobs)} 行")
        if not df_jobs.empty: st.dataframe(df_jobs.sort_values(by="JobID", ascending=False), use_container_width=True)

if role in ["admin","op"]:
    idx = 1 if role=="admin" else 0
    with tabs[idx]:
        st.subheader("📋 所有任務 - Job#排序 + 下載POD")
        st.write(f"共 {len(df_jobs)} 張單")
        if df_jobs.empty:
            st.warning("讀唔到Jobs，試下按左邊『清除Cache刷新』")
            # 顯示原始資料debug
            try:
                ws=get_sh().worksheet("Jobs"); st.write(ws.get_all_values()[:5])
            except Exception as e: st.error(str(e))
        else:
            df_show = df_jobs.sort_values(by="JobID", ascending=False) if "JobID" in df_jobs.columns else df_jobs
            st.dataframe(df_show, use_container_width=True, height=400)
            st.divider()
            st.write("**📸 POD下載區 - Job# Prefix**")
            for _, r in df_show.iterrows():
                jid=str(r.get("JobID","")).strip()
                if not jid: continue
                b64img=str(r.get("POD_Image","")) if "POD_Image" in df_show.columns else ""
                link=str(r.get("POD_Link",""))
                has_pod = b64img.startswith("data:") or (link.startswith("http") and link!="")
                if has_pod:
                    with st.container(border=True):
                        c1,c2=st.columns([3,1])
                        c1.markdown(f"**{jid}** | {r.get('Client','')} | Status:{r.get('Status','')}")
                        if b64img.startswith("data:"):
                            try:
                                _, bdata = b64img.split(",",1)
                                fb=base64.b64decode(bdata)
                                c1.image(fb, width=250, caption=f"{jid}_POD.jpg")
                                c2.download_button(f"⬇️ 下載 {jid}_POD.jpg", fb, file_name=f"{jid}_POD.jpg", mime="image/jpeg", key=f"dl_{jid}", type="primary", use_container_width=True)
                            except Exception as e: c1.write(f"圖片錯誤 {e}")
                        elif link.startswith("http"):
                            c2.link_button("🔗 查看Drive", link, use_container_width=True)

if role in ["admin","op"]:
    idx = 2 if role=="admin" else 1
    with tabs[idx]:
        st.subheader("📤 上傳Order")
        st.download_button("📥 下載模板", data=get_template_excel(), file_name="TMS_Template.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
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
        st.subheader("🟢 可接任務池 V14 - 已修復空白問題")
        st.write(f"Sheet共 {len(df_jobs)} 張，篩選中...")
        if df_jobs.empty:
            st.warning("暫無資料，按左邊刷新")
        else:
            def is_pending(s): return str(s).strip().lower()=="pending"
            def is_empty(d): return str(d).strip()=="" or str(d).strip().lower() in ["nan","none",""]
            mask_p = df_jobs["Status"].apply(is_pending) if "Status" in df_jobs.columns else pd.Series([False]*len(df_jobs))
            mask_e = df_jobs["Driver"].apply(is_empty) if "Driver" in df_jobs.columns else pd.Series([False]*len(df_jobs))
            avail = df_jobs[mask_p | mask_e].copy()
            avail = avail[~avail["Status"].astype(str).str.lower().isin(["delivered","completed"])] if "Status" in df_jobs.columns else avail
            avail = avail.sort_values(by="JobID", ascending=False) if not avail.empty else avail
            st.success(f"可接：{len(avail)} 張")
            st.dataframe(avail, use_container_width=True)
            for _, row in avail.iterrows():
                jid=row["JobID"]
                with st.container(border=True):
                    c1,c2=st.columns([4,1])
                    c1.markdown(f"**{jid}** | {row.get('Client','')} | {row.get('From','')}->{row.get('To','')}")
                    if c2.button("搶此單", key=f"claim_{jid}", type="primary", use_container_width=True):
                        ws=get_sh().worksheet("Jobs"); cell=ws.find(jid)
                        if not cell: st.error("搵唔到")
                        else:
                            curr=str(ws.cell(cell.row, 6).value or "").strip()
                            if curr!="": st.error(f"已被 {curr} 搶")
                            else:
                                ws.update_cell(cell.row, 6, operator); ws.update_cell(cell.row, 7, "Accepted")
                                st.cache_data.clear(); st.success("搶單成功"); st.rerun()
    with tabs[1]:
        st.subheader("🚚 我的任務")
        my_jobs = df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else pd.DataFrame()
        my_jobs = my_jobs.sort_values(by="JobID", ascending=False) if not my_jobs.empty else my_jobs
        st.dataframe(my_jobs, use_container_width=True)

with tabs[3 if role=="admin" else 2 if role=="op" else 2]:
    st.subheader("📸 上傳POD - Job# Prefix")
    my_pool = df_jobs if role in ["admin","op"] else df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else pd.DataFrame()
    my_pool = my_pool.sort_values(by="JobID", ascending=False) if not my_pool.empty else my_pool
    if my_pool.empty: st.info("無任務")
    else:
        jid=st.selectbox("選擇JobID", my_pool["JobID"].astype(str).tolist())
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
    st.subheader("💰 雜費"); st.dataframe(df_exp, use_container_width=True)

if role=="admin":
    with tabs[5]:
        st.subheader("用戶"); st.dataframe(load_df("Users"), use_container_width=True)
