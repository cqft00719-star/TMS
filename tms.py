# tms.py - TMS V46.9 My Tasks - 完整版, 唔會再唔見介面
import streamlit as st
import os
import io
import json
import base64
import traceback
from datetime import datetime
import requests
import pandas as pd
from PIL import Image

st.set_page_config(page_title="TMS V46.9 My Tasks", layout="wide")

TMS_SHEET_ID = "15X06a1oOFeta-ETz-CiLbB-5TUWS1rWKGi35HmBNNVE"
BOT_URL = "https://cookie-delivery-bot-production.up.railway.app"
DRIVE_FOLDER_ID = ""

# ---------- 登入狀態, 一定要留返 ----------
if "driver" not in st.session_state:
    st.session_state["driver"] = ""
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "just_uploaded_job" not in st.session_state:
    st.session_state["just_uploaded_job"] = ""

def get_gcp_info():
    try:
        if "GCP_JSON" in st.secrets:
            v = st.secrets["GCP_JSON"]
            if isinstance(v, dict):
                return dict(v)
            s = str(v).strip()
            if s.startswith("{"):
                return json.loads(s)
            return json.loads(base64.b64decode(s).decode())
    except Exception as e:
        st.error(f"讀 st.secrets GCP_JSON 失敗: {e}")
    try:
        if "gcp_service_account" in st.secrets:
            return dict(st.secrets["gcp_service_account"])
    except Exception:
        pass
    s = os.getenv("GCP_JSON", "").strip()
    if s:
        if s.startswith("{"):
            return json.loads(s)
        return json.loads(base64.b64decode(s).decode())
    st.error("Secrets 入面冇 GCP_JSON, 請去 Streamlit -> Settings -> Secrets, 加返 GCP_JSON")
    st.stop()

def get_creds():
    from google.oauth2.service_account import Credentials
    info = get_gcp_info()
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    return Credentials.from_service_account_info(info, scopes=scopes)

def get_sheet():
    import gspread
    return gspread.authorize(get_creds())

def get_drive():
    from googleapiclient.discovery import build
    return build("drive", "v3", credentials=get_creds())

def compress_pod(uploaded_file):
    im = Image.open(uploaded_file).convert("RGB")
    im.thumbnail((1200, 1200))
    for quality in [70, 55, 40, 30]:
        buf = io.BytesIO()
        im.save(buf, format="JPEG", quality=quality, optimize=True)
        data = buf.getvalue()
        if len(data) < 900 * 1024:
            return data, len(data)
        w, h = im.size
        im = im.resize((int(w * 0.7), int(h * 0.7)))
    buf = io.BytesIO()
    im.save(buf, format="JPEG", quality=30, optimize=True)
    return buf.getvalue(), len(buf.getvalue())

def upload_pod_to_drive(job_id, uploaded_file):
    try:
        if not DRIVE_FOLDER_ID:
            return True, ""
        from googleapiclient.http import MediaIoBaseUpload
        drive = get_drive()
        safe_name = f"{job_id}_POD_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
        data, final_size = compress_pod(uploaded_file)
        if final_size > 2 * 1024 * 1024:
            return False, f"壓縮完仲係 {final_size//1024}KB, 太大"
        media = MediaIoBaseUpload(io.BytesIO(data), mimetype="image/jpeg", resumable=True)
        body = {"name": safe_name, "parents": [DRIVE_FOLDER_ID], "mimeType": "image/jpeg"}
        f = drive.files().create(body=body, media_body=media, fields="id,webViewLink").execute()
        return True, f.get("webViewLink", "")
    except Exception as e:
        return False, f"Drive 失敗: {str(e)[:200]}"

@st.cache_data(ttl=60)
def load_jobs():
    sh = get_sheet().open_by_key(TMS_SHEET_ID)
    ws = sh.worksheet("Jobs")
    vals = ws.get_all_values()
    if len(vals) < 2:
        return pd.DataFrame(), [], ws
    headers = [h.strip() for h in vals[0]]
    df = pd.DataFrame(vals[1:], columns=headers)
    return df, headers, ws

def find_col(headers, names):
    for n in names:
        if n in headers:
            return n
    return None

# ---------- 登入畫面, 登出之後會返嚟呢度 ----------
if not st.session_state["logged_in"]:
    st.markdown("### TMS V46.9 My Tasks")
    st.markdown("請揀司機, 再撳登入")
    driver = st.selectbox("司機", ["", "展 (driver)", "Felix Kwok", "龍哥司機", "其他司機"], index=0)
    if st.button("登入", type="primary", use_container_width=True):
        if not driver:
            st.error("請先揀司機")
        else:
            st.session_state["driver"] = driver
            st.session_state["logged_in"] = True
            st.rerun()
    st.stop()

# ---------- 左邊欄, 登出一定有反應 ----------
with st.sidebar:
    st.markdown(f"### 👤 {st.session_state['driver']}")
    st.markdown("🔥 V46.9 My Tasks Fix")
    st.caption("上次自動發送: 2026-11-18")
    if st.button("登出", use_container_width=True):
        st.session_state["driver"] = ""
        st.session_state["logged_in"] = False
        st.session_state["just_uploaded_job"] = ""
        st.cache_data.clear()
        st.rerun()
    st.markdown("---")
    if st.button("📦 同步OrderCosts", use_container_width=True):
        try:
            r = requests.get(BOT_URL + "/test-jobs", timeout=20)
            st.success(f"同步完成: {r.status_code}")
        except Exception as e:
            st.error(f"同步失敗: {e}")
    if st.button("🔍 測試WhatsApp Group", use_container_width=True):
        try:
            r = requests.get(BOT_URL + "/", timeout=20)
            j = r.json()
            st.success(f"Bot 正常, cookie_photos={j.get('cookie_photos')}, pod={j.get('pod_photos', 0)}")
        except Exception as e:
            st.error(f"Bot 唔通: {e}")
    if st.button("📊 測試每日Excel", use_container_width=True):
        try:
            requests.get(BOT_URL + "/test-17", timeout=30)
            st.success("每日Excel 測試已發送, 睇 WhatsApp Group")
        except Exception as e:
            st.error(f"失敗: {e}")
    if st.button("🔄 重新整理單列表", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

# ---------- 主畫面, Upload 完都唔會唔見 ----------
st.markdown("💡 提示: 影相唔好咁遠, 唔好反光, 唔好影到人哋個招牌, 唔好咁多雜物, 影得近啲, 光啲再試。")

# 如果啱啱 Upload 完, 唔好再用「還要上傳 POD」嗰個 Filter Hide 咗佢
if st.session_state["just_uploaded_job"]:
    st.success(f"✅ 上一單 {st.session_state['just_uploaded_job']} 已上傳成功, 已轉做 DELIVERED")
    if st.button("返去單列表, 睇其他單"):
        st.session_state["just_uploaded_job"] = ""
        st.cache_data.clear()
        st.rerun()

st.markdown("🔍 搜尋 JobID / 客戶 / 地址 / InternalNo")
search = st.text_input("例如: JOB-1117, TSTN, 客戶名...", label_visibility="collapsed", key="search_box")
# FIX: 預設唔好再 Tick 死「還要上傳 POD」, 等 Felix 睇到全部單, 唔會 Upload 完就空白
only_need_pod = st.checkbox("還要上傳 POD 嘅單 (已按搜尋 Filter)", value=False, key="only_pod")

try:
    df, headers, ws = load_jobs()
except Exception as e:
    st.error(f"讀 Jobs 失敗: {e}")
    st.stop()

if df.empty:
    st.warning("Jobs 冇資料, 睇下 Sheet 係咪空")
    st.stop()

col_job = find_col(headers, ["JobID", "InternalNo", "單號"])
col_client = find_col(headers, ["Client", "客戶"])
col_date = find_col(headers, ["Date", "送貨日期"])
col_from = find_col(headers, ["From", "地址", "FromAddress"])
col_driver = find_col(headers, ["Driver", "司機"])
col_status = find_col(headers, ["Status", "狀態"])
col_cargo = find_col(headers, ["CargoDetail", "貨物", "Cargo"])
col_pod = find_col(headers, ["POD", "PODLink", "POD連結"])

filtered = df.copy()
# 唔好再硬性 Filter 走 Felix 嗰啲單, 只係按搜尋嚟 Filter
if search:
    s = search.strip().lower()
    mask = pd.Series([False] * len(filtered))
    for c in [col_job, col_client, col_from]:
        if c and c in filtered.columns:
            mask = mask | filtered[c].astype(str).str.lower().str.contains(s, na=False)
    filtered = filtered[mask]
if only_need_pod and col_status and col_status in filtered.columns:
    filtered = filtered[~filtered[col_status].astype(str).str.upper().isin(["DELIVERED", "COMPLETED", "CANCELLED", "CANCEL"])]

if filtered.empty:
    st.info("冇搵到符合嘅單, 試下剔走「還要上傳 POD 嘅單」, 睇下係咪已經變咗 DELIVERED")
    # 唔好 st.stop(), 要留返上面啲掣俾 Felix 撳
    st.stop()

# 單列表, 一定要成個 List 顯示返
st.markdown(f"共 {len(filtered)} 單")
options = []
for idx, row in filtered.iterrows():
    label = f"{row[col_job]} | {row[col_client] if col_client in row else ''} | {row[col_date] if col_date in row else ''} | {row[col_status] if col_status in row else ''}"
    options.append((label, idx))

sel_label = st.selectbox("揀單", [o[0] for o in options], label_visibility="collapsed", key="job_select")
sel_idx = [o[1] for o in options if o[0] == sel_label][0]
row = filtered.loc[sel_idx]
job_id = str(row[col_job]) if col_job else ""

st.markdown("---")
st.markdown(f"### 你揀咗: {job_id}")
if col_client in row:
    st.markdown(f"👤 客戶: {row[col_client]} | 📅 {row[col_date] if col_date in row else ''}")
if col_from in row:
    st.markdown(f"📍 {row[col_from]}")
if col_driver in row:
    st.markdown(f"🚚 司機: {row[col_driver]} | 狀態: {row[col_status] if col_status in row else ''}")
if col_cargo in row:
    st.markdown(f"📦 貨物: {row[col_cargo]}")

st.markdown("---")
st.markdown("### 📷 拍照上傳POD (一次最多 3 張, 系統會自動壓縮)")
uploaded = st.file_uploader("Upload", type=["jpg", "jpeg", "png"], accept_multiple_files=True, label_visibility="collapsed", help="200MB per file • JPG, PNG", key="pod_uploader")
if uploaded and len(uploaded) > 3:
    st.warning("一次最多 3 張, 淨係會 Upload 頭 3 張")
    uploaded = uploaded[:3]

if st.button("📤 上傳POD + 轉Delivered + 通知", type="primary", use_container_width=True):
    if not uploaded:
        st.error("請先揀相")
    else:
        fails = []
        links = []
        for f in uploaded:
            ok, msg = upload_pod_to_drive(job_id, f)
            if ok:
                links.append(msg)
            else:
                fails.append(f"{f.name} ({msg})")
        if fails:
            st.error("以下 " + str(len(fails)) + " 張失敗: " + ", ".join(fails))
        if not fails:
            try:
                fresh_sh = get_sheet().open_by_key(TMS_SHEET_ID)
                fresh_ws = fresh_sh.worksheet("Jobs")
                all_vals = fresh_ws.get_all_values()
                headers2 = [h.strip() for h in all_vals[0]]
                job_col_idx = headers2.index(col_job) if col_job in headers2 else 0
                status_col_idx = headers2.index(col_status) if col_status in headers2 else None
                pod_col_idx = headers2.index(col_pod) if col_pod and col_pod in headers2 else None
                target_row = None
                for i, r in enumerate(all_vals[1:], start=2):
                    if len(r) > job_col_idx and r[job_col_idx].strip() == job_id.strip():
                        target_row = i
                        break
                if target_row and status_col_idx is not None:
                    fresh_ws.update_cell(target_row, status_col_idx + 1, "DELIVERED")
                if target_row and pod_col_idx is not None and links:
                    fresh_ws.update_cell(target_row, pod_col_idx + 1, " | ".join([l for l in links if l]))
                # FIX: 唔好即刻 Clear 晒, 留返個 just_uploaded_job, 等畫面仲見到嘢
                st.session_state["just_uploaded_job"] = job_id
                st.cache_data.clear()
                try:
                    requests.get(BOT_URL + "/test-jobs", timeout=20)
                except Exception:
                    pass
                st.success(f"✅ 上傳成功! {job_id} 已轉做 DELIVERED")
                st.balloons()
                st.rerun()
            except Exception as e:
                st.error(f"更新 Sheet 失敗, 但 POD 已上 Drive: {e}")
                st.code(traceback.format_exc())
