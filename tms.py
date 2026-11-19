# tms.py - TMS V46.9 My Tasks Fix
# Streamlit TMS Apps: 司機睇單 + Upload POD
# 呢份係成份, 直接取代舊 tms.py, app.py 唔使郁

import streamlit as st
import os
import re
import io
import json
import base64
import traceback
from datetime import datetime, timedelta
import requests
import pandas as pd
from PIL import Image

st.set_page_config(page_title="TMS V46.9 My Tasks", layout="wide")

# ============ 設定 - 唔好改, 同 app.py 嗰邊對返 ============
TMS_SHEET_ID = os.getenv("TMS_SHEET_ID", "15X06a1oOFeta-ETz-CiLbB-5TUWS1rWKGi35HmBNNVE")
GCP_JSON = os.getenv("GCP_JSON", "")
BOT_URL = os.getenv("BOT_URL", "https://cookie-delivery-bot-production.up.railway.app")
DRIVE_FOLDER_ID = os.getenv("POD_DRIVE_FOLDER_ID", "")
# 如果你 Drive POD Folder 冇 ID, 留空都得, 會淨係更新 Sheet, 唔會 Fail

def get_creds():
    from google.oauth2.service_account import Credentials
    s = GCP_JSON.strip()
    if not s.startswith("{"):
        s = base64.b64decode(s).decode()
    info = json.loads(s)
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    return Credentials.from_service_account_info(info, scopes=scopes)

def get_sheet():
    import gspread
    return gspread.authorize(get_creds())

def get_drive():
    from googleapiclient.discovery import build
    return build("drive", "v3", credentials=get_creds())

# ============ POD 壓縮 - 呢度係今次 Fix 嘅重點 ============
def compress_pod(uploaded_file):
    # 一定要用壓縮完嘅 data, 唔好用原圖 size 嚟判
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
    # 返 (ok, link_or_error)
    try:
        if not DRIVE_FOLDER_ID:
            return True, ""
        from googleapiclient.http import MediaIoBaseUpload
        drive = get_drive()
        # 檔名唔好有括號、空格, 用 JOB + 時間
        safe_name = f"{job_id}_POD_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
        # FIX: 用 compress 完嗰個 data, 唔好用 uploaded_file.getvalue()
        data, final_size = compress_pod(uploaded_file)
        if final_size > 2 * 1024 * 1024:
            return False, f"壓縮完仲係 {final_size//1024}KB, 太大"
        media = MediaIoBaseUpload(io.BytesIO(data), mimetype="image/jpeg", resumable=True)
        body = {"name": safe_name, "parents": [DRIVE_FOLDER_ID], "mimeType": "image/jpeg"}
        f = drive.files().create(body=body, media_body=media, fields="id,webViewLink").execute()
        return True, f.get("webViewLink", "")
    except Exception as e:
        return False, f"Drive 失敗: {str(e)[:200]}"

# ============ 讀 Jobs ============
@st.cache_data(ttl=60)
def load_jobs():
    try:
        sh = get_sheet().open_by_key(TMS_SHEET_ID)
        ws = sh.worksheet("Jobs")
        vals = ws.get_all_values()
        if len(vals) < 2:
            return pd.DataFrame(), [], ws
        headers = [h.strip() for h in vals[0]]
        df = pd.DataFrame(vals[1:], columns=headers)
        return df, headers, ws
    except Exception as e:
        st.error(f"讀 Jobs 失敗: {e}")
        return pd.DataFrame(), [], None

def find_col(headers, names):
    for n in names:
        if n in headers:
            return n
    return None

# ============ 左邊欄 - 原本功能全部保留 ============
with st.sidebar:
    st.markdown("### 👤 展 (driver)")
    st.markdown("🔥 V46.9 My Tasks Fix")
    st.caption(f"上次自動發送: 2026-11-18")
    if st.button("登出"):
        st.session_state.clear()
        st.rerun()
    st.markdown("---")
    if st.button("📦 同步OrderCosts"):
        st.info("同步OrderCosts 已觸發, 請睇 Sheet 有冇更新")
        try:
            r = requests.get(BOT_URL + "/test-jobs", timeout=20)
            st.success(f"同步完成: {r.status_code}")
        except Exception as e:
            st.error(f"同步失敗: {e}")
    if st.button("🔍 測試WhatsApp Group"):
        try:
            r = requests.get(BOT_URL + "/", timeout=20)
            j = r.json()
            st.success(f"Bot 正常, cookie_photos={j.get('cookie_photos')}, pod={j.get('pod_photos', 0)}")
            st.json(j.get("last_webhook", {}))
        except Exception as e:
            st.error(f"Bot 唔通: {e}")
    if st.button("📊 測試每日Excel"):
        try:
            r = requests.get(BOT_URL + "/test-17", timeout=30)
            st.success("每日Excel 測試已發送, 睇 WhatsApp Group")
        except Exception as e:
            st.error(f"失敗: {e}")

# ============ 主畫面 ============
st.markdown("💡 提示: 影相唔好咁遠, 唔好反光, 唔好影到人哋個招牌, 唔好咁多雜物, 影得近啲, 光啲, 唔好咁多字, 唔好咁多嘢, 影得近啲, 光啲再試。")
st.markdown("🔍 搜尋 JobID / 客戶 / 地址 / InternalNo")
search = st.text_input("例如: JOB-1117, TSTN, 客戶名...", label_visibility="collapsed")
only_need_pod = st.checkbox("還要上傳 POD 嘅單 (已按搜尋 Filter)", value=True)

df, headers, ws = load_jobs()
if df.empty:
    st.warning("Jobs 冇資料, 睇下 Sheet 係咪空")
    st.stop()

# 欄名對應 - 唔好改, 要對返你 Sheet 嗰啲欄
col_job = find_col(headers, ["JobID", "InternalNo", "單號"])
col_client = find_col(headers, ["Client", "客戶"])
col_date = find_col(headers, ["Date", "送貨日期"])
col_from = find_col(headers, ["From", "地址", "FromAddress"])
col_driver = find_col(headers, ["Driver", "司機"])
col_status = find_col(headers, ["Status", "狀態"])
col_cargo = find_col(headers, ["CargoDetail", "貨物", "Cargo"])
col_pod = find_col(headers, ["POD", "PODLink", "POD連結"])

# Filter
filtered = df.copy()
if search:
    s = search.strip().lower()
    mask = pd.Series([False] * len(filtered))
    for c in [col_job, col_client, col_from]:
        if c and c in filtered.columns:
            mask = mask | filtered[c].astype(str).str.lower().str.contains(s, na=False)
    filtered = filtered[mask]
if only_need_pod and col_status and col_status in filtered.columns:
    # 未 Delivered 嘅先係要上傳 POD
    filtered = filtered[~filtered[col_status].astype(str).str.upper().isin(["DELIVERED", "COMPLETED", "CANCELLED", "CANCEL"])]

if filtered.empty:
    st.info("冇搵到符合嘅單")
    st.stop()

# 下拉揀單 - 格式同你 Cap 圖一樣
def job_label(row):
    parts = []
    if col_job in row:
        parts.append(str(row[col_job]))
    if col_client in row:
        parts.append(str(row[col_client]))
    if col_job in row:
        # 唔好重複, 呢度加 InternalNo 嗰類
        pass
    if col_date in row:
        parts.append(str(row[col_date]))
    return " | ".join([p for p in parts if p])

options = []
for idx, row in filtered.iterrows():
    label = job_label(row)
    # 加埋 InternalNo / 3RTSA 嗰啲, 等司機易搵
    if col_job in row:
        label = f"{row[col_job]} | {row[col_client] if col_client in row else ''} | {row[col_date] if col_date in row else ''}"
    options.append((label, idx))

sel_label = st.selectbox("揀單", [o[0] for o in options], label_visibility="collapsed")
sel_idx = [o[1] for o in options if o[0] == sel_label][0]
row = filtered.loc[sel_idx]
job_id = str(row[col_job]) if col_job else ""

# 顯示單嘅詳情
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

# ============ POD 上傳 - Fix 咗喺呢度 ============
st.markdown("---")
st.markdown("### 📷 拍照上傳POD (一次最多 3 張, 系統會自動壓縮)")
uploaded = st.file_uploader("Upload", type=["jpg", "jpeg", "png"], accept_multiple_files=True, label_visibility="collapsed", help="200MB per file • JPG, PNG")
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
            # 唔好再寫「請再影過, 企近啲、光啲再試」, 因為唔係影相問題
            st.error("以下 " + str(len(fails)) + " 張失敗: " + ", ".join(fails))
        if not fails:
            # 全部 Upload 成功, 先至更新 Sheet 做 Delivered
            try:
                # 搵返 Jobs 嗰頁, 嗰單係第幾行
                all_vals = ws.get_all_values()
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
                    ws.update_cell(target_row, status_col_idx + 1, "DELIVERED")
                if target_row and pod_col_idx is not None and links:
                    ws.update_cell(target_row, pod_col_idx + 1, " | ".join([l for l in links if l]))
                st.success(f"✅ 上傳成功! 已轉做 DELIVERED, Bot 會喺 5 分鐘內喺 運輸 (VIP單) Group 發通知")
                st.balloons()
                # 叫 Bot 即刻 Check 一次, 唔使等 5 分鐘
                try:
                    requests.get(BOT_URL + "/test-jobs", timeout=20)
                except Exception:
                    pass
            except Exception as e:
                st.error(f"更新 Sheet 失敗, 但 POD 已上 Drive: {e}")
                st.code(traceback.format_exc())
