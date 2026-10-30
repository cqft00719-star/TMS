import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime
from datetime import timezone, timedelta
import uuid, io, base64, requests, zipfile, os, time
from PIL import Image, ImageDraw, ImageFont, ImageOps

HKT = timezone(timedelta(hours=8))
def now_hk(): return datetime.datetime.now(HKT)
def today_hk(): return now_hk().date()

try:
    from streamlit_autorefresh import st_autorefresh
    st_autorefresh(interval=60*1000, key="tms_autorefresh_60s")
except:
    pass

st.set_page_config(page_title="TMS V46.3 修復版", layout="wide", page_icon="🚚")

NEW_FEES = ["隧道費","登記費","停車費","交倉附加費","其他"]
ULTRAMSG_INSTANCE = "instance194419"
ULTRAMSG_TOKEN = "ji5hd4t2li7vf3ys"
ULTRAMSG_GROUP_ID = "120363413020795341@g.us"
DAILY_AUTO_HOUR = 17
DAILY_AUTO_MINUTE = 0

def send_ultramsg_message(to, body):
    try:
        url = f"https://api.ultramsg.com/{ULTRAMSG_INSTANCE}/messages/chat"
        payload = {"token": ULTRAMSG_TOKEN, "to": to, "body": body}
        r = requests.post(url, data=payload, timeout=20)
        return r.status_code == 200, r.text[:500]
    except Exception as e:
        return False, str(e)

def send_ultramsg_document(to, file_path, caption=""):
    try:
        with open(file_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("utf-8")
        url = f"https://api.ultramsg.com/{ULTRAMSG_INSTANCE}/messages/document"
        payload = {"token": ULTRAMSG_TOKEN, "to": to, "document": b64, "filename": os.path.basename(file_path), "caption": caption}
        r = requests.post(url, data=payload, timeout=60)
        return r.status_code == 200, r.text[:1000]
    except Exception as e:
        return False, str(e)

def send_whatsapp_to_ops(message):
    return send_ultramsg_message(ULTRAMSG_GROUP_ID, message)[0], ""

def parse_company_export_template(file):
    df = pd.read_excel(file, sheet_name=0, header=0, dtype=str).fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    if "要貨單號" in df.columns:
        df = df[~df["要貨單號"].astype(str).str.contains("合計|總計", na=False)]
        df = df[df["要貨單號"].astype(str).str.strip()!=""]
    grouped = []
    if "要貨單號" not in df.columns:
        return df
    for order_no, group in df.groupby("要貨單號", sort=False):
        first = group.iloc[0].to_dict()
        cargo_lines = [str(r.get("商品編碼","")).strip() for _, r in group.iterrows() if str(r.get("商品編碼","")).strip()]
        merged = {
            "商戶類型": first.get("商戶類型",""), "商戶名稱": first.get("商戶名稱",""),
            "發貨倉庫": first.get("發貨倉庫",""), "發貨人": first.get("發貨人",""),
            "發貨聯絡電話": first.get("發貨聯絡電話",""), "接貨時間": first.get("接貨時間",""),
            "收貨地址": first.get("收貨地址",""), "收貨人": first.get("收貨人",""),
            "聯係電話": first.get("聯係電話",""), "送貨日期": first.get("送貨日期",""),
            "要貨單號": order_no, "創建日期": first.get("創建日期",""), "整單備注": first.get("整單備注",""),
            "CargoDetail": "\n".join(cargo_lines), "商品明細數": len(group),
            "InternalNo": order_no, "Client": str(first.get("商戶名稱","")).strip(),
            "From": str(first.get("發貨倉庫","")).strip(), "To": str(first.get("收貨地址","")).strip(),
            "Date": str(first.get("送貨日期","")).split(" ")[0] if first.get("送貨日期") else str(today_hk()),
        }
        grouped.append(merged)
    return pd.DataFrame(grouped)

def generate_and_send_daily_excel():
    df_jobs = load_df("Jobs")
    if df_jobs.empty:
        send_ultramsg_message(ULTRAMSG_GROUP_ID, f"📦 TMS每日提醒 {now_hk().strftime('%Y-%m-%d')} 未來兩日無Order")
        return False, "無Jobs"
    date_col = None
    for c in ["Date","送貨日期","DeliveryDate"]:
        if c in df_jobs.columns:
            date_col = c
            break
    if not date_col: return False, "無Date欄"
    df_jobs[date_col] = pd.to_datetime(df_jobs[date_col], errors="coerce")
    tomorrow = today_hk() + datetime.timedelta(days=1)
    day_after = tomorrow + datetime.timedelta(days=1)
    filtered = df_jobs[df_jobs[date_col].dt.date.isin([tomorrow, day_after])]
    if filtered.empty:
        send_ultramsg_message(ULTRAMSG_GROUP_ID, f"📦 TMS每日提醒 {now_hk().strftime('%Y-%m-%d')} 聽日 {tomorrow} 同後日 {day_after} 無Order")
        return True, f"無單已發文字 {tomorrow}~{day_after}"
    out_path = f"/tmp/TMS_Orders_{tomorrow}_to_{day_after}.xlsx"
    export_cols = [c for c in ["InternalNo","JobID","Date","Client","Zone","From","To","CargoDetail","DeliveryTime","Notes"] if c in filtered.columns]
    filtered[export_cols].to_excel(out_path, index=False)
    caption = f"📦 TMS每日提醒 {now_hk().strftime('%Y-%m-%d %H:%M HKT')}\n聽日 {tomorrow} 同後日 {day_after} 共有 {len(filtered)} 張單"
    ok, info = send_ultramsg_document(ULTRAMSG_GROUP_ID, out_path, caption)
    return ok, info

def get_last_daily_log():
    try:
        df = load_df("DailyLog")
        if df.empty or "Date" not in df.columns: return None
        df = df[df["Date"].astype(str).str.strip()!=""]
        if df.empty: return None
        return str(df.iloc[-1]["Date"]).strip()
    except: return None

def write_daily_log(status):
    try:
        sh = get_sh()
        try: ws = sh.worksheet("DailyLog")
        except:
            ws = sh.add_worksheet(title="DailyLog", rows=1000, cols=10)
            ws.update('A1', [["Date","Time_HKT","Status","Detail","TriggeredBy"]])
        ws.append_row([str(today_hk()), now_hk().strftime("%H:%M:%S"), status, str(now_hk()), "auto"], value_input_option='USER_ENTERED')
        st.cache_data.clear()
    except: pass

def check_and_auto_send_daily():
    now = now_hk()
    if now.hour < DAILY_AUTO_HOUR or (now.hour == DAILY_AUTO_HOUR and now.minute < DAILY_AUTO_MINUTE):
        return
    last = get_last_daily_log()
    if last == str(today_hk()):
        return
    ok, info = generate_and_send_daily_excel()
    write_daily_log(f"{'OK' if ok else 'FAIL'} {info[:120]}")

def set_job_delivered(jid):
    try:
        ws = get_sh().worksheet("Jobs")
        headers = [h.strip() for h in ws.row_values(1)]
        if "JobID" not in headers: return False
        jid_col = headers.index("JobID") + 1
        status_idx = [i for i,h in enumerate(headers) if h=="Status"]
        if not status_idx: return False
        status_col = status_idx[-1] + 1
        pod_col = headers.index("PodTime_HKT")+1 if "PodTime_HKT" in headers else None
        all_jids = ws.col_values(jid_col)
        for r_idx, v in enumerate(all_jids[1:], start=2):
            if str(v).strip() == str(jid).strip():
                ws.update_cell(r_idx, status_col, "Delivered")
                if pod_col:
                    ws.update_cell(r_idx, pod_col, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"))
                st.cache_data.clear()
                return True
        return False
    except: return False

def compress_pod_high_quality(uploaded_file):
    try:
        img = Image.open(io.BytesIO(uploaded_file.getvalue()))
        img = ImageOps.exif_transpose(img).convert("RGB")
        for max_side in [1600, 1400, 1200, 1024, 900]:
            for quality in [85, 75, 65, 55]:
                buf = io.BytesIO()
                tmp = img.copy()
                tmp.thumbnail((max_side, max_side), Image.LANCZOS)
                tmp.save(buf, format="JPEG", quality=quality, optimize=True)
                fb = buf.getvalue()
                if len(fb) <= 48000:
                    return fb, max_side, quality
        buf = io.BytesIO()
        img.thumbnail((800, 800), Image.LANCZOS)
        img.save(buf, format="JPEG", quality=50, optimize=True)
        return buf.getvalue(), 800, 50
    except: return None, 0, 0

def watermark_pod_image(b64_data_url, job_id, internal_no=""):
    try:
        if "," in str(b64_data_url): b64_data_url = str(b64_data_url).split(",",1)[1]
        img_bytes = base64.b64decode(b64_data_url)
        img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
        draw = ImageDraw.Draw(img, "RGBA")
        w, h = img.size; bar_h = 90
        draw.rectangle([(0, h-bar_h), (w, h)], fill=(0,0,0,180))
        try: font = ImageFont.truetype("DejaVuSans.ttf", 22)
        except: font = ImageFont.load_default()
        draw.text((10, h-bar_h+5), f"Job:{job_id}", fill=(255,255,255), font=font)
        draw.text((10, h-bar_h+30), f"Internal:{internal_no}" if internal_no else "Internal:N/A", fill=(255,220,0), font=font)
        draw.text((10, h-bar_h+55), now_hk().strftime("%Y/%m/%d %H:%M HKT"), fill=(200,200,200), font=font)
        buf = io.BytesIO(); img.save(buf, format="JPEG", quality=90)
        return buf.getvalue(), img
    except: return None, None

@st.cache_resource
def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    return gspread.authorize(creds)
def get_sh(): return get_client().open_by_key(st.secrets["connections"]["gsheets"]["spreadsheet"])
@st.cache_data(ttl=30)
def load_df(tab):
    try:
        ws=get_sh().worksheet(tab); vals=ws.get_all_values()
        if len(vals)<1: return pd.DataFrame()
        headers=[h.strip() for h in vals[0]]; clean=[]
        for i,h in enumerate(headers):
            if h=="": h=f"COL_{i}"
            if h in clean: h=f"{h}_{i}"
            clean.append(h)
        rows=[]
        for r in vals[1:]:
            if any(str(c).strip()!="" for c in r):
                if len(r)<len(clean): r=r+[""]*(len(clean)-len(r))
                rows.append(r[:len(clean)])
        return pd.DataFrame(rows, columns=clean)
    except: return pd.DataFrame()

def append_row_safe(tab,row):
    try:
        safe_row=[str(c)[:45000] for c in row]
        get_sh().worksheet(tab).append_rows([safe_row], value_input_option='USER_ENTERED')
        st.cache_data.clear()
        return True
    except Exception as e:
        if "429" in str(e):
            time.sleep(2.5)
            try:
                get_sh().worksheet(tab).append_rows([safe_row], value_input_option='USER_ENTERED')
                st.cache_data.clear()
                return True
            except: return False
        st.error(f"寫入 {tab} 失敗: {e}")
        return False

def append_row(tab,row): return append_row_safe(tab,row)
def append_row_by_header(tab, data_dict):
    ws=get_sh().worksheet(tab); headers=[h.strip() for h in ws.row_values(1)]; row=[""]*len(headers)
    for k,v in data_dict.items():
        if k in headers: row[headers.index(k)]=str(v)
    return append_row_safe(tab,row)

def sync_order_costs_sheet():
    try:
        df_jobs=load_df("Jobs"); df_exp=load_df("Expenses")
        if df_jobs.empty: return 0
        df_cost=df_jobs.copy(); df_cost["TotalExpense"]=0
        for fee in NEW_FEES: df_cost[fee]=0
        if not df_exp.empty and "JobID" in df_exp.columns:
            df_exp["Amount_num"]=pd.to_numeric(df_exp["Amount"], errors="coerce").fillna(0)
            total_map=df_exp.groupby("JobID")["Amount_num"].sum().to_dict()
            df_cost["TotalExpense"]=df_cost["JobID"].map(total_map).fillna(0)
            if "Type" in df_exp.columns:
                pivot=df_exp.pivot_table(index="JobID", columns="Type", values="Amount_num", aggfunc="sum", fill_value=0)
                for fee in NEW_FEES:
                    if fee in pivot.columns: df_cost[fee]=df_cost["JobID"].map(pivot[fee].to_dict()).fillna(0)
        headers=["InternalNo","JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","DeliveryTime","CargoDetail","Status","Driver"]+NEW_FEES+["TotalExpense","LastSync_HKT"]
        df_cost["LastSync_HKT"]=now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")
        for h in headers:
            if h not in df_cost.columns: df_cost[h]=""
        ws=get_sh().worksheet("OrderCosts"); ws.clear(); ws.update('A1',[headers])
        if not df_cost.empty:
            rows=[[str(r.get(h,"")) for h in headers] for _,r in df_cost.sort_values(by="JobID",ascending=False).iterrows()]
            for i in range(0,len(rows),100):
                ws.append_rows(rows[i:i+100], value_input_option='USER_ENTERED')
                time.sleep(0.5)
        st.cache_data.clear(); return len(df_cost)
    except Exception as e: st.error(f"OrderCosts同步失敗: {e}"); return 0

def ensure_sheets():
    sh=get_sh(); existing=[ws.title for ws in sh.worksheets()]
    def add(t,h,d=None):
        if t not in existing:
            ws=sh.add_worksheet(title=t, rows=3000, cols=40); ws.update('A1',[h])
            if d: ws.append_rows(d, value_input_option='USER_ENTERED')
    add("Users",["username","password","name","role"],[["admin","admin123","Admin","admin"],["阿傑","123","阿傑","driver"],["John","John123","John","driver"],["Mo","Mo123","Mo","driver"],["op01","op01","OP01","op"],["Kit","Kit123","Kit","driver"]])
    add("Jobs",["InternalNo","JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","DeliveryTime","Driver","Status","CargoDetail","InternalNo","Notes","CreatedBy","CreatedAt_HKT","PodTime_HKT"])
    add("PODs",["POD_ID","JobID","FileName","Size","Base64_Image","CreatedBy","CreatedAt_HKT","PodTime_HKT"])
    add("Expenses",["ExpenseID","Date","JobID","Driver","Type","Amount","Remark","CreatedAt_HKT"])
    add("OrderCosts",["InternalNo","JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","DeliveryTime","CargoDetail","Status","Driver","隧道費","登記費","停車費","交倉附加費","其他","TotalExpense","LastSync_HKT"])
    add("DailyLog",["Date","Time_HKT","Status","Detail","TriggeredBy"])

def printable_order_html(job):
    jid = job.get('JobID',''); internal = job.get('InternalNo',''); client = job.get('Client',''); zone = job.get('Zone','')
    from_addr = job.get('From',''); from_contact = job.get('From_Contact',''); from_phone = job.get('From_Phone',''); pickup = job.get('PickupTime','')
    to_addr = job.get('To',''); to_contact = job.get('To_Contact',''); to_phone = job.get('To_Phone',''); delivery = job.get('DeliveryTime','')
    cargo = job.get('CargoDetail',''); notes = job.get('Notes','') or ""; date_str = job.get('Date','')
    expected = f"{date_str} {pickup}-{delivery}".strip() or delivery or pickup or date_str
    cargo_html = str(cargo).replace('\n','<br>') if cargo else ""
    return f"""<html><head><meta charset="utf-8"><title>Printable - {jid}</title>
    <style>body{{font-family: "Microsoft JhengHei", Arial; font-size:13px;}}.page{{border:2px solid #000; max-width:830px; margin:10px auto; padding:18px 22px;}} h1{{text-align:center; font-size:22px;}} table{{width:100%; border-collapse:collapse; margin:6px 0;}} td{{border:1px solid #000; padding:6px 8px;}}.label{{background:#f2f2f2; width:110px; font-weight:bold;}}</style></head><body>
    <div class="page"><h1>🚚 運輸訂單 {jid}</h1><div>Internal: {internal} | Client: {client} | Zone: {zone}</div>
    <div>From: {from_addr} {from_contact} {from_phone} Pickup: {pickup}</div>
    <div>To: {to_addr} {to_contact} {to_phone} Delivery: {delivery}</div>
    <div>Cargo: {cargo_html}</div><div>Notes: {notes}</div><div>Date: {date_str}</div></div>
    <div style="text-align:center;"><button onclick="window.print()">🖨️ 打印</button></div></body></html>"""

def format_job_option(row):
    return f"{row.get('JobID','')} | {str(row.get('Client',''))[:15]} | {row.get('InternalNo','')} | {str(row.get('Date',''))[:10]} {row.get('DeliveryTime','')}".strip()

if "user" not in st.session_state: st.session_state.user=None
if not st.session_state.user:
    st.title("TMS V46.3 登入")
    df_u=load_df("Users"); u=st.text_input("用戶名"); p=st.text_input("密碼",type="password")
    if st.button("登入",type="primary",key="v463_login"):
        m=df_u[(df_u["username"].astype(str).str.strip()==u.strip()) & (df_u["password"].astype(str).str.strip()==p.strip())]
        if not m.empty: st.session_state.user=m.iloc[0].to_dict(); st.rerun()
        else: st.error("帳號錯")
    st.stop()

ensure_sheets()
try: check_and_auto_send_daily()
except: pass

user=st.session_state.user; role=str(user.get("role","driver")).lower(); operator=user.get("username",""); op_name=user.get("name","")
st.sidebar.title(f"👤 {op_name} ({role})")
st.sidebar.write(f"🔥 當前版本: **V46.3 修復下載POD**")
last_log = get_last_daily_log()
st.sidebar.caption(f"上次自動發送: {last_log}" if last_log else "今日未自動發送 (17:00)")

if st.sidebar.button("登出",key="v463_logout"): st.session_state.user=None; st.cache_data.clear(); st.rerun()
if st.sidebar.button("🔄 同步OrderCosts",key="v463_sync"): st.sidebar.success(f"已同步 {sync_order_costs_sheet()} 單")
if st.sidebar.button("📅 查看DailyLog",key="v463_log"): st.sidebar.dataframe(load_df("DailyLog").tail(10))

if role=="admin": tabs=st.tabs(["📊 Dashboard","📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費","💵 每單費用","📥 下載POD"])
elif role=="op": tabs=st.tabs(["📊 Dashboard","📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費","💵 每單費用","📥 下載POD"])
else: tabs=st.tabs(["🟢 可接任務池","🚚 我的任務","📸 上傳POD","💰 雜費","📋 重複打印"])

df_jobs=load_df("Jobs")
def get_my_jobs(df):
    if df.empty: return df
    mask = df["Driver"].astype(str).str.strip().str.lower().isin([operator.lower(), op_name.lower()])
    mask2 = df["Driver"].astype(str).str.contains(operator, na=False) | df["Driver"].astype(str).str.contains(op_name, na=False)
    return df[mask | mask2]

with tabs[0]:
    if role=="driver":
        st.subheader("🟢 可接任務池")
        df_pool=df_jobs[(df_jobs["Driver"].astype(str).str.strip()=="") & (df_jobs["Status"].astype(str).str.upper().isin(["PENDING",""]))] if not df_jobs.empty else pd.DataFrame()
        if df_pool.empty: st.info("暫無可接")
        else:
            for idx,row in df_pool.iterrows():
                jid=row.get("JobID","")
                with st.container(border=True):
                    st.markdown(f"**{jid}** | {row.get('Client','')} | {row.get('Date','')}")
                    if st.button("🟢 搶此單",key=f"v463_claim_{jid}_{idx}",type="primary"):
                        ws=get_sh().worksheet("Jobs"); headers=[h.strip() for h in ws.row_values(1)]; cell=ws.find(jid)
                        if cell:
                            ws.update_cell(cell.row, headers.index("Driver")+1, operator)
                            ws.update_cell(cell.row, headers.index("Status")+1, "Accepted")
                            st.cache_data.clear(); st.success(f"已搶 {jid}"); st.rerun()
    else:
        st.subheader("📊 Dashboard V46.3")
        if not df_jobs.empty: st.dataframe(df_jobs.head(200), use_container_width=True)

if role in ["admin","op"]:
    with tabs[1]:
        st.subheader("📋 任務管理")
        if not df_jobs.empty:
            for idx,row in df_jobs.sort_values(by="JobID",ascending=False).head(100).iterrows():
                jid=row.get("JobID","")
                with st.container(border=True):
                    st.markdown(f"**{jid}** | {row.get('Client','')} | {row.get('Status','')}")
    with tabs[2]:
        st.subheader("📤 上傳Order")
        with st.form("v463_order_form"):
            order_date = st.date_input("📅 訂單日期", value=today_hk())
            client=st.text_input("客戶*"); from_loc=st.text_input("起點*"); to_loc=st.text_input("終點*")
            cargo=st.text_area("貨物詳情*"); internal_no=st.text_input("InternalNo")
            if st.form_submit_button("✅ 建立Order"):
                if not client or not from_loc or not to_loc or not cargo: st.error("必填")
                else:
                    jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                    append_row_by_header("Jobs", {"InternalNo":internal_no, "JobID":jid, "Date":str(order_date), "Client":client, "From":from_loc, "To":to_loc, "CargoDetail":cargo, "Status":"Pending", "Driver":"", "CreatedBy":operator, "CreatedAt_HKT":now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")})
                    st.success(f"已建立 {jid}")

pod_idx=3 if role in ["admin","op"] else 2
with tabs[pod_idx]:
    st.subheader("📸 上傳POD")
    my_jobs = get_my_jobs(df_jobs) if role=="driver" else df_jobs
    my_pool=my_jobs.sort_values(by="JobID",ascending=False) if not my_jobs.empty else df_jobs
    if not my_pool.empty:
        labels = [format_job_option(r) for _,r in my_pool.iterrows()]
        label_to_jid = dict(zip(labels, my_pool["JobID"].tolist()))
        sel_label = st.selectbox("選擇JobID", labels, key="v463_pod_jid")
        jid = label_to_jid.get(sel_label)
        with st.form("v463_pod_form", clear_on_submit=True):
            pod_files=st.file_uploader("📷 上傳POD",type=["jpg","png","jpeg"],accept_multiple_files=True)
            if st.form_submit_button("📤 上傳POD", type="primary"):
                count_ok=0
                for pf in pod_files:
                    fb, _, _ = compress_pod_high_quality(pf)
                    if not fb: continue
                    b64=base64.b64encode(fb).decode()
                    data_url=f"data:image/jpeg;base64,{b64}"
                    pod_id=f"POD-{now_hk().strftime('%m%d%H%M%S')}-{uuid.uuid4().hex[:4]}"
                    if append_row_safe("PODs",[pod_id,jid,f"{jid}_POD.jpg",str(len(fb)),data_url,operator,now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"),now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")]): count_ok+=1
                set_job_delivered(jid)
                st.success(f"已上傳 {count_ok} 張")

expense_idx = 4 if role in ["admin","op"] else 3
with tabs[expense_idx]:
    st.subheader("💰 雜費")
    st.info("雜費功能同之前一樣")

cost_idx = 5 if role in ["admin","op"] else None
if cost_idx is not None:
    with tabs[cost_idx]:
        st.subheader("💵 每單費用")
        if st.button("🔄 同步", key="v463_cost_sync"): sync_order_costs_sheet(); st.rerun()
        df_cost=load_df("OrderCosts")
        if not df_cost.empty: st.dataframe(df_cost, use_container_width=True)
    with tabs[6]:
        st.subheader("📥 下載POD - V46.3 修復版")
        df_pods = load_df("PODs")
        df_jobs_for_pod = load_df("Jobs")
        if df_pods.empty:
            st.info("暫無POD記錄")
        else:
            df_pods["JobID"] = df_pods["JobID"].astype(str).str.strip()
            if not df_jobs_for_pod.empty and "JobID" in df_jobs_for_pod.columns:
                df_jobs_for_pod["JobID"] = df_jobs_for_pod["JobID"].astype(str).str.strip()
                df_merged = pd.merge(df_pods, df_jobs_for_pod[["JobID","Date","Client","Driver","To","InternalNo"]], on="JobID", how="left")
            else:
                df_merged = df_pods.copy()
                df_merged["Date"] = ""; df_merged["Client"] = ""

            # V46.3修復重點：安全轉日期
            df_merged["CreatedAt_HKT"] = pd.to_datetime(df_merged["CreatedAt_HKT"], errors="coerce")
            df_merged = df_merged[df_merged["CreatedAt_HKT"].notna()].copy()
            
            st.markdown("#### 🔍 篩選時段")
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                start_date = st.date_input("開始日期", value=today_hk() - datetime.timedelta(days=7), key="v463_pod_start")
            with c2:
                end_date = st.date_input("結束日期", value=today_hk(), key="v463_pod_end")
            with c3:
                driver_list = ["全部"] + sorted([x for x in df_merged["Driver"].dropna().astype(str).unique().tolist() if x.strip()]) if "Driver" in df_merged.columns else ["全部"]
                sel_driver = st.selectbox("司機", driver_list, key="v463_pod_driver")
            with c4:
                client_list = ["全部"] + sorted([x for x in df_merged["Client"].dropna().astype(str).unique().tolist() if x.strip()]) if "Client" in df_merged.columns else ["全部"]
                sel_client = st.selectbox("客戶", client_list, key="v463_pod_client")

            # 安全篩選 - 修復InvalidComparison
            filtered = df_merged.copy()
            filtered["CreatedDate"] = filtered["CreatedAt_HKT"].dt.date
            filtered = filtered[(filtered["CreatedDate"] >= start_date) & (filtered["CreatedDate"] <= end_date)]
            if sel_driver != "全部":
                filtered = filtered[filtered["Driver"].astype(str).str.contains(sel_driver, na=False)]
            if sel_client != "全部":
                filtered = filtered[filtered["Client"].astype(str).str.contains(sel_client, na=False)]

            st.success(f"找到 {len(filtered)} 張POD ({start_date} 至 {end_date})")

            if not filtered.empty:
                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    if st.button(f"📦 打包下載全部 {len(filtered)} 張 (ZIP)", type="primary", use_container_width=True, key="v463_dl_zip"):
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                            for _, r in filtered.iterrows():
                                try:
                                    b64_str = str(r.get("Base64_Image",""))
                                    if "," in b64_str: b64_str = b64_str.split(",",1)[1]
                                    img_bytes = base64.b64decode(b64_str)
                                    wm_bytes, _ = watermark_pod_image(b64_str, r.get("JobID",""), r.get("InternalNo",""))
                                    final_bytes = wm_bytes if wm_bytes else img_bytes
                                    fname = f"{r.get('JobID','')}_{r.get('FileName','POD.jpg')}"
                                    fname = "".join(c for c in fname if c.isalnum() or c in "._- ").strip()
                                    zipf.writestr(fname, final_bytes)
                                except: pass
                        zip_buffer.seek(0)
                        st.download_button("⬇️ 點擊下載ZIP", data=zip_buffer, file_name=f"POD_{start_date}_to_{end_date}.zip", mime="application/zip", key="v463_dl_btn_zip", use_container_width=True)
                with col_dl2:
                    csv_buf = filtered[["JobID","Client","Date","To","FileName","CreatedAt_HKT"]].to_csv(index=False).encode('utf-8-sig')
                    st.download_button("📄 下載清單CSV", data=csv_buf, file_name=f"POD_List_{start_date}_to_{end_date}.csv", mime="text/csv", use_container_width=True, key="v463_dl_csv")

                st.divider()
                for idx, row in filtered.sort_values(by="CreatedAt_HKT", ascending=False).head(100).iterrows():
                    with st.container(border=True):
                        c_img, c_info = st.columns([1,2])
                        with c_img:
                            try:
                                b64_str = str(row.get("Base64_Image",""))
                                if "," in b64_str: b64_str = b64_str.split(",",1)[1]
                                st.image(base64.b64decode(b64_str), use_container_width=True)
                            except: st.write("圖片失效")
                        with c_info:
                            st.markdown(f"**{row.get('JobID','')}** | {row.get('Client','')} | {row.get('To','')}")
                            st.caption(f"📅 {row.get('CreatedAt_HKT','')} | 司機: {row.get('Driver','')}")
                            try:
                                b64_str = str(row.get("Base64_Image",""))
                                if "," in b64_str: b64_str = b64_str.split(",",1)[1]
                                img_bytes = base64.b64decode(b64_str)
                                wm_bytes, _ = watermark_pod_image(b64_str, row.get("JobID",""), row.get("InternalNo",""))
                                dl_bytes = wm_bytes if wm_bytes else img_bytes
                                st.download_button(f"⬇️ 下載", data=dl_bytes, file_name=row.get('FileName','POD.jpg'), mime="image/jpeg", key=f"v463_dl_one_{idx}")
                            except: pass
