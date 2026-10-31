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

st.set_page_config(page_title="TMS V46.7 Pro Dashboard", layout="wide", page_icon="🚚")

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
    try:
        ok, info = send_ultramsg_message(ULTRAMSG_GROUP_ID, message)
        return ok, info
    except Exception as e:
        return False, str(e)

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
    except Exception as e:
        print(f"DailyLog write fail {e}")

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
    except Exception as e:
        print(f"set_job_delivered fail {e}")
        return False

def compress_pod_high_quality(uploaded_file):
    try:
        img = Image.open(io.BytesIO(uploaded_file.getvalue()))
        img = ImageOps.exif_transpose(img).convert("RGB")
        for max_side in [1200, 1024, 900, 800, 700, 600, 500, 400]:
            for quality in [70, 60, 50, 40, 35, 30]:
                buf = io.BytesIO()
                tmp = img.copy()
                tmp.thumbnail((max_side, max_side), Image.LANCZOS)
                tmp.save(buf, format="JPEG", quality=quality, optimize=True)
                fb = buf.getvalue()
                b64_len = len(base64.b64encode(fb))
                if b64_len <= 48000:
                    return fb, max_side, quality
        buf = io.BytesIO()
        img.thumbnail((400, 400), Image.LANCZOS)
        img.save(buf, format="JPEG", quality=30, optimize=True)
        return buf.getvalue(), 400, 30
    except Exception as e:
        print(f"compress fail {e}")
        return None, 0, 0

def watermark_pod_image(b64_str, job_id, internal_no=""):
    try:
        s = str(b64_str)
        if "," in s: s = s.split(",",1)[1]
        s = s.strip()
        s += "=" * (-len(s) % 4)
        img_bytes = base64.b64decode(s)
        img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
        draw = ImageDraw.Draw(img, "RGBA")
        w, h = img.size; bar_h = 90
        draw.rectangle([(0, h-bar_h), (w, h)], fill=(0,0,0,180))
        try:
            font = ImageFont.truetype("DejaVuSans.ttf", 22)
        except:
            font = ImageFont.load_default()
        draw.text((10, h-bar_h+5), f"Job:{job_id}", fill=(255,255,255), font=font)
        draw.text((10, h-bar_h+30), f"Internal:{internal_no}" if internal_no else "Internal:N/A", fill=(255,220,0), font=font)
        draw.text((10, h-bar_h+55), now_hk().strftime("%Y/%m/%d %H:%M HKT"), fill=(200,200,200), font=font)
        buf = io.BytesIO(); img.save(buf, format="JPEG", quality=85)
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
        safe_row=[]
        for c in row:
            s=str(c)
            if tab=="PODs":
                if len(s) > 1000:
                    if s.startswith("data:image"):
                        if "," in s: s = s.split(",",1)[1]
                    if len(s) > 49000:
                        s = s[:49000]
            else:
                if len(s)>45000:
                    s=s[:45000]+"...[已截斷]"
            safe_row.append(s)
        get_sh().worksheet(tab).append_rows([safe_row], value_input_option='USER_ENTERED')
        st.cache_data.clear()
        return True
    except Exception as e:
        msg=str(e)
        if "429" in msg or "Quota" in msg:
            time.sleep(2.5)
            try:
                get_sh().worksheet(tab).append_rows([safe_row], value_input_option='USER_ENTERED')
                st.cache_data.clear()
                return True
            except Exception as e2:
                st.error(f"寫入 {tab} 失敗(重試後): {e2}")
                return False
        st.error(f"寫入 {tab} 失敗: {e}")
        return False

def append_row(tab,row):
    return append_row_safe(tab,row)

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
    <style>@media print {{.no-print {{ display:none; }} @page {{ margin:10mm; }} }} body{{font-family: "Microsoft JhengHei", Arial, sans-serif; font-size:13px;}}.page{{border:2px solid #000; max-width:830px; margin:10px auto; padding:18px 22px;}} h1{{text-align:center; font-size:22px; margin:5px 0 8px 0;}}.subtitle{{text-align:center; font-size:12px; margin-bottom:12px; border-bottom:1px solid #999; padding-bottom:8px;}} table{{width:100%; border-collapse:collapse; margin:6px 0;}} td{{border:1px solid #000; padding:6px 8px;}}.label{{background:#f2f2f2; width:110px; font-weight:bold;}}.label2{{background:#f2f2f2; width:55px; font-weight:bold; text-align:center;}}.section{{font-size:15px; font-weight:bold; margin:14px 0 4px 0;}}.sign{{margin-top:18px; display:flex; justify-content:space-between; font-size:12px;}}.hint{{text-align:center; font-size:11px; color:#666; margin-top:10px;}}</style></head><body>
    <div class="page"><h1>🚚 運輸訂單 Transport Order</h1><div class="subtitle">JobID: {jid} | Internal: {internal} | 預計: {expected}</div>
    <table><tr><td class="label">客戶 Client</td><td>{client}</td><td class="label2">Zone</td><td style="width:130px;">{zone}</td></tr></table>
    <div class="section">📦 起點 Pickup</div><table><tr><td class="label">地址 From</td><td colspan="3">{from_addr}</td></tr><tr><td class="label">聯絡人</td><td>{from_contact}</td><td class="label2">電話</td><td style="width:150px;">{from_phone}</td></tr><tr><td class="label">提貨時間</td><td colspan="3">{pickup}</td></tr></table>
    <div class="section">🏁 終點 Delivery</div><table><tr><td class="label">地址 To</td><td colspan="3">{to_addr}</td></tr><tr><td class="label">聯絡人</td><td>{to_contact}</td><td class="label2">電話</td><td style="width:150px;">{to_phone}</td></tr><tr><td class="label">送貨時間</td><td colspan="3">{delivery}</td></tr></table>
    <div class="section">📋 貨物資料 Cargo</div><table><tr><td class="label">詳細明細</td><td>{cargo_html}</td></tr><tr><td class="label">備註</td><td>{notes}</td></tr></table>
    <div class="sign"><div>司機簽收: __________________</div><div>客戶簽收: __________________</div><div>日期: {date_str}</div></div></div>
    <div class="hint no-print"><button onclick="window.print()" style="margin-top:10px; padding:10px 22px; font-size:15px;">🖨️ 打印 / Save as PDF</button></div></body></html>"""

def format_job_option(row):
    jid=str(row.get("JobID",""))
    client=str(row.get("Client",""))[:15]
    internal=str(row.get("InternalNo",""))
    date=str(row.get("Date",""))[:10]
    delivery=str(row.get("DeliveryTime",""))
    return f"{jid} | {client} | {internal} | {date} {delivery}".strip()

# ================= 登入 =================
if "user" not in st.session_state: st.session_state.user=None
if not st.session_state.user:
    st.title("TMS V46.7 登入 - Pro Dashboard")
    df_u=load_df("Users"); u=st.text_input("用戶名"); p=st.text_input("密碼",type="password")
    if st.button("登入",type="primary",key="v467_login"):
        m=df_u[(df_u["username"].astype(str).str.strip()==u.strip()) & (df_u["password"].astype(str).str.strip()==p.strip())]
        if not m.empty: st.session_state.user=m.iloc[0].to_dict(); st.rerun()
        else: st.error("帳號錯")
    st.stop()

ensure_sheets()
try:
    check_and_auto_send_daily()
except Exception as e:
    print(f"auto daily fail {e}")

user=st.session_state.user; role=str(user.get("role","driver")).lower(); operator=user.get("username",""); op_name=user.get("name","")
st.sidebar.title(f"👤 {op_name} ({role})")
st.sidebar.write(f"🔥 **V46.7 Pro Light Dashboard**")
last_log = get_last_daily_log()
if last_log: st.sidebar.caption(f"上次自動發送: {last_log}")
else: st.sidebar.caption("今日未自動發送 (17:00 HKT)")

if st.sidebar.button("登出",key="v467_logout"): st.session_state.user=None; st.cache_data.clear(); st.rerun()
if st.sidebar.button("🔄 同步OrderCosts",key="v467_clear"):
    cnt=sync_order_costs_sheet(); st.sidebar.success(f"已同步 {cnt} 單")
if st.sidebar.button("🔍 測試WhatsApp Group",key="v467_test_wa"):
    ok, info = send_whatsapp_to_ops(f"[V46.7測試]{now_hk().strftime('%m/%d %H:%M')} by {operator}")
    st.sidebar.success(f"✅ {info}") if ok else st.sidebar.error(f"❌ {info}")
if st.sidebar.button("📦 測試每日Excel",key="v467_test_daily_excel"):
    with st.spinner("生成緊..."):
        ok, info = generate_and_send_daily_excel()
        st.sidebar.success(f"✅ {info}") if ok else st.sidebar.warning(f"失敗 {info}")

if role=="admin": tabs=st.tabs(["📊 Pro Dashboard","📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費","💵 每單費用","📥 下載POD"])
elif role=="op": tabs=st.tabs(["📊 Pro Dashboard","📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費","💵 每單費用","📥 下載POD"])
else: tabs=st.tabs(["🟢 可接任務池","🚚 我的任務","📸 上傳POD","💰 雜費","📋 重複打印"])

df_jobs=load_df("Jobs"); df_exp=load_df("Expenses"); df_pods=load_df("PODs")

def get_my_jobs(df):
    if df.empty: return df
    mask = df["Driver"].astype(str).str.strip().str.lower().isin([operator.lower(), op_name.lower()])
    mask2 = df["Driver"].astype(str).str.contains(operator, na=False) | df["Driver"].astype(str).str.contains(op_name, na=False)
    return df[mask | mask2]

# ================= V46.7 PRO DASHBOARD =================
with tabs[0]:
    if role=="driver":
        st.subheader("🟢 可接任務池")
        df_pool=df_jobs[(df_jobs["Driver"].astype(str).str.strip()=="") & (df_jobs["Status"].astype(str).str.upper().isin(["PENDING",""]))] if not df_jobs.empty else pd.DataFrame()
        if df_pool.empty: st.info("暫無可接")
        else:
            for idx,row in df_pool.sort_values(by="JobID",ascending=False).head(20).iterrows():
                jid=row.get("JobID","")
                with st.container(border=True):
                    st.markdown(f"**{jid}** | {row.get('Client','')} | {row.get('Date','')} {row.get('DeliveryTime','')}")
                    if st.button("🟢 搶此單",key=f"v467_claim_{jid}_{idx}",type="primary"):
                        ws=get_sh().worksheet("Jobs"); headers=[h.strip() for h in ws.row_values(1)]; cell=ws.find(jid)
                        if cell:
                            if "Driver" in headers: ws.update_cell(cell.row, headers.index("Driver")+1, operator)
                            if "Status" in headers: ws.update_cell(cell.row, headers.index("Status")+1, "Accepted")
                            st.cache_data.clear(); st.success(f"已搶 {jid}"); st.rerun()
    else:
        st.markdown("""
        <style>
        .kpi-card {background:white; border-radius:12px; padding:16px 18px; box-shadow:0 2px 8px rgba(0,0,0,0.06); border:1px solid #eef1f6;}
        .kpi-title {font-size:13px; color:#6b7280; margin-bottom:4px;}
        .kpi-value {font-size:28px; font-weight:800; color:#111827;}
        .kpi-delta {font-size:12px; margin-top:4px;}
        .section-card {background:white; border-radius:14px; padding:18px; box-shadow:0 2px 12px rgba(0,0,0,0.05); border:1px solid #eef1f6; margin-bottom:16px;}
        </style>
        """, unsafe_allow_html=True)
        st.markdown("### 🚚 LogiTrack HK | Pro Dashboard <span style='font-size:14px;color:#6b7280'>Light Mode</span>", unsafe_allow_html=True)
        st.caption(f"Updated just now • Last sync: {now_hk().strftime('%H:%M HKT')} • Data as of {today_hk()}")

        if df_jobs.empty:
            st.info("暫無Jobs數據")
        else:
            # 解析
            df_jobs["Date_parsed"] = pd.to_datetime(df_jobs["Date"], errors="coerce")
            df_jobs["Status_clean"] = df_jobs["Status"].astype(str).str.strip().str.title()
            df_pods["JobID"] = df_pods["JobID"].astype(str).str.strip() if not df_pods.empty else ""
            df_jobs["JobID"] = df_jobs["JobID"].astype(str).str.strip()

            today_count = len(df_jobs[df_jobs["Date_parsed"].dt.date == today_hk()])
            pending_count = len(df_jobs[df_jobs["Status_clean"].str.upper()=="PENDING"])
            in_transit_count = len(df_jobs[df_jobs["Status_clean"].str.upper().isin(["ACCEPTED","IN TRANSIT","IN_TRANSIT"])])
            delivered_count = len(df_jobs[df_jobs["Status_clean"].str.upper()=="DELIVERED"])
            total_jobs = len(df_jobs)
            pod_rate = (len(df_pods["JobID"].unique()) / delivered_count * 100) if delivered_count>0 else 0

            # KPI 5卡
            k1,k2,k3,k4,k5 = st.columns(5)
            with k1: st.markdown(f'<div class="kpi-card"><div class="kpi-title">📅 Total Orders Today</div><div class="kpi-value">{today_count}</div><div class="kpi-delta" style="color:#16a34a">共 {total_jobs} 單在系統</div></div>', unsafe_allow_html=True)
            with k2: st.markdown(f'<div class="kpi-card"><div class="kpi-title">⏰ Pending</div><div class="kpi-value">{pending_count}</div><div class="kpi-delta" style="color:#d97706">待接單</div></div>', unsafe_allow_html=True)
            with k3: st.markdown(f'<div class="kpi-card"><div class="kpi-title">🚚 In Transit</div><div class="kpi-value" style="color:#0ea5e9">{in_transit_count}</div><div class="kpi-delta" style="color:#0ea5e9">運送中</div></div>', unsafe_allow_html=True)
            with k4: st.markdown(f'<div class="kpi-card"><div class="kpi-title">✅ Delivered</div><div class="kpi-value" style="color:#16a34a">{delivered_count}</div><div class="kpi-delta" style="color:#16a34a">已完成</div></div>', unsafe_allow_html=True)
            with k5: st.markdown(f'<div class="kpi-card"><div class="kpi-title">📸 POD Upload Rate</div><div class="kpi-value" style="color:#7c3aed">{pod_rate:.1f}%</div><div class="kpi-delta" style="color:#7c3aed">{len(df_pods)} 張POD</div></div>', unsafe_allow_html=True)

            st.write("")
            c_left, c_right = st.columns([2,1])
            with c_left:
                st.markdown('<div class="section-card">', unsafe_allow_html=True)
                st.markdown("**📈 Daily Deliveries - Last 7 Days**")
                last7 = [(today_hk() - datetime.timedelta(days=i)) for i in range(6,-1,-1)]
                daily_counts = []
                for d in last7:
                    cnt = len(df_jobs[df_jobs["Date_parsed"].dt.date == d])
                    daily_counts.append({"Date": d.strftime("%m/%d"), "Orders": cnt})
                df_chart = pd.DataFrame(daily_counts)
                st.bar_chart(df_chart, x="Date", y="Orders", height=280)
                st.markdown('</div>', unsafe_allow_html=True)

                st.markdown('<div class="section-card">', unsafe_allow_html=True)
                st.markdown("**🏆 司機排行榜 - 本月 Delivered**")
                if not df_jobs.empty and "Driver" in df_jobs.columns:
                    df_delivered = df_jobs[df_jobs["Status_clean"].str.upper()=="DELIVERED"]
                    if not df_delivered.empty:
                        leaderboard = df_delivered["Driver"].value_counts().reset_index()
                        leaderboard.columns = ["Driver","Delivered"]
                        st.dataframe(leaderboard.head(10), use_container_width=True, height=260)
                    else:
                        st.info("本月暫無Delivered")
                else:
                    st.info("無Driver數據")
                st.markdown('</div>', unsafe_allow_html=True)

            with c_right:
                st.markdown('<div class="section-card">', unsafe_allow_html=True)
                st.markdown("**🗺️ Zone Distribution**")
                if "Zone" in df_jobs.columns and not df_jobs["Zone"].dropna().empty:
                    zone_counts = df_jobs["Zone"].value_counts()
                    st.bar_chart(zone_counts, height=200)
                    for z,c in zone_counts.head(5).items():
                        pct = c/total_jobs*100 if total_jobs else 0
                        st.caption(f"{z}: {c} ({pct:.0f}%)")
                else:
                    st.caption("無Zone數據")
                st.markdown('</div>', unsafe_allow_html=True)

                st.markdown('<div class="section-card">', unsafe_allow_html=True)
                st.markdown("**🗺️ Live Delivery Routes - Hong Kong (Static)**")
                st.markdown("""
                <div style="background:#f8fafc; border-radius:10px; padding:12px; border:1px solid #e5e7eb; text-align:center;">
                <div style="font-size:42px;">🗺️</div>
                <div style="font-weight:700; margin-top:6px;">Hong Kong Delivery Network</div>
                <div style="font-size:12px; color:#6b7280; margin-top:4px;">Central / Kowloon • Kwai Tsing Port • Sha Tin • Tseung Kwan O</div>
                <div style="margin-top:10px; display:flex; justify-content:center; gap:12px; font-size:12px;">
                <span style="background:#dbeafe; padding:4px 10px; border-radius:20px;">● Active Delivery</span>
                <span style="background:#fef3c7; padding:4px 10px; border-radius:20px;">● 14 Routes</span>
                </div>
                </div>
                """, unsafe_allow_html=True)
                st.caption("靜態展示，唔用Google Map，慳Quota")
                st.markdown('</div>', unsafe_allow_html=True)

                st.markdown('<div class="section-card">', unsafe_allow_html=True)
                st.markdown("**⚠️ POD未上傳清單 (Delivered但無POD)**")
                if not df_jobs.empty and not df_pods.empty:
                    delivered_jobs = set(df_jobs[df_jobs["Status_clean"].str.upper()=="DELIVERED"]["JobID"].astype(str))
                    pod_jobs = set(df_pods["JobID"].astype(str))
                    missing = delivered_jobs - pod_jobs
                    if missing:
                        df_missing = df_jobs[df_jobs["JobID"].isin(missing)][["JobID","Client","Date","Driver","To"]].head(10)
                        st.dataframe(df_missing, use_container_width=True, height=200)
                        st.warning(f"共有 {len(missing)} 單已送達但未上傳POD")
                    else:
                        st.success("✅ 全部Delivered已上傳POD")
                else:
                    st.info("暫無數據")
                st.markdown('</div>', unsafe_allow_html=True)

            st.markdown('<div class="section-card">', unsafe_allow_html=True)
            st.markdown("**📋 Recent Jobs**")
            recent = df_jobs.sort_values(by="Date_parsed", ascending=False).head(10)[["JobID","Client","Zone","Status","Driver","Date","To"]]
            st.dataframe(recent, use_container_width=True, height=320)
            st.markdown('</div>', unsafe_allow_html=True)

# 其餘Tabs保持V46.6
if role in ["admin","op"]:
    with tabs[1]:
        st.subheader("📋 任務管理 + 打印")
        if df_jobs.empty: st.info("無任務")
        else:
            for idx,row in df_jobs.sort_values(by="JobID",ascending=False).head(100).iterrows():
                jid=row.get("JobID","")
                with st.container(border=True):
                    c1,c2 = st.columns([3,1])
                    with c1:
                        st.markdown(f"**{jid}** | {row.get('Client','')} | {row.get('Status','')} | Driver:{row.get('Driver','')}")
                        st.caption(f"{row.get('From','')} -> {row.get('To','')} | 📅 {row.get('Date','')} {row.get('DeliveryTime','')}")
                    with c2:
                        if st.button("🖨️ 打印", key=f"v467_print_{jid}_{idx}"): st.session_state[f"print_{jid}"] = True
                    if st.session_state.get(f"print_{jid}"):
                        st.components.v1.html(printable_order_html(row.to_dict()), height=750, scrolling=True)
    with tabs[2]:
        st.subheader("📤 上傳Order")
        up_file = st.file_uploader("📤 上傳Excel", type=["xlsx"], key="v467_excel_uploader")
        if up_file:
            try:
                df_check = pd.read_excel(up_file, sheet_name=0, header=0, nrows=2)
                is_new_template = "要貨單號" in df_check.columns or "商戶名稱" in df_check.columns
                up_file.seek(0)
                df_up = parse_company_export_template(up_file) if is_new_template else pd.read_excel(up_file, dtype=str).fillna("")
                st.dataframe(df_up.head(20), use_container_width=True)
                if st.button(f"✅ 確認批量建立 {len(df_up)} 單", type="primary", key="v467_excel_confirm"):
                    created=0
                    for _, r in df_up.iterrows():
                        client = str(r.get("商戶名稱","") or r.get("Client",""))
                        to_addr = str(r.get("收貨地址","") or r.get("To",""))
                        date_val = str(r.get("送貨日期","") or r.get("Date","") or today_hk())
                        if client.strip()=="" and to_addr.strip()=="": continue
                        jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                        data={"InternalNo":str(r.get("InternalNo","") or r.get("要貨單號","")), "JobID":jid, "Date":date_val, "Client":client, "Zone":"", "From":str(r.get("發貨倉庫","") or r.get("From","")), "From_Contact":str(r.get("發貨人","")), "From_Phone":str(r.get("發貨聯絡電話","")), "PickupTime":str(r.get("接貨時間","")), "To":to_addr, "To_Contact":str(r.get("收貨人","")), "To_Phone":str(r.get("聯係電話","")), "DeliveryTime":"", "CargoDetail":str(r.get("CargoDetail","")), "Notes":str(r.get("整單備注","") or r.get("Notes","")), "Status":"Pending", "Driver":"", "CreatedBy":operator, "CreatedAt_HKT":now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")}
                        append_row_by_header("Jobs", data); created+=1
                    st.success(f"🎉 批量建立 {created} 單")
            except Exception as e: st.error(f"Excel讀取失敗: {e}")
        with st.form("v467_order_form_full"):
            order_date = st.date_input("📅 訂單日期", value=today_hk())
            colA, colB = st.columns(2)
            with colA:
                client=st.text_input("客戶*"); zone=st.selectbox("Zone", ["", "HKI","KLN","NT","九龍","香港","新界","離島","OTHER"])
                from_loc=st.text_input("起點*"); from_contact=st.text_input("起點聯絡人"); from_phone=st.text_input("起點電話"); pickup_time=st.text_input("提貨時間")
            with colB:
                internal_no=st.text_input("InternalNo*"); to_loc=st.text_input("終點*"); to_contact=st.text_input("終點聯絡人"); to_phone=st.text_input("終點電話"); delivery_time=st.text_input("送貨時間")
            cargo=st.text_area("貨物詳情*", height=120); notes=st.text_input("備註")
            if st.form_submit_button("✅ 建立Order", type="primary", use_container_width=True):
                jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                data={"InternalNo":internal_no, "JobID":jid, "Date":str(order_date), "Client":client, "Zone":zone, "From":from_loc, "From_Contact":from_contact, "From_Phone":from_phone, "PickupTime":pickup_time, "To":to_loc, "To_Contact":to_contact, "To_Phone":to_phone, "DeliveryTime":delivery_time, "CargoDetail":cargo, "Notes":notes, "Status":"Pending", "Driver":"", "CreatedBy":operator, "CreatedAt_HKT":now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")}
                append_row_by_header("Jobs", data); st.success(f"🎉 已建立 {jid}")

pod_idx=3 if role in ["admin","op"] else 2
with tabs[pod_idx]:
    st.subheader("📸 上傳POD - V46.7")
    my_jobs = get_my_jobs(df_jobs) if role=="driver" else df_jobs
    my_pool=my_jobs.sort_values(by="JobID",ascending=False) if not my_jobs.empty else df_jobs
    if not my_pool.empty:
        options = my_pool.to_dict('records')
        labels = [format_job_option(r) for r in options]
        label_to_jid = dict(zip(labels, [r["JobID"] for r in options]))
        sel_label = st.selectbox("選擇JobID", labels, key="v467_pod_jid")
        jid = label_to_jid.get(sel_label, options[0]["JobID"])
        with st.form("v467_pod_form", clear_on_submit=True):
            pod_files=st.file_uploader("📷 拍照上傳POD",type=["jpg","png","jpeg"],accept_multiple_files=True)
            if st.form_submit_button("📤 上傳POD + 轉Delivered", type="primary", use_container_width=True) and pod_files:
                count_ok=0
                for pf in pod_files:
                    fb, _, _ = compress_pod_high_quality(pf)
                    if not fb: continue
                    b64=base64.b64encode(fb).decode()
                    pod_id=f"POD-{now_hk().strftime('%m%d%H%M%S')}-{uuid.uuid4().hex[:4]}"
                    if append_row_safe("PODs",[pod_id,jid,f"{jid}_POD.jpg",str(len(fb)),b64,operator,now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"),now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")]): count_ok+=1
                set_job_delivered(jid); st.success(f"🎉 上傳 {count_ok} 張，已轉Delivered")

expense_idx = 4 if role in ["admin","op"] else 3
with tabs[expense_idx]:
    st.subheader("💰 雜費")
    my_for_expense = get_my_jobs(df_jobs) if role=="driver" else df_jobs
    if not my_for_expense.empty:
        exp_labels = [format_job_option(r) for _,r in my_for_expense.iterrows()]
        sel_exp_label=st.selectbox("揀JobID", exp_labels, key="v467_exp_jid")
        sel_exp_jid = sel_exp_label.split(" | ")[0]
        with st.form("v467_expense_form"):
            amt_tunnel = st.number_input("隧道費", min_value=0.0, step=10.0)
            amt_parking = st.number_input("停車費", min_value=0.0, step=10.0)
            amt_other = st.number_input("其他", min_value=0.0, step=10.0)
            amt_reg = st.number_input("登記費", min_value=0.0, step=10.0)
            amt_wh = st.number_input("交倉附加費", min_value=0.0, step=10.0)
            if st.form_submit_button("✅ 提交雜費", type="primary"):
                exp_dict = {"隧道費":amt_tunnel,"登記費":amt_reg,"停車費":amt_parking,"交倉附加費":amt_wh,"其他":amt_other}
                for fee_type, amount in exp_dict.items():
                    if amount>0.01:
                        exp_id=f"EXP-{now_hk().strftime('%m%d%H%M%S')}-{uuid.uuid4().hex[:4]}"
                        append_row_safe("Expenses",[exp_id, str(today_hk()), sel_exp_jid, operator, fee_type, str(int(amount)), "", now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")])
                st.success("已提交")

if role=="driver":
    with tabs[1]:
        st.subheader("🚚 我的任務")
        my_jobs = get_my_jobs(df_jobs)
        if not my_jobs.empty:
            st.dataframe(my_jobs[["Date","JobID","Client","Status","From","To"]], use_container_width=True)
    with tabs[4]:
        st.subheader("📋 重複打印")
        if not df_jobs.empty:
            sel=st.selectbox("揀單打印", [format_job_option(r) for _,r in df_jobs.head(100).iterrows()])
            if sel:
                jid=sel.split(" | ")[0]; job=df_jobs[df_jobs["JobID"]==jid].iloc[0].to_dict()
                st.components.v1.html(printable_order_html(job), height=750, scrolling=True)

cost_idx = 5 if role in ["admin","op"] else None
if cost_idx is not None:
    with tabs[cost_idx]:
        st.subheader("💵 每單費用 OrderCosts")
        if st.button("🔄 重新同步", key="v467_sync_cost_btn", type="primary"): sync_order_costs_sheet(); st.rerun()
        df_cost=load_df("OrderCosts")
        if not df_cost.empty: st.dataframe(df_cost.sort_values(by="JobID",ascending=False), use_container_width=True, height=450)
    with tabs[6]:
        st.subheader("📥 下載POD - V46.7")
        df_pods = load_df("PODs")
        if df_pods.empty: st.warning("無PODs")
        else:
            df_pods["JobID"] = df_pods["JobID"].astype(str).str.strip()
            df_jobs["JobID"] = df_jobs["JobID"].astype(str).str.strip()
            df_merged = pd.merge(df_pods, df_jobs[["JobID","Date","Client","Driver","To","InternalNo"]] if not df_jobs.empty else pd.DataFrame(columns=["JobID"]), on="JobID", how="left")
            df_merged["CreatedAt_HKT_parsed"] = pd.to_datetime(df_merged["CreatedAt_HKT"], errors="coerce")
            c1,c2=st.columns(2)
            with c1: start_date = st.date_input("開始日期", value=today_hk() - datetime.timedelta(days=7))
            with c2: end_date = st.date_input("結束日期", value=today_hk())
            filtered = df_merged[(df_merged["CreatedAt_HKT_parsed"].dt.date >= start_date) & (df_merged["CreatedAt_HKT_parsed"].dt.date <= end_date)] if df_merged["CreatedAt_HKT_parsed"].notna().any() else df_merged
            st.success(f"找到 {len(filtered)} 張POD")
            if not filtered.empty and st.button(f"📦 打包下載 {len(filtered)} 張ZIP", type="primary"):
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for _, r in filtered.iterrows():
                        try:
                            b64_str = str(r.get("Base64_Image","")).strip()
                            if "," in b64_str: b64_str = b64_str.split(",",1)[1]
                            b64_str += "=" * (-len(b64_str) % 4)
                            img_bytes = base64.b64decode(b64_str)
                            zipf.writestr(f"{r.get('JobID','')}_{r.get('FileName','POD.jpg')}", img_bytes)
                        except: pass
                zip_buffer.seek(0)
                st.download_button("⬇️ 下載ZIP", data=zip_buffer, file_name=f"POD_{start_date}_to_{end_date}.zip", mime="application/zip")
            for idx, row in filtered.head(50).iterrows():
                with st.container(border=True):
                    c_img, c_info = st.columns([1,2])
                    with c_img:
                        try:
                            b64_str = str(row.get("Base64_Image","")).strip()
                            if "," in b64_str: b64_str = b64_str.split(",",1)[1]
                            b64_str += "=" * (-len(b64_str) % 4)
                            st.image(base64.b64decode(b64_str), use_container_width=True)
                        except:
                            st.error("圖片已損壞，請補上傳")
                    with c_info:
                        st.markdown(f"**{row.get('JobID','')}** | {row.get('Client','')}")
                        st.caption(f"📅 {row.get('CreatedAt_HKT','')} | 司機:{row.get('Driver','')}")
