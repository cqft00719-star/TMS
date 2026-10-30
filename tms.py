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

st.set_page_config(page_title="TMS V46.5 完整修復", layout="wide", page_icon="🚚")

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

# V46.5 修復：壓細啲確保入到Sheet 50000字限制
def compress_pod_high_quality(uploaded_file):
    try:
        img = Image.open(io.BytesIO(uploaded_file.getvalue()))
        img = ImageOps.exif_transpose(img).convert("RGB")
        for max_side in [1200, 1024, 900, 800, 700]:
            for quality in [70, 60, 50, 45]:
                buf = io.BytesIO()
                tmp = img.copy()
                tmp.thumbnail((max_side, max_side), Image.LANCZOS)
                tmp.save(buf, format="JPEG", quality=quality, optimize=True)
                fb = buf.getvalue()
                if len(fb) <= 35000:  # 改細，Base64約46666，入到50000限制
                    return fb, max_side, quality
        buf = io.BytesIO()
        img.thumbnail((600, 600), Image.LANCZOS)
        img.save(buf, format="JPEG", quality=40, optimize=True)
        return buf.getvalue(), 600, 40
    except Exception as e:
        print(f"compress fail {e}")
        return None, 0, 0

def watermark_pod_image(b64_data_url, job_id, internal_no=""):
    try:
        if "," in str(b64_data_url): b64_data_url = str(b64_data_url).split(",",1)[1]
        img_bytes = base64.b64decode(b64_data_url)
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
            # V46.5修復：PODs唔截斷Base64，其他先截
            if tab=="PODs" and len(s)>50000 and "data:image" in s:
                # 已經壓到35000，唔應該超過50000，如果超過都盡量保留
                s=s[:49900]
            elif len(s)>45000:
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

# V46.5 還原：V46.1原版完美打印樣式，唔改
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

if "user" not in st.session_state: st.session_state.user=None
if not st.session_state.user:
    st.title("TMS V46.5 登入")
    df_u=load_df("Users"); u=st.text_input("用戶名"); p=st.text_input("密碼",type="password")
    if st.button("登入",type="primary",key="v465_login"):
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
st.sidebar.write(f"🔥 當前版本: **V46.5 完整還原修復**")
st.sidebar.write(f"📱 {ULTRAMSG_INSTANCE} -> {ULTRAMSG_GROUP_ID}")
last_log = get_last_daily_log()
if last_log: st.sidebar.caption(f"上次自動發送: {last_log}")
else: st.sidebar.caption("今日未自動發送 (17:00 HKT)")

if st.sidebar.button("登出",key="v465_logout"): st.session_state.user=None; st.cache_data.clear(); st.rerun()
if st.sidebar.button("🔄 同步OrderCosts",key="v465_clear"):
    cnt=sync_order_costs_sheet(); st.sidebar.success(f"已同步 {cnt} 單")
if st.sidebar.button("🔍 測試WhatsApp Group",key="v465_test_wa"):
    test_msg = f"[V46.5測試]{now_hk().strftime('%m/%d %H:%M')} 測試通知 - by {operator}"
    ok, info = send_whatsapp_to_ops(test_msg)
    if ok: st.sidebar.success(f"✅ {info}");
    else: st.sidebar.error(f"❌ {info}")
if st.sidebar.button("📦 測試每日Excel附件發送",key="v465_test_daily_excel"):
    with st.spinner("生成緊聽日+後日Excel..."):
        ok, info = generate_and_send_daily_excel()
        if ok: 
            st.sidebar.success(f"✅ Excel附件已發去Group {info}")
            write_daily_log(f"MANUAL OK {info[:100]}")
        else: st.sidebar.warning(f"失敗 {info}")
if st.sidebar.button("📅 查看DailyLog",key="v465_view_log"):
    df_log = load_df("DailyLog")
    st.sidebar.dataframe(df_log.tail(10))

if role=="admin": tabs=st.tabs(["📊 Dashboard","📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費","💵 每單費用","📥 下載POD"])
elif role=="op": tabs=st.tabs(["📊 Dashboard","📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費","💵 每單費用","📥 下載POD"])
else: tabs=st.tabs(["🟢 可接任務池","🚚 我的任務","📸 上傳POD","💰 雜費","📋 重複打印"])

df_jobs=load_df("Jobs"); df_exp=load_df("Expenses")
def get_my_jobs(df):
    if df.empty: return df
    mask = df["Driver"].astype(str).str.strip().str.lower().isin([operator.lower(), op_name.lower()])
    mask2 = df["Driver"].astype(str).str.contains(operator, na=False) | df["Driver"].astype(str).str.contains(op_name, na=False)
    return df[mask | mask2]

with tabs[0]:
    if role=="driver":
        st.subheader("🟢 可接任務池 - 有送貨日子+預覽")
        df_pool=df_jobs[(df_jobs["Driver"].astype(str).str.strip()=="") & (df_jobs["Status"].astype(str).str.upper().isin(["PENDING",""]))] if not df_jobs.empty else pd.DataFrame()
        if df_pool.empty: st.info("暫無可接")
        else:
            if "Date" in df_pool.columns:
                try:
                    df_pool["Date_sort"]=pd.to_datetime(df_pool["Date"], errors="coerce")
                    df_pool=df_pool.sort_values(by="Date_sort", ascending=True)
                except:
                    df_pool=df_pool.sort_values(by="JobID",ascending=False)
            for idx,row in df_pool.iterrows():
                jid=row.get("JobID","")
                with st.container(border=True):
                    st.markdown(f"**{jid}** | {row.get('Client','')} | Internal:{row.get('InternalNo','')}")
                    st.write(f"📅 送貨: **{row.get('Date','')} {row.get('DeliveryTime','')}** | 客: {row.get('Client','')}")
                    st.write(f"從 {row.get('From','')} -> {row.get('To','')}")
                    st.caption(f"Cargo: {str(row.get('CargoDetail',''))[:120]}")
                    c1,c2=st.columns(2)
                    with c1:
                        if st.button("👁️ 預覽送貨單", key=f"v465_prev_{jid}_{idx}"):
                            st.session_state[f"preview_{jid}"] = not st.session_state.get(f"preview_{jid}", False)
                    with c2:
                        if st.button("🟢 搶此單",key=f"v465_claim_{jid}_{idx}",type="primary"):
                            ws=get_sh().worksheet("Jobs"); headers=[h.strip() for h in ws.row_values(1)]; cell=ws.find(jid)
                            if cell:
                                if "Driver" in headers: ws.update_cell(cell.row, headers.index("Driver")+1, operator)
                                if "Status" in headers: ws.update_cell(cell.row, headers.index("Status")+1, "Accepted")
                                st.cache_data.clear(); send_whatsapp_to_ops(f"[V46.5]JobID:{jid} 司機:{operator} 已接單"); st.success(f"已搶 {jid}"); st.rerun()
                    if st.session_state.get(f"preview_{jid}"):
                        st.components.v1.html(printable_order_html(row.to_dict()), height=700, scrolling=True)
    else:
        st.subheader("📊 Dashboard V46.5")
        if not df_jobs.empty:
            cols = ["InternalNo"] + [c for c in df_jobs.columns if c != "InternalNo"] if "InternalNo" in df_jobs.columns else df_jobs.columns.tolist()
            st.dataframe(df_jobs[cols].head(200), use_container_width=True)

if role in ["admin","op"]:
    with tabs[1]:
        st.subheader("📋 任務管理 + 打印")
        if df_jobs.empty: st.info("無任務")
        else:
            for idx,row in df_jobs.sort_values(by="JobID",ascending=False).head(100).iterrows():
                jid=row.get("JobID","")
                with st.container(border=True):
                    c1,c2,c3 = st.columns([3,1,1])
                    with c1:
                        st.markdown(f"**{jid}** | {row.get('Client','')} | Internal:{row.get('InternalNo','')} | {row.get('Status','')} | Driver:{row.get('Driver','')}")
                        st.caption(f"{row.get('From','')} -> {row.get('To','')} | 📅 {row.get('Date','')} {row.get('DeliveryTime','')}")
                    with c2:
                        if st.button("🖨️ 打印", key=f"v465_print_{jid}_{idx}"): st.session_state[f"print_{jid}"] = True
                    with c3: st.write(row.get('Zone',''))
                    if st.session_state.get(f"print_{jid}"):
                        st.components.v1.html(printable_order_html(row.to_dict()), height=750, scrolling=True)
                        if st.button("❌ 關閉打印", key=f"v465_close_{jid}_{idx}"): st.session_state[f"print_{jid}"] = False; st.rerun()
    with tabs[2]:
        st.subheader("📤 上傳Order")
        with st.expander("📊 Excel批量上傳", expanded=True):
            up_file = st.file_uploader("📤 上傳Excel批量建立", type=["xlsx"], key="v465_excel_uploader")
            if up_file:
                try:
                    df_check = pd.read_excel(up_file, sheet_name=0, header=0, nrows=2)
                    is_new_template = "要貨單號" in df_check.columns or "商戶名稱" in df_check.columns
                    up_file.seek(0)
                    if is_new_template:
                        df_up = parse_company_export_template(up_file)
                        st.success(f"✅ 識別到新公司訂單模板，已合併為 {len(df_up)} 張Order")
                        st.dataframe(df_up[["要貨單號","商戶名稱","商品明細數","CargoDetail"]].head(20), use_container_width=True)
                    else:
                        df_up = pd.read_excel(up_file, dtype=str).fillna("")
                    if st.button(f"✅ 確認批量建立 {len(df_up)} 單", type="primary", use_container_width=True, key="v465_excel_confirm"):
                        created = 0
                        for _, r in df_up.iterrows():
                            if is_new_template:
                                client = str(r.get("商戶名稱","") or r.get("Client","")); to_addr = str(r.get("收貨地址","") or r.get("To",""))
                                date_val = str(r.get("送貨日期","") or r.get("Date","") or today_hk())
                                if client.strip()=="" and to_addr.strip()=="": continue
                                jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                                data={"InternalNo":str(r.get("InternalNo","") or r.get("要貨單號","")), "JobID":jid, "Date":date_val, "Client":client, "Zone":"", "From":str(r.get("發貨倉庫","")), "From_Contact":str(r.get("發貨人","")), "From_Phone":str(r.get("發貨聯絡電話","")), "PickupTime":str(r.get("接貨時間","")), "To":to_addr, "To_Contact":str(r.get("收貨人","")), "To_Phone":str(r.get("聯係電話","")), "DeliveryTime":"", "CargoDetail":str(r.get("CargoDetail","")), "Notes":str(r.get("整單備注","")), "Status":"Pending", "Driver":"", "CreatedBy":operator, "CreatedAt_HKT":now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")}
                            else:
                                if str(r.get("Client","")).strip()=="" and str(r.get("From","")).strip()=="": continue
                                jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                                data={"InternalNo":str(r.get("InternalNo","")), "JobID":jid, "Date":str(r.get("Date","") or today_hk()), "Client":str(r.get("Client","")), "Zone":str(r.get("Zone","")), "From":str(r.get("From","")), "From_Contact":str(r.get("From_Contact","")), "From_Phone":str(r.get("From_Phone","")), "PickupTime":str(r.get("PickupTime","")), "To":str(r.get("To","")), "To_Contact":str(r.get("To_Contact","")), "To_Phone":str(r.get("To_Phone","")), "DeliveryTime":str(r.get("DeliveryTime","")), "CargoDetail":str(r.get("CargoDetail","")), "Notes":str(r.get("Notes","")), "Status":"Pending", "Driver":"", "CreatedBy":operator, "CreatedAt_HKT":now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")}
                            append_row_by_header("Jobs", data); created+=1
                        sync_order_costs_sheet(); st.success(f"🎉 批量建立 {created} 單成功！"); st.cache_data.clear()
                except Exception as e: st.error(f"Excel讀取失敗: {e}")

        with st.form("v465_order_form_full"):
            order_date = st.date_input("📅 選擇訂單日期 (Date)*", value=today_hk(), key="v465_manual_date")
            colA, colB = st.columns(2)
            with colA:
                client=st.text_input("客戶* (Client)"); zone=st.selectbox("Zone", ["", "HKI","KLN","NT","九龍","香港","新界","離島","OTHER"], key="v465_zone")
                from_loc=st.text_input("起點* (From)"); from_contact=st.text_input("起點聯絡人"); from_phone=st.text_input("起點電話"); pickup_time=st.text_input("提貨時間", placeholder="900")
            with colB:
                internal_no=st.text_input("InternalNo*"); to_loc=st.text_input("終點* (To)"); to_contact=st.text_input("終點聯絡人"); to_phone=st.text_input("終點電話"); delivery_time=st.text_input("送貨時間", placeholder="0900-1700")
            cargo=st.text_area("貨物詳情* (CargoDetail)", height=120); notes=st.text_input("備註 Notes")
            submitted = st.form_submit_button("✅ 手動建立Order", type="primary", use_container_width=True)
            if submitted:
                if not client or not from_loc or not to_loc or not cargo: st.error("請填寫 客戶* 起點* 終點* 貨物詳情*")
                else:
                    jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                    data={"InternalNo":internal_no, "JobID":jid, "Date":str(order_date), "Client":client, "Zone":zone, "From":from_loc, "From_Contact":from_contact, "From_Phone":from_phone, "PickupTime":pickup_time, "To":to_loc, "To_Contact":to_contact, "To_Phone":to_phone, "DeliveryTime":delivery_time, "CargoDetail":cargo, "Notes":notes, "Status":"Pending", "Driver":"", "CreatedBy":operator, "CreatedAt_HKT":now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")}
                    append_row_by_header("Jobs", data); sync_order_costs_sheet(); st.success(f"🎉 已建立 {jid}")

pod_idx=3 if role in ["admin","op"] else 2
with tabs[pod_idx]:
    st.subheader("📸 上傳POD - V46.5 即時預覽送貨單")
    my_jobs = get_my_jobs(df_jobs) if role=="driver" else df_jobs
    my_pool=my_jobs.sort_values(by="JobID",ascending=False) if not my_jobs.empty else df_jobs
    if not my_pool.empty:
        options = my_pool.to_dict('records')
        labels = [format_job_option(r) for r in options]
        label_to_jid = dict(zip(labels, [r["JobID"] for r in options]))
        sel_label = st.selectbox("選擇JobID (已顯示客戶+Internal+送貨日子)", labels, key="v465_pod_jid")
        jid = label_to_jid.get(sel_label, options[0]["JobID"])
        sel_row = my_pool[my_pool["JobID"]==jid]
        if not sel_row.empty:
            r=sel_row.iloc[0]
            st.info(f"已選: {r.get('Client','')} | {r.get('To','')} | 📅 {r.get('Date','')} {r.get('DeliveryTime','')}")
            with st.expander("👁️ 預覽此單送貨單 (避免揀錯)", expanded=True):
                st.components.v1.html(printable_order_html(r.to_dict()), height=650, scrolling=True)
        with st.form("v465_pod_form_final", clear_on_submit=True):
            pod_files=st.file_uploader("📷 拍照上傳POD (可多張)",type=["jpg","png","jpeg"],accept_multiple_files=True,key="v465_pod_files_final")
            submit_pod = st.form_submit_button("📤 上傳POD + 自動轉Delivered", type="primary", use_container_width=True)
            if submit_pod and pod_files:
                count_ok=0
                for idx_file, pf in enumerate(pod_files):
                    fb, max_side, qual = compress_pod_high_quality(pf)
                    if not fb: continue
                    b64=base64.b64encode(fb).decode()
                    data_url=f"data:image/jpeg;base64,{b64}"
                    pod_id=f"POD-{now_hk().strftime('%m%d%H%M%S')}-{uuid.uuid4().hex[:4]}"
                    ok_row = append_row_safe("PODs",[pod_id,jid,f"{jid}_POD_{now_hk().strftime('%Y%m%d_%H%M%S')}_{idx_file+1}.jpg",str(len(fb)),data_url,operator,now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"),now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")])
                    if ok_row: count_ok+=1
                ok_status = set_job_delivered(jid)
                sync_order_costs_sheet()
                msg = f"[V46.5 POD] JobID:{jid} 司機:{operator} 已上傳{count_ok}張高清POD 已轉Delivered:{ok_status}"
                send_whatsapp_to_ops(msg)
                st.cache_data.clear()
                if ok_status:
                    st.success(f"🎉 成功上傳 {count_ok} 張，Status已自動轉 Delivered")
                else:
                    st.warning(f"已上傳 {count_ok} 張，但Status自動轉換失敗")

expense_idx = 4 if role in ["admin","op"] else 3
with tabs[expense_idx]:
    st.subheader("💰 雜費 - 顯示客戶+日子+預覽")
    my_for_expense = get_my_jobs(df_jobs) if role=="driver" else df_jobs
    if not my_for_expense.empty:
        my_for_expense=my_for_expense.sort_values(by="JobID",ascending=False)
        exp_options = my_for_expense.to_dict('records')
        exp_labels = [format_job_option(r) for r in exp_options]
        exp_label_to_jid = dict(zip(exp_labels, [r["JobID"] for r in exp_options]))
        sel_exp_label=st.selectbox("揀JobID入雜費 (有客戶+日子)", exp_labels, key="v465_exp_jid")
        sel_exp_jid = exp_label_to_jid.get(sel_exp_label)
        sel_row = my_for_expense[my_for_expense["JobID"]==sel_exp_jid]
        if not sel_row.empty:
            r=sel_row.iloc[0]
            st.info(f"已選: {r.get('Client','')} | 📅 {r.get('Date','')} | {r.get('To','')}")
            with st.expander("👁️ 預覽此單送貨單", expanded=False):
                st.components.v1.html(printable_order_html(r.to_dict()), height=650, scrolling=True)
        with st.form("v465_expense_form_final", clear_on_submit=False):
            exp_date = st.date_input("日期", value=today_hk(), key="v465_exp_date")
            c1,c2=st.columns(2)
            with c1:
                amt_tunnel = st.number_input("隧道費", min_value=0.0, step=10.0, value=0.0)
                amt_parking = st.number_input("停車費", min_value=0.0, step=10.0, value=0.0)
                amt_other = st.number_input("其他", min_value=0.0, step=10.0, value=0.0)
            with c2:
                amt_reg = st.number_input("登記費", min_value=0.0, step=10.0, value=0.0)
                amt_wh = st.number_input("交倉附加費", min_value=0.0, step=10.0, value=0.0)
            if st.form_submit_button("✅ 提交雜費", type="primary", use_container_width=True):
                exp_dict = {"隧道費":amt_tunnel,"登記費":amt_reg,"停車費":amt_parking,"交倉附加費":amt_wh,"其他":amt_other}
                to_submit = {k:float(v) for k,v in exp_dict.items() if float(v)>0.01}
                if to_submit:
                    for fee_type, amount in to_submit.items():
                        exp_id=f"EXP-{now_hk().strftime('%m%d%H%M%S')}-{uuid.uuid4().hex[:4]}"
                        append_row_safe("Expenses",[exp_id, str(exp_date), sel_exp_jid, operator, fee_type, str(int(amount)), "", now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")])
                    sync_order_costs_sheet(); st.success("已提交")

if role=="driver":
    with tabs[1]:
        st.subheader("🚚 我的任務 - 可按送貨日子篩選")
        my_jobs = get_my_jobs(df_jobs)
        if my_jobs.empty:
            st.info("暫無任務")
        else:
            my_jobs["Date_parsed"]=pd.to_datetime(my_jobs["Date"], errors="coerce")
            c1,c2,c3=st.columns([1,1,1])
            with c1:
                show_mode=st.radio("顯示模式", ["全部","今日","聽日","指定日期"], horizontal=True, key="v465_my_mode")
            with c2:
                if show_mode=="指定日期":
                    filter_date=st.date_input("揀送貨日子", value=today_hk(), key="v465_my_filter_date")
                else:
                    filter_date=None
            with c3:
                status_filter=st.multiselect("狀態", ["Pending","Accepted","Delivered"], default=["Pending","Accepted"], key="v465_my_status")
            filtered=my_jobs.copy()
            if show_mode=="今日":
                filtered=filtered[filtered["Date_parsed"].dt.date==today_hk()]
            elif show_mode=="聽日":
                filtered=filtered[filtered["Date_parsed"].dt.date==today_hk()+datetime.timedelta(days=1)]
            elif show_mode=="指定日期" and filter_date:
                filtered=filtered[filtered["Date_parsed"].dt.date==filter_date]
            if status_filter:
                filtered=filtered[filtered["Status"].isin(status_filter)]
            st.success(f"找到 {len(filtered)} 單 | {show_mode}")
            st.dataframe(filtered.sort_values(by="Date_parsed")[["Date","DeliveryTime","JobID","Client","InternalNo","Status","From","To"]], use_container_width=True, height=400)
            if not filtered.empty:
                sel_my = st.selectbox("選一張預覽送貨單", [format_job_option(r) for _,r in filtered.iterrows()], key="v465_my_preview_sel")
                if sel_my:
                    jid_my = sel_my.split(" | ")[0]
                    row_my = filtered[filtered["JobID"]==jid_my]
                    if not row_my.empty:
                        st.components.v1.html(printable_order_html(row_my.iloc[0].to_dict()), height=650, scrolling=True)
    with tabs[4]:
        st.subheader("📋 重複打印 - 有客戶+日子+預覽")
        if not df_jobs.empty:
            all_opts=df_jobs.sort_values(by="JobID",ascending=False).head(200).to_dict('records')
            all_labels=[format_job_option(r) for r in all_opts]
            label_map=dict(zip(all_labels, [r["JobID"] for r in all_opts]))
            sel=st.selectbox("揀單打印 (有客戶+日子)", all_labels, key="v465_driver_print_sel")
            if sel:
                jid=label_map.get(sel); job=df_jobs[df_jobs["JobID"]==jid].iloc[0].to_dict()
                st.components.v1.html(printable_order_html(job), height=750, scrolling=True)

cost_idx = 5 if role in ["admin","op"] else None
if cost_idx is not None:
    with tabs[cost_idx]:
        st.subheader("💵 每單費用 OrderCosts V46.5")
        if st.button("🔄 立即重新同步", key="v465_sync_cost_btn", type="primary", use_container_width=True):
            cnt=sync_order_costs_sheet(); st.success(f"已同步 {cnt} 單"); st.rerun()
        df_cost=load_df("OrderCosts")
        if not df_cost.empty:
            st.dataframe(df_cost.sort_values(by="JobID",ascending=False), use_container_width=True, height=450)
    with tabs[6]:
        st.subheader("📥 下載POD - V46.5 修復版")
        df_pods = load_df("PODs")
        df_jobs_for_pod = load_df("Jobs")
        if df_pods.empty:
            st.warning("⚠️ PODs分頁暫無記錄")
        else:
            df_pods["JobID"] = df_pods["JobID"].astype(str).str.strip()
            df_jobs_for_pod["JobID"] = df_jobs_for_pod["JobID"].astype(str).str.strip() if not df_jobs_for_pod.empty else ""
            df_merged = pd.merge(df_pods, df_jobs_for_pod[["JobID","Date","Client","Driver","To","InternalNo"]] if not df_jobs_for_pod.empty else pd.DataFrame(columns=["JobID"]), on="JobID", how="left")

            # V46.5修復：雙日期容錯，舊記錄用Date
            df_merged["CreatedAt_HKT_parsed"] = pd.to_datetime(df_merged["CreatedAt_HKT"], errors="coerce")
            df_merged["Date_parsed"] = pd.to_datetime(df_merged["Date"], errors="coerce")
            df_merged["FilterDate"] = df_merged["CreatedAt_HKT_parsed"]
            df_merged["FilterDate"] = df_merged["FilterDate"].fillna(df_merged["Date_parsed"])

            st.markdown("#### 🔍 篩選時段")
            c1, c2, c3, c4 = st.columns(4)
            with c1: start_date = st.date_input("開始日期", value=today_hk() - datetime.timedelta(days=7), key="v465_start")
            with c2: end_date = st.date_input("結束日期", value=today_hk(), key="v465_end")
            with c3:
                driver_list = ["全部"] + sorted([x for x in df_merged["Driver"].dropna().astype(str).unique() if x.strip()]) if "Driver" in df_merged.columns else ["全部"]
                sel_driver = st.selectbox("司機", driver_list, key="v465_driver")
            with c4:
                client_list = ["全部"] + sorted([x for x in df_merged["Client"].dropna().astype(str).unique() if x.strip()]) if "Client" in df_merged.columns else ["全部"]
                sel_client = st.selectbox("客戶", client_list, key="v465_client")

            filtered = df_merged.copy()
            # 有FilterDate先篩，無就顯示全部避免空
            if filtered["FilterDate"].notna().any():
                filtered_valid = filtered[filtered["FilterDate"].notna()].copy()
                filtered_valid["FilterDateOnly"] = filtered_valid["FilterDate"].dt.date
                filtered_valid = filtered_valid[(filtered_valid["FilterDateOnly"] >= start_date) & (filtered_valid["FilterDateOnly"] <= end_date)]
                # 加埋無日期嘅都顯示，避免漏
                filtered_invalid = filtered[filtered["FilterDate"].isna()]
                filtered = pd.concat([filtered_valid, filtered_invalid])
            if sel_driver != "全部": filtered = filtered[filtered["Driver"].astype(str).str.contains(sel_driver, na=False)]
            if sel_client != "全部": filtered = filtered[filtered["Client"].astype(str).str.contains(sel_client, na=False)]

            st.success(f"找到 {len(filtered)} 張POD | 原始總數 {len(df_merged)} 張")

            if not filtered.empty:
                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    if st.button(f"📦 打包下載全部 {len(filtered)} 張 (ZIP)", type="primary", use_container_width=True, key="v465_dl_zip_all"):
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                            for _, r in filtered.iterrows():
                                try:
                                    b64_str = str(r.get("Base64_Image",""))
                                    if "," in b64_str: b64_str = b64_str.split(",",1)[1]
                                    if len(b64_str)<100: continue
                                    img_bytes = base64.b64decode(b64_str)
                                    wm_bytes, _ = watermark_pod_image(b64_str, r.get("JobID",""), r.get("InternalNo",""))
                                    final_bytes = wm_bytes if wm_bytes else img_bytes
                                    fname = f"{r.get('JobID','')}_{r.get('FileName','POD.jpg')}"
                                    fname = "".join(c for c in fname if c.isalnum() or c in "._- ").strip()
                                    zipf.writestr(fname, final_bytes)
                                except: pass
                        zip_buffer.seek(0)
                        st.download_button("⬇️ 點擊下載ZIP", data=zip_buffer, file_name=f"POD_{start_date}_to_{end_date}.zip", mime="application/zip", key="v465_dl_btn", use_container_width=True)
                with col_dl2:
                    csv_buf = filtered[["JobID","Client","Date","To","FileName","CreatedAt_HKT"]].to_csv(index=False).encode('utf-8-sig')
                    st.download_button("📄 下載清單CSV", data=csv_buf, file_name=f"POD_List_{start_date}_to_{end_date}.csv", mime="text/csv", use_container_width=True, key="v465_dl_csv")

                st.divider()
                for idx, row in filtered.sort_values(by="FilterDate", ascending=False, na_position='last').head(100).iterrows():
                    with st.container(border=True):
                        c_img, c_info = st.columns([1,2])
                        with c_img:
                            try:
                                b64_str = str(row.get("Base64_Image",""))
                                if "," in b64_str: b64_str = b64_str.split(",",1)[1]
                                if len(b64_str)>100:
                                    st.image(base64.b64decode(b64_str), use_container_width=True)
                                else:
                                    st.write("⚠️ 舊記錄無圖片")
                            except:
                                st.write("圖片失效，請重新上傳")
                        with c_info:
                            st.markdown(f"**{row.get('JobID','')}** | {row.get('Client','')} | {row.get('To','')}")
                            st.caption(f"📅 POD: {row.get('CreatedAt_HKT','')} | 送貨: {row.get('Date','')} | 司機: {row.get('Driver','')}")
                            try:
                                b64_str = str(row.get("Base64_Image",""))
                                if "," in b64_str: b64_str = b64_str.split(",",1)[1]
                                if len(b64_str)>100:
                                    img_bytes = base64.b64decode(b64_str)
                                    wm_bytes, _ = watermark_pod_image(b64_str, row.get("JobID",""), row.get("InternalNo",""))
                                    dl_bytes = wm_bytes if wm_bytes else img_bytes
                                    st.download_button(f"⬇️ 下載 {row.get('FileName','')}", data=dl_bytes, file_name=row.get('FileName','POD.jpg'), mime="image/jpeg", key=f"v465_dl_one_{idx}")
                            except: pass
