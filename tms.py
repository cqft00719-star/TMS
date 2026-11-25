import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime
from datetime import timezone, timedelta
import uuid, io, base64, requests, zipfile, os, time, json
from urllib.parse import quote
from PIL import Image, ImageDraw, ImageFont, ImageOps

HKT = timezone(timedelta(hours=8))
def now_hk(): return datetime.datetime.now(HKT)
def today_hk(): return now_hk().date()

def get_gcp_info():
    try:
        if "gcp_service_account" in st.secrets:
            return dict(st.secrets["gcp_service_account"])
    except Exception:
        pass
    try:
        if "GCP_JSON" in st.secrets:
            v = st.secrets["GCP_JSON"]
            if isinstance(v, dict):
                return dict(v)
            s = str(v).strip()
            if s.startswith("{"):
                return json.loads(s)
            return json.loads(base64.b64decode(s).decode())
    except Exception:
        pass
    s = os.getenv("GCP_JSON", "").strip()
    if s:
        if s.startswith("{"):
            return json.loads(s)
        return json.loads(base64.b64decode(s).decode())
    return None

try:
    from streamlit_autorefresh import st_autorefresh
    st_autorefresh(interval=60*1000, key="tms_autorefresh_60s")
except:
    pass

st.set_page_config(page_title="TMS V48 Edit + QR", layout="wide", page_icon="🚚")

NEW_FEES = ["隧道費","登記費","停車費","交倉附加費","其他"]
SCAN_SHOPS = ["TSTN","TSTG","TST","MK","MPMK","MOKO","CENT","HSP","CPE","CWB","CYW","AT1","AT2","ATRA","APM","KLN","TM","SH"]
SCAN_MIN_CHARGE = [200,200,200,200,200,200,380,380,380,380,380,500,500,500,200,200,250,250]
SCAN_PER_BOX = 5
SCAN_BOXES = {}

def scan_cost_for(shop, boxes):
    if shop not in SCAN_SHOPS:
        return 0, 0
    mc = SCAN_MIN_CHARGE[SCAN_SHOPS.index(shop)]
    included = mc // SCAN_PER_BOX
    if boxes <= included:
        return mc, mc
    return mc, mc + (boxes - included) * SCAN_PER_BOX

def build_tongxing_matrix():
    rows = []
    rows.append(["Date"] + SCAN_SHOPS + ["總數"])
    for d in range(1, 32):
        ds = f"2026-11-{d:02d}"
        label = f"{d}-Nov"
        boxes_map = SCAN_BOXES.get(ds, {})
        if not boxes_map:
            continue
        row = [label]
        day_total = 0
        for shop in SCAN_SHOPS:
            if shop in boxes_map:
                _, actual = scan_cost_for(shop, boxes_map[shop])
                row.append(actual)
                day_total += actual
            else:
                row.append("")
        row.append(day_total)
        rows.append(row)
    return rows

def build_zongbiao_matrix():
    rows = []
    rows.append(["Date"] + SCAN_SHOPS + ["總箱數"])
    for d in range(1, 32):
        ds = f"2026-11-{d:02d}"
        label = f"{d}-Nov"
        boxes_map = SCAN_BOXES.get(ds, {})
        if not boxes_map:
            continue
        row = [label]
        day_total = 0
        for shop in SCAN_SHOPS:
            if shop in boxes_map:
                row.append(boxes_map[shop])
                day_total += boxes_map[shop]
            else:
                row.append("")
        row.append(day_total)
        rows.append(row)
    return rows

def load_scan_boxes_from_sheet():
    try:
        sh = get_sh()
        ws = sh.worksheet("運輸總表")
        vals = ws.get_all_values()
        if len(vals) < 2:
            return
        for row in vals[1:]:
            if not row or not row[0].strip():
                continue
            try:
                raw = row[0].strip()
                if "-Nov" in raw:
                    day = int(raw.split("-")[0])
                    ds = f"2026-11-{day:02d}"
                elif "-" in raw and len(raw.split("-")) == 3:
                    ds = datetime.datetime.strptime(raw, "%Y-%m-%d").strftime("%Y-%m-%d")
                else:
                    continue
            except:
                continue
            boxes = {}
            for i, shop in enumerate(SCAN_SHOPS):
                v = row[i+1].strip() if i+1 < len(row) else ""
                if v and v.isdigit():
                    boxes[shop] = int(v)
            if boxes:
                SCAN_BOXES[ds] = boxes
    except Exception as e:
        print(f"load from sheet fail: {e}")

def rebuild_tongxing_sheet():
    return "已停用：為咗唔蓋咗 app.py 22:00 寫入嘅新數，呢個一鍵寫入已經停用。請直接喺 Google Sheet 睇，唔好喺呢度寫入。"

def write_cookie_manual(manual_date, boxes_dict):
    try:
        sh = get_sh()
        sd = manual_date.strftime("%d-%b")
        total_boxes = sum(boxes_dict.values())
        fee_total = 0
        for s, v in boxes_dict.items():
            if v:
                _, c = scan_cost_for(s, v)
                fee_total += c
        for sheet_name, is_cost in [("運輸總表", False), ("通行", True)]:
            ws = sh.worksheet(sheet_name)
            vals = ws.get_all_values()
            expected = ["Date"] + SCAN_SHOPS + ["總箱數"]
            if not vals or vals[0][0].strip() == "" or "200" in str(vals[0]):
                ws.update('A1', [expected], value_input_option='USER_ENTERED')
                vals = ws.get_all_values()
            ri = None
            for i, r in enumerate(vals[1:], start=2):
                if r and r[0].strip() == sd:
                    ri = i
                    break
            if not ri:
                ri = len(vals) + 1
                ws.update_cell(ri, 1, sd)
            if is_cost:
                rv = []
                for s in SCAN_SHOPS:
                    v = boxes_dict.get(s, 0)
                    if v == 0:
                        rv.append("")
                    else:
                        _, c = scan_cost_for(s, v)
                        rv.append(c)
                ws.update(f"B{ri}:S{ri}", [rv])
                ws.update_cell(ri, 20, fee_total)
            else:
                rv = [boxes_dict.get(s, 0) if boxes_dict.get(s, 0) != 0 else "" for s in SCAN_SHOPS]
                ws.update(f"B{ri}:S{ri}", [rv])
                ws.update_cell(ri, 20, total_boxes)
            try:
                for c_idx, s in enumerate(SCAN_SHOPS, start=2):
                    if boxes_dict.get(s, 0) > 0:
                        col = chr(64 + c_idx) if c_idx <= 26 else "A" + chr(64 + c_idx - 26)
                        ws.format(f"{col}{ri}", {"backgroundColor": {"red": 1, "green": 1, "blue": 0}})
                ws.format(f"T{ri}", {"backgroundColor": {"red": 1, "green": 1, "blue": 0}})
            except:
                pass
        st.cache_data.clear()
        return True, total_boxes, fee_total, sd
    except Exception as e:
        return False, 0, 0, str(e)

def build_tongxing_excel_bytes():
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    wb = Workbook()
    ws1 = wb.active; ws1.title = "通行"
    ws2 = wb.create_sheet("運輸總表")
    fills = {"TSTN":"FFFF00","TSTG":"FFFF00","TST":"FFFF00","MK":"92D050","MPMK":"92D050","MOKO":"92D050","CENT":"92D050","HSP":"92D050","CPE":"92D050","CWB":"92D050","CYW":"92D050","AT1":"FF0000","AT2":"FF0000","ATRA":"FF0000","APM":"FFFF00","KLN":"FFFF00","TM":"00B0F0","SH":"00B0F0"}
    thin = Side(style="thin", color="B0B0B0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for ws, matrix, is_cost in [(ws1, build_tongxing_matrix(), True), (ws2, build_zongbiao_matrix(), False)]:
        header_row = 1
        for r_idx, row in enumerate(matrix, start=1):
            for c_idx, val in enumerate(row, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if r_idx == header_row and 2 <= c_idx <= 19:
                    cell.fill = PatternFill("solid", fgColor=fills.get(SCAN_SHOPS[c_idx-2], "FFFFFF"))
                    cell.font = Font(bold=True)
                if r_idx == header_row and c_idx == 20:
                    cell.font = Font(bold=True)
                if r_idx > header_row and c_idx == 1:
                    cell.font = Font(bold=True)
                if r_idx > header_row and c_idx == 20 and val != "":
                    cell.font = Font(bold=True)
        ws.column_dimensions['A'].width = 12
        for i in range(2, 20):
            ws.column_dimensions[chr(64+i)].width = 10
        ws.column_dimensions['T'].width = 12
        ws.freeze_panes = "B2"
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()

# ===== V48: Group ID 已改做真 VIP Group =====
EVO_BASE_URL = "https://evolution-api-production-5d91de27.up.railway.app"
EVO_INSTANCE = "my WhatsApps"
EVO_API_KEY = st.secrets["EVO_API_KEY"]
ULTRAMSG_GROUP_ID = "120363410834827562@g.us"
MY_WHATSAPP = st.secrets.get("MY_WHATSAPP", "")
DAILY_AUTO_HOUR = 17
DAILY_AUTO_MINUTE = 0
# QR 用：Streamlit Secrets 加 APP_URL = "https://你個-tms-網址"
APP_URL = ""
try:
    APP_URL = str(st.secrets.get("APP_URL", "")).strip()
except:
    APP_URL = ""

def get_qr_url_for_job(jid):
    # 如果有 APP_URL，QR Scan 會直接開 Apps 嗰張單；如果冇，就淨係編 JobID，司機 Scan 到都知係邊張單
    if APP_URL:
        base = APP_URL.rstrip("/")
        return f"{base}/?job={quote(str(jid).strip())}"
    return str(jid).strip()

def make_qr_img_tag(jid, size=110):
    # 唔依賴 qrcode library，用 qrserver 畫，Streamlit 唔使另外裝嘢
    data = quote(get_qr_url_for_job(jid))
    url = f"https://api.qrserver.com/v1/create-qr-code/?size={size}x{size}&data={data}"
    return url

def send_ultramsg_message(to, body):
    try:
        url = f"{EVO_BASE_URL}/message/sendText/{quote(EVO_INSTANCE)}"
        headers = {"apikey": EVO_API_KEY, "Content-Type": "application/json"}
        r = requests.post(url, json={"number": to, "text": body}, headers=headers, timeout=20)
        return r.status_code in [200, 201], r.text[:500]
    except Exception as e:
        return False, str(e)

def send_ultramsg_document(to, file_path, caption=""):
    try:
        with open(file_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("utf-8")
        url = f"{EVO_BASE_URL}/message/sendMedia/{quote(EVO_INSTANCE)}"
        headers = {"apikey": EVO_API_KEY, "Content-Type": "application/json"}
        payload = {"number": to, "mediatype": "document", "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "caption": caption, "media": b64, "fileName": os.path.basename(file_path)}
        r = requests.post(url, json=payload, headers=headers, timeout=60)
        return r.status_code in [200, 201], r.text[:1000]
    except Exception as e:
        return False, str(e)

def send_whatsapp_to_ops(message):
    return send_ultramsg_message(ULTRAMSG_GROUP_ID, message)
def send_whatsapp_to_me(message):
    if not MY_WHATSAPP:
        return False, "MY_WHATSAPP 未設定"
    return send_ultramsg_message(MY_WHATSAPP, message)

def send_excel_to_me(file_bytes, filename, caption):
    try:
        if not MY_WHATSAPP:
            return False, "MY_WHATSAPP 未設定"
        b64 = base64.b64encode(file_bytes).decode()
        url = f"{EVO_BASE_URL}/message/sendMedia/{quote(EVO_INSTANCE)}"
        headers = {"apikey": EVO_API_KEY, "Content-Type": "application/json"}
        payload = {"number": MY_WHATSAPP, "mediatype": "document", "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "caption": caption, "media": b64, "fileName": filename}
        r = requests.post(url, json=payload, headers=headers, timeout=60)
        return r.status_code in [200, 201], r.text[:500]
    except Exception as e:
        return False, str(e)

def notify_new_order(job_dict, operator_name=""):
    try:
        msg = f"🆕{job_dict.get('JobID','')} | {job_dict.get('InternalNo','')}\n👤 {job_dict.get('Client','')}\n📍 {job_dict.get('From','')} -> {job_dict.get('To','')}\n📅 {job_dict.get('Date','')} {job_dict.get('DeliveryTime','')}\n👷 By {operator_name} {now_hk().strftime('%H:%M HKT')}"
        ok, info = send_whatsapp_to_ops(msg)
        if not ok: st.error(f"WhatsApp 通知失敗: {info[:300]}")
        return ok, info
    except Exception as e:
        st.error(f"WhatsApp 通知失敗: {e}")
        return False, str(e)

def notify_bulk_orders(count, sample_jobs, operator_name=""):
    try:
        lines = "\n".join([f"- {j.get('JobID','')} {j.get('Client','')}" for j in sample_jobs[:5]])
        msg = f"📦{count} 單 By {operator_name}\n{lines}\n{'+'+str(count-5)+'單...' if count>5 else ''} {now_hk().strftime('%H:%M HKT')}"
        ok, info = send_whatsapp_to_ops(msg)
        if not ok: st.error(f"WhatsApp 通知失敗: {info[:300]}")
        return ok, info
    except Exception as e:
        st.error(f"WhatsApp 通知失敗: {e}")
        return False, str(e)

def notify_claim(jid, driver, client=""):
    try:
        ok, info = send_whatsapp_to_ops(f"🚚{driver} 已接單 {jid} | {client} | {now_hk().strftime('%H:%M HKT')}")
        if not ok: st.error(f"WhatsApp 通知失敗: {info[:300]}")
    except Exception as e:
        st.error(f"WhatsApp 通知失敗: {e}")

def notify_delivered(jid, driver, client="", pod_count=0):
    try:
        ok, info = send_whatsapp_to_ops(f"✅{jid} | {client} | 司機:{driver} | POD:{pod_count}張 | {now_hk().strftime('%H:%M HKT')}")
        if not ok: st.error(f"WhatsApp 通知失敗: {info[:300]}")
    except Exception as e:
        st.error(f"WhatsApp 通知失敗: {e}")

def notify_expense(jid, driver, client, fee_dict, total):
    try:
        parts = " ".join([f"{k}:${v:.0f}" for k,v in fee_dict.items() if v>0.01])
        ok, info = send_whatsapp_to_ops(f"💰雜費 {jid} | {client} | 司機:{driver} | {parts} 合計:${total:.0f} | {now_hk().strftime('%H:%M HKT')}")
        if not ok: st.error(f"WhatsApp 通知失敗: {info[:300]}")
    except Exception as e:
        st.error(f"WhatsApp 通知失敗: {e}")

# ===== V48: 修改 Order 通知 =====
def notify_order_edited(jid, client, operator_name, changed_fields):
    try:
        ch = "、".join(changed_fields) if changed_fields else "內容"
        ok, info = send_whatsapp_to_ops(f"✏️ Order 已修改 {jid} | {client} | 改咗: {ch} | 由 {operator_name} 改 | {now_hk().strftime('%H:%M HKT')}\n司機請睇最新地址 / 時間，唔好照舊單去。")
        if not ok: st.error(f"WhatsApp 通知失敗: {info[:300]}")
    except Exception as e:
        st.error(f"WhatsApp 通知失敗: {e}")

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
        merged = {"商戶類型": first.get("商戶類型",""), "商戶名稱": first.get("商戶名稱",""), "發貨倉庫": first.get("發貨倉庫",""), "發貨人": first.get("發貨人",""), "發貨聯絡電話": first.get("發貨聯絡電話",""), "接貨時間": first.get("接貨時間",""), "收貨地址": first.get("收貨地址",""), "收貨人": first.get("收貨人",""), "聯係電話": first.get("聯係電話",""), "送貨日期": first.get("送貨日期",""), "要貨單號": order_no, "創建日期": first.get("創建日期",""), "整單備注": first.get("整單備注",""), "CargoDetail": "\n".join(cargo_lines), "商品明細數": len(group), "InternalNo": order_no, "Client": str(first.get("商戶名稱","")).strip(), "From": str(first.get("發貨倉庫","")).strip(), "To": str(first.get("收貨地址","")).strip(), "Date": str(first.get("送貨日期","")).split(" ")[0] if first.get("送貨日期") else str(today_hk())}
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
            date_col = c; break
    if not date_col:
        return False, "無Date欄"
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
        if df.empty or "Date" not in df.columns:
            return None
        df = df[df["Date"].astype(str).str.strip()!=""]
        if df.empty: return None
        return str(df.iloc[-1]["Date"]).strip()
    except:
        return None

def write_daily_log(status):
    try:
        sh = get_sh()
        try:
            ws = sh.worksheet("DailyLog")
        except:
            ws = sh.add_worksheet(title="DailyLog", rows=1000, cols=10)
            ws.update('A1', [["Date","Time_HKT","Status","Detail","TriggeredBy"]])
        ws.append_row([str(today_hk()), now_hk().strftime("%H:%M:%S"), status, str(now_hk()), "manual-tms.py"], value_input_option='USER_ENTERED')
        st.cache_data.clear()
    except Exception as e:
        print(f"DailyLog write fail {e}")

def check_and_auto_send_daily():
    return

# ===== V48: 修改 Order - 直接寫返去同一行 =====
def update_job_in_sheet(jid, new_vals, old_row_dict):
    try:
        ws = get_sh().worksheet("Jobs")
        headers = [h.strip() for h in ws.row_values(1)]
        if "JobID" not in headers:
            return False, "Jobs 冇 JobID 欄"
        cell = ws.find(str(jid).strip())
        if not cell:
            return False, f"搵唔到 {jid}"
        row_no = cell.row
        changed = []
        for col_name, new_v in new_vals.items():
            if col_name not in headers:
                continue
            old_v = str(old_row_dict.get(col_name, ""))
            if str(new_v) != old_v:
                changed.append(col_name)
                ws.update_cell(row_no, headers.index(col_name)+1, str(new_v))
        # 喺 Notes 加返修改紀錄，等之後有得追
        try:
            if changed and "Notes" in headers:
                old_notes = str(old_row_dict.get("Notes", "") or "")
                add = f"[修改 {now_hk().strftime('%m/%d %H:%M')} by {st.session_state.get('user', {}).get('username','')}: 改咗 {','.join(changed)}]"
                ws.update_cell(row_no, headers.index("Notes")+1, (old_notes + " " + add).strip())
        except:
            pass
        st.cache_data.clear()
        return True, changed
    except Exception as e:
        return False, str(e)

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

def compress_pod_for_sheet(uploaded_file):
    try:
        img = Image.open(io.BytesIO(uploaded_file.getvalue()))
        img = ImageOps.exif_transpose(img).convert("RGB")
        for max_side, quality in [(800, 50), (600, 40), (500, 35), (400, 30), (350, 25), (300, 20)]:
            buf = io.BytesIO()
            tmp = img.copy()
            tmp.thumbnail((max_side, max_side), Image.LANCZOS)
            tmp.save(buf, format="JPEG", quality=quality, optimize=True)
            fb = buf.getvalue()
            b64 = base64.b64encode(fb).decode()
            if len(b64) < 49000:
                return fb, b64
        return None, ""
    except Exception as e:
        print(f"compress for sheet fail {e}")
        return None, ""

@st.cache_resource
def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(get_gcp_info(), scopes=scopes)
    return gspread.authorize(creds)
def get_sh():
    return get_client().open_by_key(st.secrets["connections"]["gsheets"]["spreadsheet"])
@st.cache_data(ttl=30)
def load_df(tab):
    try:
        ws=get_sh().worksheet(tab)
        vals=ws.get_all_values()
        if len(vals)<1:
            return pd.DataFrame()
        headers=[h.strip() for h in vals[0]]
        clean=[]
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
    except:
        return pd.DataFrame()

def append_row_safe(tab,row):
    try:
        safe_row=[]
        for c in row:
            s=str(c)
            if len(s) > 49000: s = s[:49000]
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
    ws=get_sh().worksheet(tab)
    headers=[h.strip() for h in ws.row_values(1)]
    row=[""]*len(headers)
    for k,v in data_dict.items():
        if k in headers:
            row[headers.index(k)]=str(v)
    return append_row_safe(tab,row)

def sync_order_costs_sheet():
    try:
        df_jobs=load_df("Jobs")
        df_exp=load_df("Expenses")
        if df_jobs.empty: return 0
        df_cost=df_jobs.copy()
        df_cost["TotalExpense"]=0
        for fee in NEW_FEES: df_cost[fee]=0
        if not df_exp.empty and "JobID" in df_exp.columns:
            df_exp["Amount_num"]=pd.to_numeric(df_exp["Amount"], errors="coerce").fillna(0)
            total_map=df_exp.groupby("JobID")["Amount_num"].sum().to_dict()
            df_cost["TotalExpense"]=df_cost["JobID"].map(total_map).fillna(0)
            if "Type" in df_exp.columns:
                pivot=df_exp.pivot_table(index="JobID", columns="Type", values="Amount_num", aggfunc="sum", fill_value=0)
                for fee in NEW_FEES:
                    if fee in pivot.columns:
                        df_cost[fee]=df_cost["JobID"].map(pivot[fee].to_dict()).fillna(0)
        headers=["InternalNo","JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","DeliveryTime","CargoDetail","Status","Driver"]+NEW_FEES+["TotalExpense","LastSync_HKT"]
        df_cost["LastSync_HKT"]=now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")
        for h in headers:
            if h not in df_cost.columns: df_cost[h]=""
        ws=get_sh().worksheet("OrderCosts")
        ws.clear()
        ws.update('A1',[headers])
        if not df_cost.empty:
            rows=[[str(r.get(h,"")) for h in headers] for _,r in df_cost.sort_values(by="JobID",ascending=False).iterrows()]
            for i in range(0,len(rows),100):
                ws.append_rows(rows[i:i+100], value_input_option='USER_ENTERED')
                time.sleep(0.5)
        st.cache_data.clear()
        return len(df_cost)
    except Exception as e:
        st.error(f"OrderCosts同步失敗: {e}")
        return 0

def ensure_sheets():
    sh=get_sh()
    existing=[ws.title for ws in sh.worksheets()]
    def add(t,h,d=None):
        if t not in existing:
            ws=sh.add_worksheet(title=t, rows=3000, cols=40)
            ws.update('A1',[h])
            if d: ws.append_rows(d, value_input_option='USER_ENTERED')
    add("Users",["username","password","name","role"],[["admin","admin123","Admin","admin"],["阿傑","123","阿傑","driver"],["John","John123","John","driver"],["Mo","Mo123","Mo","driver"],["op01","op01","OP01","op"],["Kit","Kit123","Kit","driver"],["Felix","Felix123","Felix Kwok","driver"]])
    add("Jobs",["InternalNo","JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","DeliveryTime","Driver","Status","CargoDetail","InternalNo","Notes","CreatedBy","CreatedAt_HKT","PodTime_HKT"])
    add("PODs",["POD_ID","JobID","FileName","Size","Base64_Image","CreatedBy","CreatedAt_HKT","PodTime_HKT"])
    add("Expenses",["ExpenseID","Date","JobID","Driver","Type","Amount","Remark","CreatedAt_HKT"])
    add("OrderCosts",["InternalNo","JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","DeliveryTime","CargoDetail","Status","Driver","隧道費","登記費","停車費","交倉附加費","其他","TotalExpense","LastSync_HKT"])
    add("DailyLog",["Date","Time_HKT","Status","Detail","TriggeredBy"])
    add("通行",["Date"]+SCAN_SHOPS+["總數"])
    add("運輸總表",["Date"]+SCAN_SHOPS+["總箱數"])

def printable_order_html(job):
    jid = job.get('JobID',''); internal = job.get('InternalNo',''); client = job.get('Client',''); zone = job.get('Zone','')
    from_addr = job.get('From',''); from_contact = job.get('From_Contact',''); from_phone = job.get('From_Phone',''); pickup = job.get('PickupTime','')
    to_addr = job.get('To',''); to_contact = job.get('To_Contact',''); to_phone = job.get('To_Phone',''); delivery = job.get('DeliveryTime','')
    cargo = job.get('CargoDetail',''); notes = job.get('Notes','') or ""; date_str = job.get('Date','')
    expected = f"{date_str} {pickup}-{delivery}".strip() or delivery or pickup or date_str
    cargo_html = str(cargo).replace('\n','<br>') if cargo else ""
    qr_url = make_qr_img_tag(jid, size=120)
    return f"""<html><head><meta charset="utf-8"><title>Printable - {jid}</title><style>@media print {{.no-print {{ display:none; }} @page {{ margin:10mm; }} }} body{{font-family: "Microsoft JhengHei", Arial, sans-serif; font-size:13px;}}.page{{border:2px solid #000; max-width:830px; margin:10px auto; padding:18px 22px; position:relative;}} h1{{text-align:center; font-size:22px; margin:5px 0 8px 0;}}.subtitle{{text-align:center; font-size:12px; margin-bottom:12px; border-bottom:1px solid #999; padding-bottom:8px;}} table{{width:100%; border-collapse:collapse; margin:6px 0;}} td{{border:1px solid #000; padding:6px 8px;}}.label{{background:#f2f2f2; width:110px; font-weight:bold;}}.label2{{background:#f2f2f2; width:55px; font-weight:bold; text-align:center;}}.section{{font-size:15px; font-weight:bold; margin:14px 0 4px 0;}}.sign{{margin-top:18px; display:flex; justify-content:space-between; font-size:12px;}}.hint{{text-align:center; font-size:11px; color:#666; margin-top:10px;}}.qr-box{{position:absolute; top:14px; right:18px; text-align:center;}}.qr-box img{{width:110px; height:110px; border:1px solid #000;}}.qr-box div{{font-size:10px; font-weight:bold; margin-top:2px;}}</style></head><body><div class="page"><div class="qr-box"><img src="{qr_url}"><div>Scan 入單 / POD<br>{jid}</div></div><h1>🚚 運輸訂單 Transport Order</h1><div class="subtitle">JobID: {jid} | Internal: {internal} | 預計: {expected}</div><table><tr><td class="label">客戶 Client</td><td>{client}</td><td class="label2">Zone</td><td style="width:130px;">{zone}</td></tr></table><div class="section">📦 起點 Pickup</div><table><tr><td class="label">地址 From</td><td colspan="3">{from_addr}</td></tr><tr><td class="label">聯絡人</td><td>{from_contact}</td><td class="label2">電話</td><td style="width:150px;">{from_phone}</td></tr><tr><td class="label">提貨時間</td><td colspan="3">{pickup}</td></tr></table><div class="section">🏁 終點 Delivery</div><table><tr><td class="label">地址 To</td><td colspan="3">{to_addr}</td></tr><tr><td class="label">聯絡人</td><td>{to_contact}</td><td class="label2">電話</td><td style="width:150px;">{to_phone}</td></tr><tr><td class="label">送貨時間</td><td colspan="3">{delivery}</td></tr></table><div class="section">📋 貨物資料 Cargo</div><table><tr><td class="label">詳細明細</td><td>{cargo_html}</td></tr><tr><td class="label">備註</td><td>{notes}</td></tr></table><div class="sign"><div>司機簽收: __________________</div><div>客戶簽收: __________________</div><div>日期: {date_str}</div></div></div><div class="hint no-print"><button onclick="window.print()" style="margin-top:10px; padding:10px 22px; font-size:15px;">🖨️ 打印 / Save as PDF</button></div></body></html>"""
def format_job_option(row):
    jid=str(row.get("JobID","")); client=str(row.get("Client",""))[:15]; internal=str(row.get("InternalNo","")); date=str(row.get("Date",""))[:10]; delivery=str(row.get("DeliveryTime",""))
    return f"{jid} | {client} | {internal} | {date} {delivery}".strip()

def get_pool_jobs(df):
    if df.empty:
        return df
    if "Driver" not in df.columns or "Status" not in df.columns:
        return df.iloc[0:0]
    drv = df["Driver"].astype(str).str.strip().str.lower()
    stt = df["Status"].astype(str).str.strip().str.upper()
    mask = (drv == "") | (drv == "nan") | (drv == "none")
    mask = mask & stt.isin(["PENDING", "", "NAN", "NONE"])
    return df[mask].copy()

if st.query_params.get("auto_daily") == "1":
    ok, info = generate_and_send_daily_excel()
    write_daily_log(f"MANUAL auto_daily=1 {'OK' if ok else 'FAIL'} {info[:120]}")
    st.success(f"手動每日提醒已執行 (auto_daily=1): {info}")
    st.stop()

if "user" not in st.session_state:
    st.session_state.user=None
if not st.session_state.user:
    st.title("TMS V48 登入 - Edit + QR")
    df_u=load_df("Users")
    u=st.text_input("用戶名")
    p=st.text_input("密碼",type="password")
    if st.button("登入",type="primary",key="v48_login"):
        m=df_u[(df_u["username"].astype(str).str.strip()==u.strip()) & (df_u["password"].astype(str).str.strip()==p.strip())]
        if not m.empty:
            st.session_state.user=m.iloc[0].to_dict()
            st.rerun()
        else:
            st.error("帳號錯")
    st.stop()
ensure_sheets()
user=st.session_state.user
role=str(user.get("role","driver")).lower()
operator=user.get("username","")
op_name=user.get("name","")
# ===== V48: QR Scan 入嚟，記低個 JobID =====
qr_job = ""
try:
    qr_job = str(st.query_params.get("job", "") or "").strip()
    # 有啲 QR 會帶 TMS:JOB-xxx 或者成條 link，抽返個 JOB- 出嚟
    if qr_job and "JOB-" in qr_job.upper():
        import re as _re
        _m = _re.search(r"JOB-[A-Z0-9\-]+", qr_job.upper())
        if _m: qr_job = _m.group(0)
except:
    qr_job = ""
if qr_job:
    st.session_state["qr_job"] = qr_job
qr_job = st.session_state.get("qr_job", "")
st.sidebar.title(f"👤 {op_name} ({role})")
st.sidebar.write(f"🔥 **V48 Edit + QR**")
st.sidebar.caption("OCR 已停用，曲奇改做手動入數")
if qr_job:
    st.sidebar.success(f"📷 QR 已認到：{qr_job}")
    if st.sidebar.button("❌ 清除 QR 單", key="v48_clear_qr"):
        st.session_state["qr_job"] = ""
        st.query_params.clear()
        st.rerun()
last_log = get_last_daily_log()
if last_log:
    st.sidebar.caption(f"上次自動發送: {last_log}")
else:
    st.sidebar.caption("今日未自動發送 (17:00 HKT, 由 app.py 負責)")
if st.sidebar.button("登出",key="v48_logout"):
    st.session_state.user=None
    st.cache_data.clear()
    st.rerun()
if st.sidebar.button("🔄 同步OrderCosts",key="v48_sync"):
    cnt=sync_order_costs_sheet()
    st.sidebar.success(f"已同步 {cnt} 單")
if st.sidebar.button("🔍 測試WhatsApp Group",key="v48_test_wa"):
    test_msg = f"[V48測試]{now_hk().strftime('%m/%d %H:%M')} by {operator}"
    ok, info = send_whatsapp_to_ops(test_msg)
    if ok: st.sidebar.success(f"✅ 已發送: {info[:200]}")
    else: st.sidebar.error(f"❌ 失敗: {info[:300]}")
if st.sidebar.button("📦 測試每日Excel (手動, 唔會自動 Send)",key="v48_test_daily_excel"):
    with st.spinner("生成緊..."):
        ok, info = generate_and_send_daily_excel()
        if ok:
            st.sidebar.success(f"✅ {info} (手動)")
            write_daily_log(f"MANUAL OK {info[:100]}")
        else:
            st.sidebar.warning(f"失敗 {info}")

if role=="admin":
    tabs=st.tabs(["📊 Pro Dashboard","📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費","💵 每單費用","📥 下載POD","📦 通行 SCAN"])
elif role=="op":
    tabs=st.tabs(["📊 Pro Dashboard","📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費","💵 每單費用","📥 下載POD","📦 通行 SCAN"])
else:
    tabs=st.tabs(["🟢 可接任務池","🚚 我的任務","📸 上傳POD","💰 雜費","📋 重複打印"])

df_jobs=load_df("Jobs")
df_exp=load_df("Expenses")
df_pods=load_df("PODs")

def get_my_jobs(df):
    if df.empty: return df
    if "Driver" not in df.columns: return df.iloc[0:0]
    drv = df["Driver"].astype(str).str.strip()
    mask = (drv.str.lower() == operator.strip().lower()) | (drv.str.lower() == op_name.strip().lower())
    return df[mask].copy()
def get_my_active_jobs(df):
    my = get_my_jobs(df)
    if my.empty: return my
    if "Status" not in my.columns: return my
    stt = my["Status"].astype(str).str.strip().str.upper()
    active = stt.isin(["ACCEPTED","IN TRANSIT","IN_TRANSIT","IN-TRANSIT","PICKED UP","PICKEDUP","ON THE WAY","ON_THE_WAY","ASSIGNED"])
    return my[active].copy()
def get_my_history_jobs(df):
    my = get_my_jobs(df)
    if my.empty: return my
    stt = my["Status"].astype(str).str.strip().str.upper()
    return my[stt.isin(["DELIVERED","COMPLETED","CANCELLED","CANCEL"] )].copy()

with tabs[0]:
    if role=="driver":
        st.subheader("🟢 可接任務池")
        # QR 嚟嘅單，置頂 show
        if qr_job and not df_jobs.empty:
            qr_row = df_jobs[df_jobs["JobID"].astype(str).str.strip() == qr_job]
            if not qr_row.empty:
                r = qr_row.iloc[0]
                st.success(f"📷 QR 認到 {qr_job}，已經幫你搵到喇，唔使再打字搵。")
                with st.container(border=True):
                    st.markdown(f"**{r.get('JobID','')}** | {r.get('Client','')} | {r.get('Date','')} {r.get('DeliveryTime','')} | Status: {r.get('Status','')} | Driver: {r.get('Driver','')}")
                    st.caption(f"{r.get('From','')} -> {r.get('To','')} | Internal: {r.get('InternalNo','')}")
                    st.write(f"📦 {str(r.get('CargoDetail',''))[:200]}")
                    if str(r.get('Status','')).strip().upper() == "PENDING" and not str(r.get('Driver','')).strip():
                        if st.button(f"🟢 搶呢張 QR 單 {qr_job}", key=f"v48_claim_qr_{qr_job}", type="primary"):
                            ws=get_sh().worksheet("Jobs")
                            headers=[h.strip() for h in ws.row_values(1)]
                            cell=ws.find(qr_job)
                            if cell:
                                if "Driver" in headers: ws.update_cell(cell.row, headers.index("Driver")+1, operator)
                                if "Status" in headers: ws.update_cell(cell.row, headers.index("Status")+1, "Accepted")
                                st.cache_data.clear()
                                notify_claim(qr_job, operator, r.get('Client',''))
                                st.success(f"已搶 {qr_job}")
                                st.rerun()
                    else:
                        st.info("呢張單已經有人接咗 / 唔係 Pending，請去「上傳POD」嗰頁，嗰度已經幫你揀好咗呢張單。")
                st.divider()
            else:
                st.warning(f"QR 嗰張 {qr_job} 搵唔到，可能 JobID 打錯咗，請試下下面搜尋。")
        df_pool = get_pool_jobs(df_jobs)
        st.caption(f"而家有 {len(df_pool)} 張可接單 (Driver 空白 + Status Pending)。如果見唔到某張，請用下面搜尋打客戶名 / JobID。")
        if df_pool.empty:
            st.info("暫無可接")
        else:
            pool_search = st.text_input("🔍 搜尋可接單 (打 Gucci / JobID / 地址)", placeholder="例如：Gucci、JOB-1104、Times Square...", key="v48_pool_search")
            df_pool_show = df_pool.copy()
            if pool_search.strip():
                s = pool_search.strip().lower()
                mask = df_pool_show.apply(lambda r: s in str(r.get("JobID","")).lower() or s in str(r.get("Client","")).lower() or s in str(r.get("To","")).lower() or s in str(r.get("InternalNo","")).lower() or s in str(r.get("From","")).lower(), axis=1)
                df_pool_show = df_pool_show[mask]
                if df_pool_show.empty:
                    st.warning(f"搵唔到包含「{pool_search}」嘅可接單。呢張單可能已經俾人搶咗 (Driver 已填)，或者 Status 唔係 Pending，請去 Sheet 睇 N / O 欄。")
                else:
                    st.success(f"搵到 {len(df_pool_show)} 張")
            for idx,row in df_pool_show.sort_values(by="JobID",ascending=False).iterrows():
                jid=row.get("JobID","")
                with st.container(border=True):
                    st.markdown(f"**{jid}** | {row.get('Client','')} | {row.get('Date','')} {row.get('DeliveryTime','')}")
                    st.caption(f"{row.get('From','')} -> {row.get('To','')} | Internal: {row.get('InternalNo','')}")
                    if st.button("🟢 搶此單",key=f"v48_claim_{jid}_{idx}",type="primary"):
                        ws=get_sh().worksheet("Jobs")
                        headers=[h.strip() for h in ws.row_values(1)]
                        cell=ws.find(jid)
                        if cell:
                            if "Driver" in headers: ws.update_cell(cell.row, headers.index("Driver")+1, operator)
                            if "Status" in headers: ws.update_cell(cell.row, headers.index("Status")+1, "Accepted")
                            st.cache_data.clear()
                            notify_claim(jid, operator, row.get('Client',''))
                            st.success(f"已搶 {jid}，已通知WhatsApp")
                            st.rerun()
    else:
        st.markdown("### 🚚 LogiTrack HK | Pro Dashboard V48")
        st.caption(f"Updated just now • Last sync: {now_hk().strftime('%H:%M HKT')} • Data as of {today_hk()} | ✅ 即時WhatsApp已啟用")
        if df_jobs.empty:
            st.info("暫無Jobs數據")
        else:
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
            k1,k2,k3,k4,k5 = st.columns(5)
            with k1: st.metric("📅 Total Today", today_count)
            with k2: st.metric("⏰ Pending", pending_count)
            with k3: st.metric("🚚 In Transit", in_transit_count)
            with k4: st.metric("✅ Delivered", delivered_count)
            with k5: st.metric("📸 POD Rate", f"{pod_rate:.1f}%")
            st.write("")
            c_left, c_right = st.columns([2,1])
            with c_left:
                st.markdown("**📈 Daily Deliveries - Last 7 Days**")
                last7 = [(today_hk() - datetime.timedelta(days=i)) for i in range(6,-1,-1)]
                daily_counts = []
                for d in last7:
                    daily_counts.append({"Date": d.strftime("%m/%d"), "Orders": len(df_jobs[df_jobs["Date_parsed"].dt.date == d])})
                st.bar_chart(pd.DataFrame(daily_counts), x="Date", y="Orders", height=280)
                st.markdown("**🏆 司機排行榜**")
                if not df_jobs.empty and "Driver" in df_jobs.columns:
                    df_delivered = df_jobs[df_jobs["Status_clean"].str.upper()=="DELIVERED"]
                    if not df_delivered.empty:
                        leaderboard = df_delivered["Driver"].value_counts().reset_index()
                        leaderboard.columns = ["Driver","Delivered"]
                        st.dataframe(leaderboard.head(10), use_container_width=True, height=260)
                    else:
                        st.info("本月暫無Delivered")
            with c_right:
                st.markdown("**🗺️ Zone Distribution**")
                if "Zone" in df_jobs.columns and not df_jobs["Zone"].dropna().empty:
                    st.bar_chart(df_jobs["Zone"].value_counts(), height=200)
                st.markdown("**⚠️ POD未上傳清單**")
                if not df_jobs.empty and not df_pods.empty:
                    delivered_jobs = set(df_jobs[df_jobs["Status_clean"].str.upper()=="DELIVERED"]["JobID"].astype(str))
                    pod_jobs = set(df_pods["JobID"].astype(str))
                    missing = delivered_jobs - pod_jobs
                    if missing:
                        st.dataframe(df_jobs[df_jobs["JobID"].isin(missing)][["JobID","Client","Date","Driver","To"]].head(10), use_container_width=True, height=200)
            st.markdown("**📋 Recent Jobs**")
            st.dataframe(df_jobs.sort_values(by="Date_parsed", ascending=False).head(10)[["JobID","Client","Zone","Status","Driver","Date","To"]], use_container_width=True, height=320)

if role in ["admin","op"]:
    with tabs[1]:
        st.subheader("📋 任務管理 + 打印 + 修改")
        st.caption("✏️ 修改：只有 admin / op 先改到。Delivered 咗嘅單唔俾改。改完會自動通知 VIP Group。")
        if qr_job and not df_jobs.empty:
            st.info(f"📷 QR 認到 {qr_job}，下面嗰張就係嗰張單。")
        if df_jobs.empty:
            st.info("無任務")
        else:
            # QR 嚟嘅單排最前
            df_sorted = df_jobs.sort_values(by="JobID",ascending=False)
            if qr_job:
                df_sorted = pd.concat([df_sorted[df_sorted["JobID"].astype(str).str.strip()==qr_job], df_sorted[df_sorted["JobID"].astype(str).str.strip()!=qr_job]])
            for idx,row in df_sorted.head(100).iterrows():
                jid=row.get("JobID","")
                with st.container(border=True):
                    c1,c2,c3 = st.columns([3,1,1])
                    with c1:
                        st.markdown(f"**{jid}** | {row.get('Client','')} | {row.get('Status','')} | Driver:{row.get('Driver','')}")
                        st.caption(f"{row.get('From','')} -> {row.get('To','')} | 📅 {row.get('Date','')} {row.get('DeliveryTime','')}")
                    with c2:
                        if st.button("🖨️ 打印", key=f"v48_print_{jid}_{idx}"):
                            st.session_state[f"print_{jid}"] = True
                    with c3:
                        if st.button("✏️ 修改", key=f"v48_edit_{jid}_{idx}"):
                            st.session_state[f"edit_{jid}"] = True
                    if st.session_state.get(f"print_{jid}"):
                        st.components.v1.html(printable_order_html(row.to_dict()), height=750, scrolling=True)
                    # ===== 修改 Form =====
                    if st.session_state.get(f"edit_{jid}"):
                        if str(row.get("Status","")).strip().upper() == "DELIVERED":
                            st.warning(f"⚠️ {jid} 已經 Delivered，唔俾喺度改。如真係要改，請去 Google Sheet 直接改，或者開張新單。")
                            if st.button("關閉修改", key=f"v48_edit_close_delivered_{jid}"):
                                st.session_state[f"edit_{jid}"] = False
                                st.rerun()
                        else:
                            st.markdown(f"**✏️ 修改 {jid}** - 改完撳儲存，會寫返去同一行，唔會開新單。")
                            with st.form(f"v48_edit_form_{jid}"):
                                ec1, ec2 = st.columns(2)
                                with ec1:
                                    e_client = st.text_input("客戶", value=str(row.get("Client","")), key=f"e_client_{jid}")
                                    e_zone = st.text_input("Zone", value=str(row.get("Zone","")), key=f"e_zone_{jid}")
                                    e_from = st.text_input("起點", value=str(row.get("From","")), key=f"e_from_{jid}")
                                    e_from_contact = st.text_input("起點聯絡人", value=str(row.get("From_Contact","")), key=f"e_from_c_{jid}")
                                    e_from_phone = st.text_input("起點電話", value=str(row.get("From_Phone","")), key=f"e_from_p_{jid}")
                                    e_pickup = st.text_input("提貨時間", value=str(row.get("PickupTime","")), key=f"e_pickup_{jid}")
                                with ec2:
                                    e_internal = st.text_input("InternalNo", value=str(row.get("InternalNo","")), key=f"e_internal_{jid}")
                                    e_date = st.text_input("日期 (YYYY-MM-DD)", value=str(row.get("Date","")), key=f"e_date_{jid}")
                                    e_to = st.text_input("終點", value=str(row.get("To","")), key=f"e_to_{jid}")
                                    e_to_contact = st.text_input("終點聯絡人", value=str(row.get("To_Contact","")), key=f"e_to_c_{jid}")
                                    e_to_phone = st.text_input("終點電話", value=str(row.get("To_Phone","")), key=f"e_to_p_{jid}")
                                    e_delivery = st.text_input("送貨時間", value=str(row.get("DeliveryTime","")), key=f"e_delivery_{jid}")
                                e_cargo = st.text_area("貨物詳情", value=str(row.get("CargoDetail","")), height=100, key=f"e_cargo_{jid}")
                                e_notes = st.text_input("備註", value=str(row.get("Notes","")), key=f"e_notes_{jid}")
                                e_driver = st.text_input("司機 (留空 = 未派)", value=str(row.get("Driver","")), key=f"e_driver_{jid}")
                                e_status = st.selectbox("狀態", ["Pending","Accepted","In Transit","Delivered","Cancelled"], index=["Pending","Accepted","In Transit","Delivered","Cancelled"].index(str(row.get("Status","Pending")).strip().title()) if str(row.get("Status","Pending")).strip().title() in ["Pending","Accepted","In Transit","Delivered","Cancelled"] else 0, key=f"e_status_{jid}")
                                bc1, bc2 = st.columns(2)
                                with bc1:
                                    do_save = st.form_submit_button("✅ 儲存修改 + 通知 Group", type="primary", use_container_width=True)
                                with bc2:
                                    do_cancel = st.form_submit_button("取消", use_container_width=True)
                                if do_cancel:
                                    st.session_state[f"edit_{jid}"] = False
                                    st.rerun()
                                if do_save:
                                    new_vals = {"Client": e_client, "Zone": e_zone, "From": e_from, "From_Contact": e_from_contact, "From_Phone": e_from_phone, "PickupTime": e_pickup, "InternalNo": e_internal, "Date": e_date, "To": e_to, "To_Contact": e_to_contact, "To_Phone": e_to_phone, "DeliveryTime": e_delivery, "CargoDetail": e_cargo, "Notes": e_notes, "Driver": e_driver, "Status": e_status}
                                    ok_e, changed_e = update_job_in_sheet(jid, new_vals, row.to_dict())
                                    if ok_e:
                                        if changed_e:
                                            notify_order_edited(jid, e_client, operator, changed_e)
                                            st.success(f"🎉 已修改 {jid}，改咗：{','.join(changed_e)}，已通知 VIP Group")
                                        else:
                                            st.info("冇嘢改到，唔使通知。")
                                        st.session_state[f"edit_{jid}"] = False
                                        st.cache_data.clear()
                                        st.rerun()
                                    else:
                                        st.error(f"修改失敗：{changed_e}")
    with tabs[2]:
        st.subheader("📤 上傳Order (自動通知WhatsApp)")
        up_file = st.file_uploader("📤 上傳Excel", type=["xlsx"], key="v48_excel_uploader")
        if up_file:
            try:
                df_check = pd.read_excel(up_file, sheet_name=0, header=0, nrows=2)
                is_new_template = "要貨單號" in df_check.columns or "商戶名稱" in df_check.columns
                up_file.seek(0)
                df_up = parse_company_export_template(up_file) if is_new_template else pd.read_excel(up_file, dtype=str).fillna("")
                st.dataframe(df_up.head(20), use_container_width=True)
                if st.button(f"✅ 確認批量建立 {len(df_up)} 單 + 發WhatsApp", type="primary", key="v48_excel_confirm"):
                    created=0; created_list=[]
                    for _, r in df_up.iterrows():
                        client = str(r.get("商戶名稱","") or r.get("Client",""))
                        to_addr = str(r.get("收貨地址","") or r.get("To",""))
                        date_val = str(r.get("送貨日期","") or r.get("Date","") or today_hk())
                        if client.strip()=="" and to_addr.strip()=="":
                            continue
                        jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                        data={"InternalNo":str(r.get("InternalNo","") or r.get("要貨單號","")), "JobID":jid, "Date":date_val, "Client":client, "Zone":"", "From":str(r.get("發貨倉庫","") or r.get("From","")), "From_Contact":str(r.get("發貨人","")), "From_Phone":str(r.get("發貨聯絡電話","")), "PickupTime":str(r.get("接貨時間","")), "To":to_addr, "To_Contact":str(r.get("收貨人","")), "To_Phone":str(r.get("聯係電話","")), "DeliveryTime":"", "CargoDetail":str(r.get("CargoDetail","")), "Notes":str(r.get("整單備注","") or r.get("Notes","")), "Status":"Pending", "Driver":"", "CreatedBy":operator, "CreatedAt_HKT":now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")}
                        if append_row_by_header("Jobs", data):
                            created+=1; created_list.append(data)
                    notify_bulk_orders(created, created_list, operator)
                    st.success(f"🎉 批量建立 {created} 單，已自動通知WhatsApp")
            except Exception as e:
                st.error(f"Excel讀取失敗: {e}")
        with st.form("v48_order_form_full"):
            order_date = st.date_input("📅 訂單日期", value=today_hk())
            colA, colB = st.columns(2)
            with colA:
                client=st.text_input("客戶*")
                zone=st.selectbox("Zone", ["", "HKI","KLN","NT","九龍","香港","新界","離島","OTHER"])
                from_loc=st.text_input("起點*")
                from_contact=st.text_input("起點聯絡人")
                from_phone=st.text_input("起點電話")
                pickup_time=st.text_input("提貨時間")
            with colB:
                internal_no=st.text_input("InternalNo*")
                to_loc=st.text_input("終點*")
                to_contact=st.text_input("終點聯絡人")
                to_phone=st.text_input("終點電話")
                delivery_time=st.text_input("送貨時間")
            cargo=st.text_area("貨物詳情*", height=120)
            notes=st.text_input("備註")
            if st.form_submit_button("✅ 建立Order + 發WhatsApp", type="primary", use_container_width=True):
                jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                data={"InternalNo":internal_no, "JobID":jid, "Date":str(order_date), "Client":client, "Zone":zone, "From":from_loc, "From_Contact":from_contact, "From_Phone":from_phone, "PickupTime":pickup_time, "To":to_loc, "To_Contact":to_contact, "To_Phone":to_phone, "DeliveryTime":delivery_time, "CargoDetail":cargo, "Notes":notes, "Status":"Pending", "Driver":"", "CreatedBy":operator, "CreatedAt_HKT":now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")}
                if append_row_by_header("Jobs", data):
                    notify_new_order(data, operator)
                    st.success(f"🎉 已建立 {jid}，已通知WhatsApp")

pod_idx=3 if role in ["admin","op"] else 2
with tabs[pod_idx]:
    st.subheader("📸 上傳POD - 自動通知")
    st.caption("💡 提示：影相嗰陣盡量企近啲、光啲，系統會自動壓縮到入到 Sheet 為止。")
    if qr_job:
        st.success(f"📷 QR 認到 {qr_job}，下面已經幫你揀好咗呢張單，唔使再搵。")
    my_jobs_for_pod = get_my_jobs(df_jobs) if role=="driver" else df_jobs
    if role in ["admin","op"] and not df_jobs.empty:
        active_all = df_jobs[df_jobs["Status"].astype(str).str.strip().str.upper().isin(["PENDING","ACCEPTED","IN TRANSIT","IN_TRANSIT"])]
        my_jobs_for_pod = active_all if not active_all.empty else df_jobs
    my_pool=my_jobs_for_pod.sort_values(by="JobID",ascending=False) if not my_jobs_for_pod.empty else df_jobs
    if my_pool.empty:
        st.info("暫無可上傳 POD 嘅單")
    else:
        search = st.text_input("🔍 搜尋 JobID / 客戶 / 地址 / InternalNo", placeholder="例如：JOB-1117、TSTN、客戶名...", key="v48_pod_search")
        filtered_pool = my_pool.copy()
        if search.strip():
            s = search.strip().lower()
            mask = filtered_pool.apply(lambda r: s in str(r.get("JobID","")).lower() or s in str(r.get("Client","")).lower() or s in str(r.get("To","")).lower() or s in str(r.get("InternalNo","")).lower() or s in str(r.get("From","")).lower(), axis=1)
            filtered_pool = filtered_pool[mask]
            if filtered_pool.empty:
                st.warning(f"搵唔到包含「{search}」嘅單，請試第二個關鍵字")
                st.stop()
            st.success(f"搵到 {len(filtered_pool)} 單")
        options = filtered_pool.to_dict('records')
        labels = [format_job_option(r) for r in options]
        label_to_job = {format_job_option(r): r for r in options}
        # QR 嚟嘅單，預設揀佢
        default_idx = 0
        if qr_job:
            for _i, _o in enumerate(options):
                if str(_o.get("JobID","")).strip() == qr_job:
                    default_idx = _i
                    break
        sel_label = st.selectbox("✅ 選擇要上傳 POD 嘅單（已按搜尋 Filter）", labels, index=default_idx, key="v48_pod_jid_v2")
        sel_job = label_to_job.get(sel_label, options[0])
        jid = sel_job.get("JobID","")
        sel_client = sel_job.get('Client','')
        with st.container(border=True):
            st.markdown(f"**你揀咗：{jid}**")
            st.write(f"👤 客戶：{sel_job.get('Client','')} | 📅 {sel_job.get('Date','')} {sel_job.get('DeliveryTime','')}")
            st.write(f"📍 {sel_job.get('From','')} → {sel_job.get('To','')}")
            st.write(f"🚚 司機：{sel_job.get('Driver','')} | 狀態：{sel_job.get('Status','')}")
            st.write(f"📦 貨物：{str(sel_job.get('CargoDetail',''))[:200]}")
        pod_files = st.file_uploader("📷 拍照上傳POD（一次最多 3 張，系統會自動壓縮）", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="v48_pod_files_noform")
        if pod_files and len(pod_files) > 3:
            st.warning("⚠️ 一次最多 3 張，系統只會上傳頭 3 張")
            pod_files = pod_files[:3]
        if st.button("📤 上傳POD + 轉Delivered + 通知", type="primary", use_container_width=True, key="v48_pod_upload_noform"):
            if not pod_files:
                st.error("請先揀相, 揀完先再撳上傳")
            else:
                count_ok = 0; fail_list = []
                for pf in pod_files:
                    try:
                        fb, b64 = compress_pod_for_sheet(pf)
                        if not fb or not b64:
                            fail_list.append(f"{pf.name} (壓極都仲係太大, 試下影近啲、光啲)")
                            continue
                        pod_id = f"POD-{now_hk().strftime('%m%d%H%M%S')}-{uuid.uuid4().hex[:4]}"
                        ok = append_row_safe("PODs", [pod_id, jid, f"{jid}_POD.jpg", str(len(fb)), b64, operator, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"), now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")])
                        if ok: count_ok += 1
                        else: fail_list.append(f"{pf.name} (寫入 PODs 失敗)")
                    except Exception as e:
                        fail_list.append(f"{pf.name} ({e})")
                if count_ok > 0:
                    set_job_delivered(jid)
                    notify_delivered(jid, operator, sel_client, count_ok)
                    st.success(f"🎉 上傳 {count_ok} 張，已轉 Delivered，已通知WhatsApp")
                    st.balloons()
                    st.cache_data.clear()
                if fail_list:
                    st.error(f"❌ 以下 {len(fail_list)} 張失敗：{', '.join(fail_list)}")

expense_idx = 4 if role in ["admin","op"] else 3
with tabs[expense_idx]:
    st.subheader("💰 雜費")
    my_for_expense = get_my_active_jobs(df_jobs) if role=="driver" else df_jobs
    if my_for_expense.empty and role=="driver":
        my_for_expense = get_my_jobs(df_jobs)
    if not my_for_expense.empty:
        exp_labels = [format_job_option(r) for _,r in my_for_expense.iterrows()]
        sel_exp_label=st.selectbox("揀JobID", exp_labels, key="v48_exp_jid")
        sel_exp_jid = sel_exp_label.split(" | ")[0]
        sel_exp_client = next((str(r.get("Client","")) for _,r in my_for_expense.iterrows() if str(r.get("JobID",""))==sel_exp_jid), "")
        with st.form("v48_expense_form"):
            amt_tunnel = st.number_input("隧道費", min_value=0.0, step=10.0)
            amt_parking = st.number_input("停車費", min_value=0.0, step=10.0)
            amt_other = st.number_input("其他", min_value=0.0, step=10.0)
            amt_reg = st.number_input("登記費", min_value=0.0, step=10.0)
            amt_wh = st.number_input("交倉附加費", min_value=0.0, step=10.0)
            if st.form_submit_button("✅ 提交雜費", type="primary"):
                exp_dict = {"隧道費":amt_tunnel,"登記費":amt_reg,"停車費":amt_parking,"交倉附加費":amt_wh,"其他":amt_other}
                total_fee = sum(exp_dict.values())
                ok_any = False
                for fee_type, amount in exp_dict.items():
                    if amount>0.01:
                        exp_id=f"EXP-{now_hk().strftime('%m%d%H%M%S')}-{uuid.uuid4().hex[:4]}"
                        if append_row_safe("Expenses",[exp_id, str(today_hk()), sel_exp_jid, operator, fee_type, str(int(amount)), "", now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")]):
                            ok_any = True
                if ok_any:
                    notify_expense(sel_exp_jid, operator, sel_exp_client, exp_dict, total_fee)
                    st.success(f"已提交雜費 合計 ${total_fee:.0f}，已通知WhatsApp VIP群")
                else:
                    st.warning("請至少輸入一項雜費")

if role=="driver":
    with tabs[1]:
        st.subheader("🚚 我的任務 - 只顯示已搶、未完成")
        my_active = get_my_active_jobs(df_jobs)
        my_hist = get_my_history_jobs(df_jobs)
        st.caption(f"你好 {op_name}，而家有 {len(my_active)} 單未完成")
        if my_active.empty:
            st.info("暫無未完成任務，去「可接任務池」搶單")
        else:
            for _, row in my_active.sort_values(by="Date", ascending=True).iterrows():
                with st.container(border=True):
                    st.markdown(f"**{row.get('JobID','')}** | {row.get('Client','')} | {row.get('Status','')}")
                    st.caption(f"{row.get('From','')} -> {row.get('To','')} | 📅 {row.get('Date','')} {row.get('DeliveryTime','')}")
        show_hist = st.checkbox("☑️ 顯示已完成歷史 (Delivered)", value=False, key="v48_show_hist")
        if show_hist:
            if my_hist.empty:
                st.info("暫無歷史")
            else:
                st.dataframe(my_hist[["Date","JobID","Client","Status","From","To"]].sort_values(by="Date", ascending=False), use_container_width=True)
    with tabs[4]:
        st.subheader("📋 重複打印")
        if not df_jobs.empty:
            sel=st.selectbox("揀單打印", [format_job_option(r) for _,r in df_jobs.head(100).iterrows()])
            if sel:
                jid=sel.split(" | ")[0]
                job=df_jobs[df_jobs["JobID"]==jid].iloc[0].to_dict()
                st.components.v1.html(printable_order_html(job), height=750, scrolling=True)

cost_idx = 5 if role in ["admin","op"] else None
if cost_idx is not None:
    with tabs[cost_idx]:
        st.subheader("💵 每單費用 OrderCosts")
        if st.button("🔄 重新同步", key="v48_sync_cost_btn", type="primary"):
            sync_order_costs_sheet()
            st.rerun()
        df_cost=load_df("OrderCosts")
        if not df_cost.empty:
            st.dataframe(df_cost.sort_values(by="JobID",ascending=False), use_container_width=True, height=450)
    with tabs[6]:
        st.subheader("📥 下載POD - V48")
        df_pods = load_df("PODs")
        if df_pods.empty:
            st.warning("無PODs")
        else:
            df_pods["JobID"] = df_pods["JobID"].astype(str).str.strip()
            df_jobs["JobID"] = df_jobs["JobID"].astype(str).str.strip()
            df_merged = pd.merge(df_pods, df_jobs[["JobID","Date","Client","Driver","To","InternalNo"]] if not df_jobs.empty else pd.DataFrame(columns=["JobID"]), on="JobID", how="left")
            df_merged["CreatedAt_HKT_parsed"] = pd.to_datetime(df_merged["CreatedAt_HKT"], errors="coerce")
            c0, c1, c2 = st.columns(3)
            with c0:
                pod_job_ids = sorted(df_merged["JobID"].dropna().unique().tolist())
                job_options = ["全部 (唔 Filter)"] + [f"{j} | {df_merged[df_merged['JobID']==j].iloc[0].get('Client','')}" for j in pod_job_ids]
                sel_pod_job = st.selectbox("🎯 揀邊張單嘅 POD（留空 = 全部）", job_options, key="v48_pod_dl_job")
            with c1:
                start_date = st.date_input("開始日期", value=today_hk() - datetime.timedelta(days=7))
            with c2:
                end_date = st.date_input("結束日期", value=today_hk())
            c3, c4 = st.columns(2)
            with c3:
                driver_list = ["全部司機"] + sorted([d for d in df_merged["Driver"].dropna().unique().tolist() if str(d).strip()])
                sel_driver = st.selectbox("👷 揀司機", driver_list, key="v48_pod_dl_driver")
            with c4:
                only_with_pod = st.checkbox("只顯示有 POD 嘅單", value=True)
            filtered = df_merged.copy()
            if filtered["CreatedAt_HKT_parsed"].notna().any():
                filtered = filtered[(filtered["CreatedAt_HKT_parsed"].dt.date >= start_date) & (filtered["CreatedAt_HKT_parsed"].dt.date <= end_date)]
            if sel_pod_job != "全部 (唔 Filter)":
                target_jid = sel_pod_job.split(" | ")[0].strip()
                filtered = filtered[filtered["JobID"].astype(str).str.strip() == target_jid]
            if sel_driver != "全部司機":
                filtered = filtered[filtered["Driver"].astype(str).str.strip() == sel_driver]
            if only_with_pod:
                filtered = filtered[filtered["Base64_Image"].astype(str).str.len() > 10]
            st.success(f"找到 {len(filtered)} 張POD" + (f"（單：{sel_pod_job}）" if sel_pod_job != "全部 (唔 Filter)" else ""))
            if not filtered.empty:
                st.dataframe(filtered[["JobID","Client","Driver","Date","CreatedAt_HKT","FileName"]].head(50), use_container_width=True, height=250)
                if st.button(f"📦 打包下載 {len(filtered)} 張ZIP", type="primary", key="v48_pod_dl_zip_btn"):
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        for _, r in filtered.iterrows():
                            try:
                                raw = str(r.get("Base64_Image","")).strip()
                                if raw.startswith("http"):
                                    rr = requests.get(raw, timeout=20)
                                    img_bytes = rr.content
                                else:
                                    b64_str = raw
                                    if "," in b64_str:
                                        b64_str = b64_str.split(",",1)[1]
                                    b64_str += "=" * (-len(b64_str) % 4)
                                    img_bytes = base64.b64decode(b64_str)
                                zipf.writestr(f"{r.get('JobID','')}_{r.get('FileName','POD.jpg')}", img_bytes)
                            except:
                                pass
                    zip_buffer.seek(0)
                    dl_name = f"POD_{sel_pod_job.split(' | ')[0] if sel_pod_job!='全部 (唔 Filter)' else 'ALL'}_{start_date}_to_{end_date}.zip"
                    st.download_button("⬇️ 下載ZIP", data=zip_buffer, file_name=dl_name, mime="application/zip")
            else:
                st.info("呢個條件下冇 POD，試下放寬日期或者揀返「全部」")

if role in ["admin","op"]:
    with tabs[-1]:
        st.subheader("📦 通行 SCAN - 橫表")
        st.caption("通行 = 錢 · 運輸總表 = 箱數 · OCR 已停用，改做下面手動入數")
        try:
            load_scan_boxes_from_sheet()
        except Exception as e:
            print(f"tongxing load fail {e}")
        st.markdown("### ✍️ 手動入曲奇數 (取代 OCR)")
        st.caption("睇住張 SGL 相，逐間打箱數，冇貨就留 0。撳掣會直接寫入 運輸總表 + 通行，唔經 OCR，唔會再估錯。入錯咗可以再入多次，會覆蓋返同一日嗰行。")
        with st.form("manual_cookie_form_v48", clear_on_submit=False):
            manual_date = st.date_input("📅 送貨日期", value=today_hk() - datetime.timedelta(days=1), key="manual_cookie_date")
            ds_key = manual_date.strftime("%Y-%m-%d")
            existing_boxes = SCAN_BOXES.get(ds_key, {})
            if existing_boxes:
                st.info(f"呢日 {manual_date} 已經有數，上面會預設載入返，你改完再撳寫入就會覆蓋。")
            cols_m = st.columns(6)
            manual_boxes = {}
            for i, shop in enumerate(SCAN_SHOPS):
                with cols_m[i % 6]:
                    manual_boxes[shop] = st.number_input(shop, min_value=0, max_value=100, value=int(existing_boxes.get(shop, 0)), key=f"manual_v48_{shop}")
            manual_total_preview = sum(manual_boxes.values())
            manual_fee_preview = sum(scan_cost_for(s, v)[1] for s, v in manual_boxes.items() if v)
            st.info(f"預覽：總箱數 {manual_total_preview} 箱 | 通行 ${manual_fee_preview}")
            if st.form_submit_button("✅ 寫入 運輸總表 + 通行", type="primary", use_container_width=True):
                if manual_total_preview == 0:
                    st.warning("總箱數係 0，係咪未打數？如果你真係想清晒呢日嘅數，請確認一次。")
                else:
                    ok_m, tot_m, fee_m, sd_m = write_cookie_manual(manual_date, manual_boxes)
                    if ok_m:
                        st.success(f"🎉 已寫入 {sd_m}：總箱數 {tot_m} 箱，通行 ${fee_m}，已寫入 運輸總表 + 通行")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error(f"寫入失敗：{sd_m}")
        st.divider()
        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("合計箱數", f"{sum(sum(v.values()) for v in SCAN_BOXES.values())} 箱")
        with c2:
            st.metric("日數", f"{len(SCAN_BOXES)} 日")
        with c3:
            st.metric("有數日期", ", ".join(sorted(SCAN_BOXES.keys())[-3:]))
        st.error("⛔ 注意：舊版嗰個「一鍵寫入兩頁」已經停用，唔會再清 Sheet。想寫數請用上面「手動入曲奇數」。")
        if st.button("⚠️ 顯示點解停用 (唔會寫入)", key="tongxing_rebuild_disabled"):
            msg = rebuild_tongxing_sheet()
            st.warning(msg)
        st.markdown("**通行 - 錢**")
        matrix = build_tongxing_matrix()
        if len(matrix) > 1:
            df_show = pd.DataFrame(matrix[1:], columns=["日期"]+SCAN_SHOPS+["日總"])
            st.dataframe(df_show, use_container_width=True, height=300)
        else:
            st.info("暫無通行數據，請用上面手動入數，或者去 Sheet 睇")
        st.markdown("**運輸總表 - 箱數**")
        matrix2 = build_zongbiao_matrix()
        if len(matrix2) > 1:
            df_show2 = pd.DataFrame(matrix2[1:], columns=["日期"]+SCAN_SHOPS+["總箱數"])
            st.dataframe(df_show2, use_container_width=True, height=300)
        else:
            st.info("暫無運輸總表數據")
        excel_bytes = build_tongxing_excel_bytes()
        st.download_button("⬇️ 下載 Excel (通行 + 運輸總表 兩頁)", data=excel_bytes, file_name=f"每月物流成本節省記錄-{today_hk().strftime('%b%Y')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if st.button("📩 Send Excel 俾我自己 (兩頁，唔Send去Group)", key="tongxing_send_wa"):
            ok, info = send_excel_to_me(excel_bytes, f"每月物流成本節省記錄-{today_hk().strftime('%b%Y')}.xlsx", f"📋 每月物流成本節省記錄 {today_hk().strftime('%b%Y')} - 通行 + 運輸總表")
            st.success(f"已 Send Excel 俾你自己: {info[:200]}" if ok else f"失敗: {info} - 記得喺 Streamlit Secrets 加 MY_WHATSAPP")
