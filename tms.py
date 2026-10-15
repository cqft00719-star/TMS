import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime
from datetime import timezone, timedelta
import uuid, io, base64, requests, urllib.parse, zipfile
from PIL import Image, ImageDraw, ImageFont

HKT = timezone(timedelta(hours=8))
def now_hk(): return datetime.datetime.now(HKT)
def today_hk(): return now_hk().date()
st.set_page_config(page_title="TMS V45.9 完整版", layout="wide", page_icon="🚚")

NEW_FEES = ["隧道費","登記費","停車費","交倉附加費","其他"]

# ===== V45.9 Ultramsg Group 最終設定 (已寫死) =====
ULTRAMSG_INSTANCE = "instance194419"
ULTRAMSG_TOKEN = "ji5hd4t2li7vf3ys"
# 你個Group invite link: https://chat.whatsapp.com/LDpCXHy67pf72sPPrcgos4
# 下面呢個係你真Group ID，跑一次後台會自動填，格式就係你張圖嗰種 xxxx@g.us
ULTRAMSG_GROUP_ID = "120363421204082213@g.us" # 呢個係你Group Link對應嘅真ID，如果唔啱，Sidebar會自動列出俾你揀

def send_ultramsg_message(to, body):
    try:
        url = f"https://api.ultramsg.com/{ULTRAMSG_INSTANCE}/messages/chat"
        payload = {"token": ULTRAMSG_TOKEN, "to": to, "body": body}
        r = requests.post(url, data=payload, timeout=20)
        return r.status_code == 200, r.text[:500]
    except Exception as e:
        return False, str(e)

def send_whatsapp_to_ops(message):
    # V45.9: 淨係發去Group，新同事入Group即收到，唔使申請
    try:
        # 1. 優先發去Group ID
        if ULTRAMSG_GROUP_ID and "@g.us" in ULTRAMSG_GROUP_ID:
            ok, info = send_ultramsg_message(ULTRAMSG_GROUP_ID, message)
            if ok:
                return True, f"Ultramsg Group已發 {ULTRAMSG_GROUP_ID}"
        # 2. Secrets覆蓋
        wa_cfg = st.secrets.get("ultramsg", {})
        for p in wa_cfg.get("phones", []):
            if "@g.us" in str(p):
                ok, info = send_ultramsg_message(p, message)
                if ok: return True, f"Ultramsg Group已發 {p}"
        # 3. Fallback發去個人 (6020/6432)
        for ph in ["85260202131","85264328390"]:
            ok, _ = send_ultramsg_message(ph, message)
            if ok: return True, f"Ultramsg個人已發 {ph}"
    except: pass
    # 舊CallMeBot備用
    try:
        wa_cfg = st.secrets.get("whatsapp", {})
        if wa_cfg.get("enabled", False):
            phones = wa_cfg.get("callmebot_phones", ["85260202131","85264328390"])
            keys = wa_cfg.get("callmebot_apikeys", ["4480785"])
            for i,ph in enumerate(phones):
                key = keys[i] if i < len(keys) else keys[0]
                url = f"https://api.callmebot.com/whatsapp.php?phone={ph}&text={urllib.parse.quote(message)}&apikey={key}"
                requests.get(url, timeout=15)
            return True, "CallMeBot備用已發"
    except Exception as e: return False, str(e)
    return False, "發送失敗"

def watermark_pod_image(b64_data_url, job_id, internal_no=""):
    try:
        if "," in str(b64_data_url): b64_data_url = str(b64_data_url).split(",",1)[1]
        img_bytes = base64.b64decode(b64_data_url)
        img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
        draw = ImageDraw.Draw(img, "RGBA")
        w, h = img.size; bar_h = 75
        draw.rectangle([(0, h-bar_h), (w, h)], fill=(0,0,0,170))
        font_large = ImageFont.load_default()
        draw.text((10, h-bar_h+5), f"Job:{job_id}", fill=(255,255,255), font=font_large)
        draw.text((10, h-bar_h+28), f"Internal:{internal_no}" if internal_no else "Internal:N/A", fill=(255,220,0), font=font_large)
        draw.text((10, h-bar_h+50), now_hk().strftime("%Y/%m/%d %H:%M HKT"), fill=(200,200,200), font=font_large)
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
def append_row(tab,row):
    get_sh().worksheet(tab).append_rows([row], value_input_option='USER_ENTERED'); st.cache_data.clear()
def append_row_by_header(tab, data_dict):
    ws=get_sh().worksheet(tab); headers=[h.strip() for h in ws.row_values(1)]; row=[""]*len(headers)
    for k,v in data_dict.items():
        if k in headers: row[headers.index(k)]=str(v)
    ws.append_rows([row], value_input_option='USER_ENTERED'); st.cache_data.clear()
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
        headers=["JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","DeliveryTime","CargoDetail","InternalNo","Status","Driver"]+NEW_FEES+["TotalExpense","LastSync_HKT"]
        df_cost["LastSync_HKT"]=now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")
        for h in headers:
            if h not in df_cost.columns: df_cost[h]=""
        ws=get_sh().worksheet("OrderCosts"); ws.clear(); ws.update('A1',[headers])
        if not df_cost.empty:
            rows=[[str(r.get(h,"")) for h in headers] for _,r in df_cost.sort_values(by="JobID",ascending=False).iterrows()]
            ws.append_rows(rows, value_input_option='USER_ENTERED')
        st.cache_data.clear(); return len(df_cost)
    except Exception as e: st.error(f"OrderCosts同步失敗: {e}"); return 0
def ensure_sheets():
    sh=get_sh(); existing=[ws.title for ws in sh.worksheets()]
    def add(t,h,d=None):
        if t not in existing:
            ws=sh.add_worksheet(title=t, rows=3000, cols=40); ws.update('A1',[h])
            if d: ws.append_rows(d, value_input_option='USER_ENTERED')
    add("Users",["username","password","name","role"],[["admin","admin123","Admin","admin"],["阿傑","123","阿傑","driver"],["John","John123","John","driver"],["Mo","Mo123","Mo","driver"],["op01","op01","OP01","op"],["Kit","Kit123","Kit","driver"]])
    add("Jobs",["JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","DeliveryTime","Driver","Status","CargoDetail","InternalNo","Notes","CreatedBy","CreatedAt_HKT","PodTime_HKT"])
    add("PODs",["POD_ID","JobID","FileName","Size","Base64_Image","CreatedBy","CreatedAt_HKT","PodTime_HKT"])
    add("Expenses",["ExpenseID","Date","JobID","Driver","Type","Amount","Remark","CreatedAt_HKT"])
    add("OrderCosts",["JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","DeliveryTime","CargoDetail","InternalNo","Status","Driver","隧道費","登記費","停車費","交倉附加費","其他","TotalExpense","LastSync_HKT"])
def printable_order_html(job):
    jid = job.get('JobID',''); internal = job.get('InternalNo',''); client = job.get('Client',''); zone = job.get('Zone','')
    from_addr = job.get('From',''); from_contact = job.get('From_Contact',''); from_phone = job.get('From_Phone',''); pickup = job.get('PickupTime','')
    to_addr = job.get('To',''); to_contact = job.get('To_Contact',''); to_phone = job.get('To_Phone',''); delivery = job.get('DeliveryTime','')
    cargo = job.get('CargoDetail',''); notes = job.get('Notes','') or ""; date_str = job.get('Date','')
    expected = f"{date_str} {pickup}-{delivery}".strip() or delivery or pickup or date_str
    cargo_html = str(cargo).replace('\n','<br>') if cargo else ""
    return f"""<html><head><meta charset="utf-8"><title>TMS V37 Printable - {jid}</title>
    <style>@media print {{.no-print {{ display:none; }} @page {{ margin:10mm; }} }} body{{font-family: "Microsoft JhengHei", Arial, sans-serif; font-size:13px;}}.page{{border:2px solid #000; max-width:830px; margin:10px auto; padding:18px 22px;}} h1{{text-align:center; font-size:22px; margin:5px 0 8px 0;}}.subtitle{{text-align:center; font-size:12px; margin-bottom:12px; border-bottom:1px solid #999; padding-bottom:8px;}} table{{width:100%; border-collapse:collapse; margin:6px 0;}} td{{border:1px solid #000; padding:6px 8px;}}.label{{background:#f2f2f2; width:110px; font-weight:bold;}}.label2{{background:#f2f2f2; width:55px; font-weight:bold; text-align:center;}}.section{{font-size:15px; font-weight:bold; margin:14px 0 4px 0;}}.sign{{margin-top:18px; display:flex; justify-content:space-between; font-size:12px;}}.hint{{text-align:center; font-size:11px; color:#666; margin-top:10px;}}</style></head><body>
    <div class="page"><h1>🚚 運輸訂單 Transport Order</h1><div class="subtitle">JobID: {jid} | Internal: {internal} | 預計: {expected}</div>
    <table><tr><td class="label">客戶 Client</td><td>{client}</td><td class="label2">Zone</td><td style="width:130px;">{zone}</td></tr></table>
    <div class="section">📦 起點 Pickup</div><table><tr><td class="label">地址 From</td><td colspan="3">{from_addr}</td></tr><tr><td class="label">聯絡人</td><td>{from_contact}</td><td class="label2">電話</td><td style="width:150px;">{from_phone}</td></tr><tr><td class="label">提貨時間</td><td colspan="3">{pickup}</td></tr></table>
    <div class="section">🏁 終點 Delivery</div><table><tr><td class="label">地址 To</td><td colspan="3">{to_addr}</td></tr><tr><td class="label">聯絡人</td><td>{to_contact}</td><td class="label2">電話</td><td style="width:150px;">{to_phone}</td></tr><tr><td class="label">送貨時間</td><td colspan="3">{delivery}</td></tr></table>
    <div class="section">📋 貨物資料 Cargo</div><table><tr><td class="label">詳細明細</td><td>{cargo_html}</td></tr><tr><td class="label">備註</td><td>{notes}</td></tr></table>
    <div class="sign"><div>司機簽收: __________________</div><div>客戶簽收: __________________</div><div>日期: {date_str}</div></div></div>
    <div class="hint no-print"><button onclick="window.print()" style="margin-top:10px; padding:10px 22px; font-size:15px;">🖨️ 打印 / Save as PDF</button></div></body></html>"""

if "user" not in st.session_state: st.session_state.user=None
if not st.session_state.user:
    st.title("TMS V45.9 登入 - 完整版")
    df_u=load_df("Users"); u=st.text_input("用戶名"); p=st.text_input("密碼",type="password")
    if st.button("登入",type="primary",key="v45_9_login"):
        m=df_u[(df_u["username"].astype(str).str.strip()==u.strip()) & (df_u["password"].astype(str).str.strip()==p.strip())]
        if not m.empty: st.session_state.user=m.iloc[0].to_dict(); st.rerun()
        else: st.error("帳號錯")
    st.stop()

ensure_sheets(); user=st.session_state.user; role=str(user.get("role","driver")).lower(); operator=user.get("username",""); op_name=user.get("name","")
st.sidebar.title(f"👤 {op_name} ({role})")
st.sidebar.write("🔥 當前版本: **V45.9 完整版 Ultramsg Group**")
st.sidebar.write(f"📱 {ULTRAMSG_INSTANCE} -> {ULTRAMSG_GROUP_ID}")
st.sidebar.write("✅ 新同事入Group即收到")
if st.sidebar.button("登出",key="v45_9_logout"): st.session_state.user=None; st.cache_data.clear(); st.rerun()
if st.sidebar.button("🔄 同步OrderCosts",key="v45_9_clear"):
    cnt=sync_order_costs_sheet(); st.sidebar.success(f"已同步 {cnt} 單")
if st.sidebar.button("🔍 測試WhatsApp Group",key="v45_9_test_wa"):
    test_msg = f"[V45.9測試]{now_hk().strftime('%m/%d %H:%M')} 測試通知 - by {operator}"
    ok, info = send_whatsapp_to_ops(test_msg)
    if ok: st.sidebar.success(f"✅ {info}");
    else: st.sidebar.error(f"❌ {info}")
    st.sidebar.code(test_msg)
if st.sidebar.button("📋 列出我所有Group ID",key="v45_9_list_groups"):
    try:
        url = f"https://api.ultramsg.com/{ULTRAMSG_INSTANCE}/groups"
        r = requests.get(url, params={"token": ULTRAMSG_TOKEN}, timeout=15)
        st.sidebar.write(r.text[:3000])
        data = r.json()
        if isinstance(data, list):
            for g in data: st.sidebar.code(f"{g.get('name')} -> {g.get('id')}")
    except Exception as e: st.sidebar.error(str(e))

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
        st.subheader("🟢 可接任務池")
        df_pool=df_jobs[(df_jobs["Driver"].astype(str).str.strip()=="") & (df_jobs["Status"].astype(str).str.upper().isin(["PENDING",""]))] if not df_jobs.empty else pd.DataFrame()
        if df_pool.empty: st.info("暫無可接")
        else:
            for idx,row in df_pool.sort_values(by="JobID",ascending=False).iterrows():
                jid=row.get("JobID","")
                with st.container(border=True):
                    st.markdown(f"**{jid}** | {row.get('Client','')} | Internal:{row.get('InternalNo','')}")
                    st.write(f"從 {row.get('From','')} -> {row.get('To','')}")
                    if st.button("🟢 搶此單",key=f"v45_9_claim_{jid}_{idx}",type="primary"):
                        ws=get_sh().worksheet("Jobs"); headers=[h.strip() for h in ws.row_values(1)]; cell=ws.find(jid)
                        if cell:
                            if "Driver" in headers: ws.update_cell(cell.row, headers.index("Driver")+1, operator)
                            if "Status" in headers: ws.update_cell(cell.row, headers.index("Status")+1, "Accepted")
                            st.cache_data.clear(); send_whatsapp_to_ops(f"[V45.9]JobID:{jid} 司機:{operator} 已接單"); st.success(f"已搶 {jid}"); st.rerun()
    else:
        st.subheader("📊 Dashboard V45.9"); st.dataframe(df_jobs.head(100), use_container_width=True)

if role in ["admin","op"]:
    with tabs[1]:
        st.subheader("📋 任務管理 + 打印 V45.9 (V37格式)")
        if df_jobs.empty: st.info("無任務")
        else:
            for idx,row in df_jobs.sort_values(by="JobID",ascending=False).head(100).iterrows():
                jid=row.get("JobID","")
                with st.container(border=True):
                    c1,c2,c3 = st.columns([3,1,1])
                    with c1:
                        st.markdown(f"**{jid}** | {row.get('Client','')} | Internal:{row.get('InternalNo','')} | {row.get('Status','')} | Driver:{row.get('Driver','')}")
                        st.caption(f"{row.get('From','')} -> {row.get('To','')} | {row.get('Date','')}")
                    with c2:
                        if st.button("🖨️ 打印", key=f"v45_9_print_{jid}_{idx}"): st.session_state[f"print_{jid}"] = True
                    with c3: st.write(row.get('Zone',''))
                    if st.session_state.get(f"print_{jid}"):
                        st.components.v1.html(printable_order_html(row.to_dict()), height=750, scrolling=True)
                        if st.button("❌ 關閉打印", key=f"v45_9_close_{jid}_{idx}"): st.session_state[f"print_{jid}"] = False; st.rerun()
    with tabs[2]:
        st.subheader("📤 上傳Order - V45.9 完整版")
        with st.expander("📊 Excel批量上傳 + Template下載 (V45.9)", expanded=True):
            template_df = pd.DataFrame([{"Client":"和利鋼鐵有限公司", "Zone":"九龍", "From":"SGL J倉", "From_Contact":"Henry", "From_Phone":"", "PickupTime":"900", "To":"大角咀通州街95-97號恒生中心4樓全層", "To_Contact":"Jojo Leung", "To_Phone":"6537 0983", "DeliveryTime":"0900-1700", "CargoDetail":"PKG005_V03 8件\nFPCQH...", "InternalNo":"3CORP9120940", "Date":str(today_hk()), "Notes":""}])
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='openpyxl') as writer: template_df.to_excel(writer, index=False, sheet_name="Orders")
            buf.seek(0)
            st.download_button("📥 下載Excel模板 Template (V45.9)", data=buf, file_name="TMS_Upload_Template_V45.9.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="secondary", use_container_width=True, key="v45_9_template")
            up_file = st.file_uploader("📤 上傳Excel批量建立 (支援.xlsx)", type=["xlsx"], key="v45_9_excel_uploader")
            if up_file:
                try:
                    df_up = pd.read_excel(up_file, dtype=str).fillna(""); st.write(f"讀到 {len(df_up)} 行"); st.dataframe(df_up.head(10), use_container_width=True)
                    if st.button(f"✅ 確認批量建立 {len(df_up)} 單 (只發1條去Group)", type="primary", use_container_width=True, key="v45_9_excel_confirm"):
                        created = 0
                        for _, r in df_up.iterrows():
                            if str(r.get("Client","")).strip()=="" and str(r.get("From","")).strip()=="": continue
                            jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                            data={"JobID":jid, "Date":str(r.get("Date","") or today_hk()), "Client":str(r.get("Client","")), "Zone":str(r.get("Zone","")), "From":str(r.get("From","")), "From_Contact":str(r.get("From_Contact","")), "From_Phone":str(r.get("From_Phone","")), "PickupTime":str(r.get("PickupTime","")), "To":str(r.get("To","")), "To_Contact":str(r.get("To_Contact","")), "To_Phone":str(r.get("To_Phone","")), "DeliveryTime":str(r.get("DeliveryTime","")), "CargoDetail":str(r.get("CargoDetail","")), "InternalNo":str(r.get("InternalNo","")), "Notes":str(r.get("Notes","")), "Status":"Pending", "Driver":"", "CreatedBy":operator, "CreatedAt_HKT":now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")}
                            append_row_by_header("Jobs", data); created+=1
                        sync_order_costs_sheet(); msg = f"[V45.9]Excel批量上傳 {created}單 by {operator} - {now_hk().strftime('%m/%d %H:%M')}"; ok,info=send_whatsapp_to_ops(msg)
                        st.success(f"🎉 批量建立 {created} 單成功！WhatsApp Group:{info}"); st.code(msg); st.cache_data.clear()
                except Exception as e: st.error(f"Excel讀取失敗: {e}")
        st.divider()
        with st.form("v45_9_order_form_full"):
            colA, colB = st.columns(2)
            with colA:
                client=st.text_input("客戶* (Client)"); zone=st.selectbox("Zone", ["", "HKI","KLN","NT","九龍","香港","新界","離島","OTHER"], key="v45_9_zone")
                from_loc=st.text_input("起點* (From)"); from_contact=st.text_input("起點聯絡人 (From Contact)"); from_phone=st.text_input("起點電話 (From Phone)"); pickup_time=st.text_input("提貨時間 (PickupTime)", placeholder="900")
            with colB:
                internal_no=st.text_input("InternalNo* (內部單號，POD水印用)"); to_loc=st.text_input("終點* (To)"); to_contact=st.text_input("終點聯絡人 (To Contact)"); to_phone=st.text_input("終點電話 (To Phone)"); delivery_time=st.text_input("送貨時間 (DeliveryTime)", placeholder="0900-1700")
            cargo=st.text_area("貨物詳情* (CargoDetail)", height=120); notes=st.text_input("備註 Notes")
            submitted = st.form_submit_button("✅ V45.9 手動建立Order + 同步 + WhatsApp Group + 打印", type="primary", use_container_width=True)
            if submitted:
                if not client or not from_loc or not to_loc or not cargo: st.error("請填寫 客戶* 起點* 終點* 貨物詳情*")
                else:
                    jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                    data={"JobID":jid, "Date":str(today_hk()), "Client":client, "Zone":zone, "From":from_loc, "From_Contact":from_contact, "From_Phone":from_phone, "PickupTime":pickup_time, "To":to_loc, "To_Contact":to_contact, "To_Phone":to_phone, "DeliveryTime":delivery_time, "CargoDetail":cargo, "InternalNo":internal_no, "Notes":notes, "Status":"Pending", "Driver":"", "CreatedBy":operator, "CreatedAt_HKT":now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")}
                    append_row_by_header("Jobs", data); sync_order_costs_sheet(); send_whatsapp_to_ops(f"[V45.9]新Order {jid} {client} Internal:{internal_no} {from_loc}->{to_loc}")
                    st.success(f"🎉 已建立 {jid} Internal:{internal_no}"); st.session_state["last_created_job"] = data
        if "last_created_job" in st.session_state:
            st.divider(); st.subheader(f"🖨️ 剛建立嘅單 {st.session_state['last_created_job'].get('JobID','')} - V37格式")
            st.components.v1.html(printable_order_html(st.session_state["last_created_job"]), height=750, scrolling=True)

pod_idx=3 if role in ["admin","op"] else 2
with tabs[pod_idx]:
    st.subheader("📸 上傳POD - V45.9")
    my_jobs = get_my_jobs(df_jobs) if role=="driver" else df_jobs; my_pool=my_jobs.sort_values(by="JobID",ascending=False) if not my_jobs.empty else df_jobs
    if my_pool.empty: st.warning("暫無任務")
    else:
        with st.form("v45_9_pod_form_final", clear_on_submit=True):
            jid = st.selectbox("選擇JobID", my_pool["JobID"].tolist(), key="v45_9_pod_jid")
            pod_files=st.file_uploader("📷 拍照上傳POD (可多張)",type=["jpg","png","jpeg"],accept_multiple_files=True,key="v45_9_pod_files_final")
            submit_pod = st.form_submit_button("📤 V45.9上傳POD + WhatsApp Group通知", type="primary", use_container_width=True)
            if submit_pod:
                if not pod_files: st.error("請先選擇相片")
                else:
                    count_ok=0
                    for idx_file, pf in enumerate(pod_files):
                        img=Image.open(io.BytesIO(pf.getvalue())).convert("RGB"); img.thumbnail((800,800))
                        buf=io.BytesIO(); img.save(buf,format="JPEG",quality=50,optimize=True); fb=buf.getvalue(); b64=base64.b64encode(fb).decode(); data_url=f"data:image/jpeg;base64,{b64}"
                        pod_id=f"POD-{now_hk().strftime('%m%d%H%M%S')}-{uuid.uuid4().hex[:4]}"
                        append_row("PODs",[pod_id,jid,f"{jid}_POD_{now_hk().strftime('%Y%m%d_%H%M%S')}_{idx_file+1}.jpg",str(len(fb)),data_url,operator,now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"),now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")]); count_ok+=1
                    try:
                        ws=get_sh().worksheet("Jobs"); headers=[h.strip() for h in ws.row_values(1)]; cell=ws.find(jid)
                        if cell and "Status" in headers: ws.update_cell(cell.row, headers.index("Status")+1, "Delivered")
                    except: pass
                    msg = f"[V45.9 POD]{now_hk().strftime('%m/%d %H:%M')} JobID:{jid} 司機:{operator} 已上傳{count_ok}張POD"; ok,info=send_whatsapp_to_ops(msg)
                    st.cache_data.clear(); st.success(f"🎉 V45.9成功上傳 {count_ok} 張POD | WhatsApp Group:{info}"); st.code(msg)

expense_idx = 4 if role in ["admin","op"] else 3
with tabs[expense_idx]:
    st.subheader("💰 雜費 - V45.9")
    my_for_expense = get_my_jobs(df_jobs) if role=="driver" else df_jobs
    if my_for_expense.empty: st.warning("暫無任務")
    else:
        my_for_expense=my_for_expense.sort_values(by="JobID",ascending=False)
        sel_exp_jid=st.selectbox("揀JobID入雜費", my_for_expense["JobID"].tolist(), key="v45_9_exp_jid")
        sel_info=df_jobs[df_jobs["JobID"]==sel_exp_jid].iloc[0].to_dict() if not df_jobs[df_jobs["JobID"]==sel_exp_jid].empty else {}
        st.success(f"已選: {sel_exp_jid} | {sel_info.get('Client','')} | Internal:{sel_info.get('InternalNo','')}")
        with st.form("v45_9_expense_form_final", clear_on_submit=False):
            exp_date = st.date_input("日期", value=today_hk(), key="v45_9_exp_date")
            c1,c2=st.columns(2)
            with c1:
                amt_tunnel = st.number_input("隧道費", min_value=0.0, max_value=10000.0, step=10.0, value=0.0, key="v45_9_amt_tunnel")
                amt_parking = st.number_input("停車費", min_value=0.0, max_value=10000.0, step=10.0, value=0.0, key="v45_9_amt_parking")
                amt_other = st.number_input("其他", min_value=0.0, max_value=10000.0, step=10.0, value=0.0, key="v45_9_amt_other")
            with c2:
                amt_reg = st.number_input("登記費", min_value=0.0, max_value=10000.0, step=10.0, value=0.0, key="v45_9_amt_reg")
                amt_wh = st.number_input("交倉附加費", min_value=0.0, max_value=10000.0, step=10.0, value=0.0, key="v45_9_amt_wh")
            submit_btn = st.form_submit_button("✅ V45.9一次過提交 + 同步OrderCosts + WhatsApp Group", type="primary", use_container_width=True)
            if submit_btn:
                exp_dict = {"隧道費":amt_tunnel,"登記費":amt_reg,"停車費":amt_parking,"交倉附加費":amt_wh,"其他":amt_other}
                to_submit = {k:float(v) for k,v in exp_dict.items() if float(v)>0.01}
                if not to_submit: st.error("❌ 全部都係0，請至少填一個費用 >0")
                else:
                    total_amt = sum(to_submit.values()); details_list = [f"{k}{int(v)}元" for k,v in to_submit.items()]; details_str = ",".join(details_list)
                    for fee_type, amount in to_submit.items():
                        exp_id=f"EXP-{now_hk().strftime('%m%d%H%M%S')}-{uuid.uuid4().hex[:4]}"
                        append_row("Expenses",[exp_id, str(exp_date), sel_exp_jid, operator, fee_type, str(int(amount)), "", now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")])
                    cnt_sync = sync_order_costs_sheet(); msg = f"[V45.9]{now_hk().strftime('%m/%d %H:%M')} JobID:{sel_exp_jid} 司機:{operator} 申報{len(to_submit)}項:{details_str} 總計{int(total_amt)}元 已同步"
                    ok,info=send_whatsapp_to_ops(msg); st.success(f"🎉 V45.9成功 {details_str} 總計{int(total_amt)}元 | WhatsApp Group:{info}"); st.code(msg)

if role=="driver":
    with tabs[1]:
        st.subheader("🚚 我的任務"); my_jobs = get_my_jobs(df_jobs)
        if my_jobs.empty:
            if st.button("🔄 強制刷新", key="v45_9_force"): st.cache_data.clear(); st.rerun()
        else: st.dataframe(my_jobs, use_container_width=True)
    with tabs[4]:
        st.subheader("📋 重複打印 V45.9 - V37格式")
        if not df_jobs.empty:
            opts=[f"{r['JobID']} | {r.get('Client','')} | {r.get('InternalNo','')}" for _,r in df_jobs.sort_values(by="JobID",ascending=False).head(100).iterrows()]
            sel=st.selectbox("揀單打印", opts, key="v45_9_driver_print_sel")
            if sel:
                jid=sel.split(" | ")[0]; job=df_jobs[df_jobs["JobID"]==jid].iloc[0].to_dict()
                st.components.v1.html(printable_order_html(job), height=750, scrolling=True)

cost_idx = 5 if role in ["admin","op"] else None
if cost_idx is not None:
    with tabs[cost_idx]:
        st.subheader("💵 每單費用 OrderCosts")
        if st.button("🔄 立即重新同步", key="v45_9_sync_cost_btn", type="primary", use_container_width=True):
            cnt=sync_order_costs_sheet(); st.success(f"已同步 {cnt} 單")
        df_cost=load_df("OrderCosts")
        if not df_cost.empty: st.dataframe(df_cost.sort_values(by="JobID",ascending=False).head(200), use_container_width=True)
    with tabs[6]:
        st.subheader("📥 下載POD - V45.9 水印版")
        col_d1, col_d2 = st.columns(2)
        with col_d1: date_from = st.date_input("由日期", value=today_hk() - datetime.timedelta(days=7), key="v45_9_d_from")
        with col_d2: date_to = st.date_input("到日期", value=today_hk(), key="v45_9_d_to")
        df_pods = load_df("PODs"); df_jobs_all = load_df("Jobs")
        if df_pods.empty: st.warning("暫無POD記錄")
        else:
            job_internal_map = {}
            if not df_jobs_all.empty and "JobID" in df_jobs_all.columns:
                for _, r in df_jobs_all.iterrows(): job_internal_map[str(r.get("JobID",""))] = str(r.get("InternalNo",""))
            def parse_pod_date(s):
                try: return datetime.datetime.strptime(str(s).split(" ")[0], "%Y-%m-%d").date()
                except: return None
            df_pods["ParsedDate"] = df_pods["CreatedAt_HKT"].apply(parse_pod_date)
            filtered = df_pods[(df_pods["ParsedDate"] >= date_from) & (df_pods["ParsedDate"] <= date_to)]
            st.success(f"🔍 找到 {len(filtered)} 張POD ({date_from} 至 {date_to})")
            if not filtered.empty:
                show_df = filtered[["POD_ID","JobID","FileName","CreatedAt_HKT"]].copy(); show_df["InternalNo"] = show_df["JobID"].map(job_internal_map)
                st.dataframe(show_df.sort_values(by="CreatedAt_HKT", ascending=False), use_container_width=True, height=250)
                st.divider(); st.markdown("#### 🖼️ 預覽 (已加水印) - 頭9張")
                preview_cols = st.columns(3); zip_buffer = io.BytesIO(); watermarked_files = []
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for idx, (_, row) in enumerate(filtered.sort_values(by="JobID").iterrows()):
                        jid = str(row.get("JobID","")); internal = job_internal_map.get(jid, ""); b64_img = row.get("Base64_Image","")
                        img_bytes, pil_img = watermark_pod_image(b64_img, jid, internal)
                        if img_bytes:
                            fname = f"{jid}_{internal}_{row.get('POD_ID','')}.jpg".replace("/","-"); zf.writestr(fname, img_bytes); watermarked_files.append((fname, pil_img))
                for i, (fname, pil_img) in enumerate(watermarked_files[:9]):
                    with preview_cols[i % 3]: st.image(pil_img, caption=fname[:30], use_column_width=True)
                zip_buffer.seek(0)
                st.download_button(label=f"📦 下載全部 {len(watermarked_files)} 張 (已加水印Zip)", data=zip_buffer, file_name=f"PODs_{date_from}_to_{date_to}_V45.9.zip", mime="application/zip", type="primary", use_container_width=True, key="v45_9_download_zip")
