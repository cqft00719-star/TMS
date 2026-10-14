import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime
from datetime import timezone, timedelta
import uuid, io, base64, requests, urllib.parse

HKT = timezone(timedelta(hours=8))
def now_hk(): return datetime.datetime.now(HKT)
def today_hk(): return now_hk().date()
st.set_page_config(page_title="TMS V39 CallMeBot", layout="wide", page_icon="🚚")

FEE_MAP = {"油費":"隧道費","過路費":"登記費","維修":"交倉附加費","隧道":"隧道費"}
NEW_FEES = ["隧道費","登記費","停車費","交倉附加費","其他"]
UOM_OPTIONS = ["袋","箱","板","件","車","托","罐","樽","包","件","其他"]

def send_whatsapp_to_ops(message):
    try:
        wa_cfg = st.secrets.get("whatsapp", {})
        if not wa_cfg.get("enabled", False):
            return False, "未啟用"
        phones = wa_cfg.get("callmebot_phones", ["85260202131"])
        keys = wa_cfg.get("callmebot_apikeys", ["4480785"])
        ok=0
        logs=[]
        for i,ph in enumerate(phones):
            key = keys[i] if i < len(keys) else keys[0]
            url = f"https://api.callmebot.com/whatsapp.php?phone={ph}&text={urllib.parse.quote(message)}&apikey={key}"
            try:
                r = requests.get(url, timeout=15)
                logs.append(r.text[:100])
                if r.status_code==200: ok+=1
            except Exception as e:
                logs.append(str(e))
        return True, f"已發 {ok}/{len(phones)}"
    except Exception as e:
        return False, str(e)

def notify_new_order(job):
    msg = f"🆕新Order {now_hk().strftime('%m/%d %H:%M')}\nJobID:{job.get('JobID')}\nClient:{job.get('Client')} Internal:{job.get('InternalNo')}\nFrom:{job.get('From')} {job.get('From_Contact')} {job.get('From_Phone')} 提貨:{job.get('PickupTime')}\nTo:{job.get('To')} {job.get('To_Contact')} {job.get('To_Phone')} 送貨:{job.get('DeliveryTime')}\n貨:{job.get('CargoDetail')}"
    return send_whatsapp_to_ops(msg)
def notify_status(jid,driver,status,client=""):
    return send_whatsapp_to_ops(f"🔄{now_hk().strftime('%m/%d %H:%M')}\nJobID:{jid}\nClient:{client}\n司機:{driver} 已{status}")
def notify_pod(jid,driver,client="",cnt=1):
    return send_whatsapp_to_ops(f"✅{now_hk().strftime('%m/%d %H:%M')}\nJobID:{jid}\nClient:{client}\n司機:{driver} 已上傳{cnt}張POD Delivered")

@st.cache_resource
def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    return gspread.authorize(creds)
def get_sh():
    return get_client().open_by_key(st.secrets["connections"]["gsheets"]["spreadsheet"])

@st.cache_data(ttl=30)
def load_df(tab):
    try:
        ws=get_sh().worksheet(tab)
        vals=ws.get_all_values()
        if len(vals)<1: return pd.DataFrame()
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
        df=pd.DataFrame(rows, columns=clean)
        return df
    except: return pd.DataFrame()

def append_rows(tab, rows):
    get_sh().worksheet(tab).append_rows(rows, value_input_option='USER_ENTERED')
    st.cache_data.clear()
def append_row(tab,row): append_rows(tab,[row])
def append_row_by_header(tab, data_dict):
    ws=get_sh().worksheet(tab)
    headers=[h.strip() for h in ws.row_values(1)]
    row=[""]*len(headers)
    for k,v in data_dict.items():
        if k in headers: row[headers.index(k)]=str(v)
    ws.append_rows([row], value_input_option='USER_ENTERED')
    st.cache_data.clear()

def ensure_sheets():
    sh=get_sh()
    existing=[ws.title for ws in sh.worksheets()]
    def add(t,h,d=None):
        if t not in existing:
            ws=sh.add_worksheet(title=t, rows=3000, cols=40)
            ws.update('A1',[h])
            if d: ws.append_rows(d, value_input_option='USER_ENTERED')
    add("Users",["username","password","name","role"],[["admin","admin123","Admin","admin"],["阿傑","123","阿傑","driver"],["John","John123","John","driver"],["Mo","Mo123","Mo","driver"],["op01","op01","OP01","op"]])
    add("Jobs",["JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","DeliveryTime","Driver","Status","ProductName","QTY","UOM","CargoDetail","InternalNo","Notes","CreatedBy","CreatedAt_HKT","POD_Link","PodTime_HKT","POD_Image"])
    add("PODs",["POD_ID","JobID","FileName","Size","Base64_Image","CreatedBy","CreatedAt_HKT","PodTime_HKT"])
    add("Expenses",["ExpenseID","Date","JobID","Driver","Type","Amount","Remark","CreatedAt_HKT"])
    add("OrderCosts",["JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","DeliveryTime","CargoDetail","InternalNo","Status","Driver","隧道費","登記費","停車費","交倉附加費","其他","TotalExpense","LastSync_HKT"])

def printable_order_html(job):
    return f"""<html><head><meta charset="utf-8"><style>@media print{{button{{display:none!important}}}} body{{font-family:Arial,'Microsoft JhengHei';padding:15px;background:white}} table{{border-collapse:collapse}}</style></head><body>
    <div style="border:2px solid #000;max-width:800px;margin:auto;padding:15px;background:white">
    <h2 style="text-align:center">🚚 運輸訂單 {job.get('JobID','')}</h2>
    <p style="text-align:center">Internal:{job.get('InternalNo','')} | {job.get('Date','')} {job.get('DeliveryTime','')}</p><hr>
    <table style="width:100%;font-size:14px"><tr><td style="border:1px solid #000;padding:6px;background:#eee">客戶</td><td style="border:1px solid #000;padding:6px">{job.get('Client','')}</td><td style="border:1px solid #000;padding:6px;background:#eee">Zone</td><td style="border:1px solid #000;padding:6px">{job.get('Zone','')}</td></tr></table>
    <h3>📦 起點</h3><table style="width:100%;font-size:14px"><tr><td style="border:1px solid #000;padding:6px;background:#eee;width:20%">From</td><td style="border:1px solid #000;padding:6px" colspan="3">{job.get('From','')}</td></tr><tr><td style="border:1px solid #000;padding:6px;background:#eee">聯絡</td><td style="border:1px solid #000;padding:6px">{job.get('From_Contact','')}</td><td style="border:1px solid #000;padding:6px;background:#eee">電話</td><td style="border:1px solid #000;padding:6px">{job.get('From_Phone','')}</td></tr><tr><td style="border:1px solid #000;padding:6px;background:#eee">提貨</td><td style="border:1px solid #000;padding:6px" colspan="3">{job.get('PickupTime','')}</td></tr></table>
    <h3>🏁 終點</h3><table style="width:100%;font-size:14px"><tr><td style="border:1px solid #000;padding:6px;background:#eee;width:20%">To</td><td style="border:1px solid #000;padding:6px" colspan="3">{job.get('To','')}</td></tr><tr><td style="border:1px solid #000;padding:6px;background:#eee">聯絡</td><td style="border:1px solid #000;padding:6px">{job.get('To_Contact','')}</td><td style="border:1px solid #000;padding:6px;background:#eee">電話</td><td style="border:1px solid #000;padding:6px">{job.get('To_Phone','')}</td></tr><tr><td style="border:1px solid #000;padding:6px;background:#eee">送貨</td><td style="border:1px solid #000;padding:6px" colspan="3">{job.get('DeliveryTime','')}</td></tr></table>
    <h3>📋 貨物</h3><table style="width:100%;font-size:14px"><tr><td style="border:1px solid #000;padding:6px;background:#eee;width:25%">明細</td><td style="border:1px solid #000;padding:6px"><pre style="margin:0;white-space:pre-wrap;font-family:inherit">{job.get('CargoDetail','')}</pre></td></tr><tr><td style="border:1px solid #000;padding:6px;background:#eee">備註</td><td style="border:1px solid #000;padding:6px">{job.get('Notes','')}</td></tr></table></div>
    <div style="text-align:center;margin-top:15px"><button onclick="window.print()" style="background:#ff0000;color:white;padding:12px 30px;border:none;border-radius:8px;font-size:16px;cursor:pointer">🖨️ 立即打印此單</button></div></body></html>"""

if "user" not in st.session_state: st.session_state.user=None
if not st.session_state.user:
    st.title("TMS V39 CallMeBot 登入")
    df_u=load_df("Users")
    u=st.text_input("用戶名"); p=st.text_input("密碼",type="password")
    if st.button("登入",type="primary",key="v39_login"):
        m=df_u[(df_u["username"].astype(str).str.strip()==u.strip()) & (df_u["password"].astype(str).str.strip()==p.strip())]
        if not m.empty: st.session_state.user=m.iloc[0].to_dict(); st.rerun()
        else: st.error("帳號錯")
    st.stop()

if "cargo_builder" not in st.session_state: st.session_state.cargo_builder=[]
if "last_job" not in st.session_state: st.session_state.last_job=None

ensure_sheets()
user=st.session_state.user
role=str(user.get("role","driver")).lower()
operator=user.get("username","")
st.sidebar.title(f"👤 {user.get('name')} ({role})")
wa_enabled = st.secrets.get("whatsapp", {}).get("enabled", False)
st.sidebar.caption(f"WhatsApp: {'🟢已啟用 85260202131' if wa_enabled else '🔴未啟用'}")
if st.sidebar.button("登出", key="v39_logout"): st.session_state.user=None; st.cache_data.clear(); st.rerun()
if st.sidebar.button("🔄 清除Cache", key="v39_clear"): st.cache_data.clear(); st.rerun()
with st.sidebar.expander("📱 測試WhatsApp"):
    test_msg=st.text_area("測試訊息", value="TMS測試 85260202131")
    if st.button("發送測試", key="v39_test_wa"):
        ok,info=send_whatsapp_to_ops(test_msg)
        if ok: st.success(info)
        else: st.error(info)

if role=="admin": tabs=st.tabs(["📊 Dashboard","📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費","💵 每單費用","👥 用戶","🧹 維護"])
elif role=="op": tabs=st.tabs(["📊 Dashboard","📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費","💵 每單費用","🧹 維護"])
else: tabs=st.tabs(["🟢 可接任務池","🚚 我的任務","📸 上傳POD","📋 重複打印"])

df_jobs=load_df("Jobs")

with tabs[0]:
    if role=="driver":
        st.subheader("🟢 可接任務池 - 卡片+接單制")
        if df_jobs.empty: st.info("暫無")
        else:
            df_pool=df_jobs[(df_jobs["Driver"].astype(str).str.strip()=="") & (df_jobs["Status"].astype(str).str.upper().isin(["PENDING",""]))]
            if df_pool.empty: st.info("無可接")
            else:
                for idx,row in df_pool.sort_values(by="JobID",ascending=False).iterrows():
                    jid=row.get("JobID","")
                    with st.container(border=True):
                        c1,c2=st.columns([4,1])
                        c1.markdown(f"**{jid}** | {row.get('Client','')} | {row.get('Date','')} {row.get('DeliveryTime','')} | Internal:{row.get('InternalNo','')}")
                        if c2.button(f"🟢 搶此單",key=f"v39_claim_{jid}_{idx}",type="primary",use_container_width=True):
                            ws=get_sh().worksheet("Jobs")
                            headers=[h.strip() for h in ws.row_values(1)]
                            cell=ws.find(jid)
                            if cell:
                                if "Driver" in headers: ws.update_cell(cell.row, headers.index("Driver")+1, operator)
                                if "Status" in headers: ws.update_cell(cell.row, headers.index("Status")+1, "Accepted")
                                st.cache_data.clear()
                                ok,info=notify_status(jid,operator,"Accepted",row.get('Client',''))
                                st.success(f"已搶 {jid} | WhatsApp:{info}")
                                st.rerun()
                        cA,cB=st.columns(2)
                        with cA: st.write(f"📦起點: {row.get('From','')} {row.get('From_Contact','')} {row.get('From_Phone','')} 提貨:{row.get('PickupTime','')}")
                        with cB: st.write(f"🏁終點: {row.get('To','')} {row.get('To_Contact','')} {row.get('To_Phone','')} 送貨:{row.get('DeliveryTime','')}")
                        st.code(row.get('CargoDetail',''))
                        with st.expander(f"👁️ 預覽/打印 {jid}"):
                            st.components.v1.html(printable_order_html(row.to_dict()),height=700,scrolling=True)
    else:
        st.subheader("📊 Dashboard")
        st.dataframe(df_jobs.head(50))

if role in ["admin","op"]:
    with tabs[1]:
        st.subheader("📋 任務管理")
        if not df_jobs.empty:
            st.dataframe(df_jobs.head(100),use_container_width=True)
            opts=[f"{r['JobID']} | {r.get('Client','')} | {r.get('From','')}->{r.get('To','')}" for _,r in df_jobs.sort_values(by="JobID",ascending=False).head(100).iterrows()]
            sel=st.selectbox("揀單打印",opts,key="v39_task_print")
            if sel:
                jid=sel.split(" | ")[0]
                job=df_jobs[df_jobs["JobID"]==jid].iloc[0].to_dict()
                st.components.v1.html(printable_order_html(job),height=850,scrolling=True)

if role in ["admin","op"]:
    with tabs[2]:
        st.subheader("📤 上傳Order + 自動WhatsApp Op")
        ZONES=["港島","九龍","新界","離島","機場","中港","澳門","內地","其他"]
        cb1,cb2,cb3=st.columns([3,1,1])
        cb_prod=cb1.text_input("貨名",placeholder="FPCQHKGBP003",key="v39_cb_prod")
        cb_qty=cb2.number_input("數量",min_value=1,value=40,step=1,key="v39_cb_qty")
        cb_uom=cb3.selectbox("UOM",UOM_OPTIONS,index=6,key="v39_cb_uom")
        if st.button("➕ 添加",key="v39_cb_add",type="primary"):
            if cb_prod.strip(): st.session_state.cargo_builder.append(f"{cb_prod.strip()} {int(cb_qty)}{cb_uom}")
        if st.session_state.cargo_builder: st.success(" | ".join(st.session_state.cargo_builder))
        if st.button("🗑️ 清空",key="v39_cb_clear2"): st.session_state.cargo_builder=[]
        with st.form("v39_form"):
            c1,c2,c3=st.columns([2,1,1])
            client=c1.text_input("客戶*"); internal_no=c2.text_input("Internal#"); zone=c3.selectbox("Zone*",ZONES,index=1)
            c_date,c_pick,c_del=st.columns(3)
            date=c_date.date_input("日期*",value=today_hk()); pickup_time=c_pick.text_input("提貨時間"); delivery_time=c_del.text_input("送貨時間*",placeholder="0900-1700")
            c_from,c_from_c,c_from_p=st.columns([3,1,1])
            from_loc=c_from.text_input("起點*"); from_contact=c_from_c.text_input("起點聯絡"); from_phone=c_from_p.text_input("起點電話")
            c_to,c_to_c,c_to_p=st.columns([3,1,1])
            to_loc=c_to.text_input("終點*"); to_contact=c_to_c.text_input("終點聯絡"); to_phone=c_to_p.text_input("終點電話")
            cargo_detail=st.text_area("CargoDetail*",value="\n".join(st.session_state.cargo_builder),height=120)
            notes=st.text_input("備註")
            sub=st.form_submit_button("✅ 建立 + 自動WhatsApp通知Op",type="primary")
            if sub:
                if client.strip() and from_loc.strip() and to_loc.strip() and cargo_detail.strip():
                    jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                    job_dict={"JobID":jid,"Date":str(date),"Client":client,"Zone":zone,"From":from_loc,"From_Contact":from_contact,"From_Phone":from_phone,"PickupTime":pickup_time,"To":to_loc,"To_Contact":to_contact,"To_Phone":to_phone,"DeliveryTime":delivery_time,"CargoDetail":cargo_detail,"InternalNo":internal_no,"Notes":notes}
                    append_row_by_header("Jobs",{"JobID":jid,"Date":str(date),"Client":client,"Zone":zone,"From":from_loc,"From_Contact":from_contact,"From_Phone":from_phone,"PickupTime":pickup_time,"To":to_loc,"To_Contact":to_contact,"To_Phone":to_phone,"DeliveryTime":delivery_time,"Driver":"","Status":"Pending","CargoDetail":cargo_detail,"InternalNo":internal_no,"Notes":notes,"CreatedBy":operator,"CreatedAt_HKT":now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")})
                    ok,info=notify_new_order(job_dict)
                    st.session_state.last_job=job_dict
                    st.session_state.cargo_builder=[]
                    st.success(f"已建立 {jid} | WhatsApp:{info}")
                    st.rerun()
        if st.session_state.last_job:
            st.components.v1.html(printable_order_html(st.session_state.last_job),height=850,scrolling=True)

pod_idx=3 if role in ["admin","op"] else 2
with tabs[pod_idx]:
    st.subheader("📸 上傳POD + 自動通知Op")
    my_pool=df_jobs if role in ["admin","op"] else df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else df_jobs
    my_pool=my_pool.sort_values(by="JobID",ascending=False) if not my_pool.empty else my_pool
    if not my_pool.empty:
        jid=st.selectbox("選擇JobID",my_pool["JobID"].tolist(),key="v39_pod_sel")
        sel_info=df_jobs[df_jobs["JobID"]==jid].iloc[0].to_dict() if not df_jobs[df_jobs["JobID"]==jid].empty else {}
        st.info(f"Internal:{sel_info.get('InternalNo','')} | {str(sel_info.get('CargoDetail',''))[:100]}")
        pod_files=st.file_uploader("拍照",type=["jpg","png","jpeg"],accept_multiple_files=True,key="v39_pod_files")
        if pod_files and st.button(f"上傳 {len(pod_files)} 張 + 通知Op",type="primary",key="v39_up_pod"):
            from PIL import Image
            for idx_file, pf in enumerate(pod_files):
                try:
                    img=Image.open(io.BytesIO(pf.getvalue())).convert("RGB")
                    img.thumbnail((500,500))
                    buf=io.BytesIO(); img.save(buf,format="JPEG",quality=40,optimize=True)
                    fb=buf.getvalue(); b64=base64.b64encode(fb).decode(); data_url=f"data:image/jpeg;base64,{b64}"
                    fname=f"{jid}_POD_{now_hk().strftime('%Y%m%d_%H%M%S')}_{idx_file+1}.jpg"
                    pod_id=f"POD-{now_hk().strftime('%m%d%H%M%S')}-{uuid.uuid4().hex[:4]}"
                    append_row("PODs",[pod_id,jid,fname,str(len(fb)),data_url,operator,now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"),now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")])
                    ws=get_sh().worksheet("Jobs"); headers=[h.strip() for h in ws.row_values(1)]; cell=ws.find(jid)
                    if cell:
                        if "Status" in headers: ws.update_cell(cell.row, headers.index("Status")+1, "Delivered")
                        if "PodTime_HKT" in headers: ws.update_cell(cell.row, headers.index("PodTime_HKT")+1, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"))
                except Exception as e: st.error(f"{e}")
            ok,info=notify_pod(jid,operator,sel_info.get('Client',''),len(pod_files))
            st.cache_data.clear()
            st.success(f"已上傳 {len(pod_files)} 張 | WhatsApp:{info}")
