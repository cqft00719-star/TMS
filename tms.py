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
st.set_page_config(page_title="TMS V44 完整版", layout="wide", page_icon="🚚")

NEW_FEES = ["隧道費","登記費","停車費","交倉附加費","其他"]
UOM_OPTIONS = ["袋","箱","板","件","車","托","罐","樽","包"]

def send_whatsapp_to_ops(message):
    try:
        wa_cfg = st.secrets.get("whatsapp", {})
        if not wa_cfg.get("enabled", False): return False, "WhatsApp未啟用"
        phones = wa_cfg.get("callmebot_phones", ["85260202131"])
        keys = wa_cfg.get("callmebot_apikeys", ["4480785"])
        ok=0
        for i,ph in enumerate(phones):
            key = keys[i] if i < len(keys) else keys[0]
            url = f"https://api.callmebot.com/whatsapp.php?phone={ph}&text={urllib.parse.quote(message)}&apikey={key}"
            r = requests.get(url, timeout=20)
            if r.status_code==200: ok+=1
        return True, f"已發 {ok}/{len(phones)}"
    except Exception as e: return False, str(e)

@st.cache_resource
def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    return gspread.authorize(creds)
def get_sh(): return get_client().open_by_key(st.secrets["connections"]["gsheets"]["spreadsheet"])

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
        return pd.DataFrame(rows, columns=clean)
    except: return pd.DataFrame()

def append_row(tab,row):
    get_sh().worksheet(tab).append_rows([row], value_input_option='USER_ENTERED')
    st.cache_data.clear()

def append_row_by_header(tab, data_dict):
    ws=get_sh().worksheet(tab)
    headers=[h.strip() for h in ws.row_values(1)]
    row=[""]*len(headers)
    for k,v in data_dict.items():
        if k in headers: row[headers.index(k)]=str(v)
    ws.append_rows([row], value_input_option='USER_ENTERED')
    st.cache_data.clear()

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
                        fee_map=pivot[fee].to_dict()
                        df_cost[fee]=df_cost["JobID"].map(fee_map).fillna(0)
        headers=["JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","DeliveryTime","CargoDetail","InternalNo","Status","Driver"]+NEW_FEES+["TotalExpense","LastSync_HKT"]
        df_cost["LastSync_HKT"]=now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")
        for h in headers:
            if h not in df_cost.columns: df_cost[h]=""
        sh=get_sh()
        ws=sh.worksheet("OrderCosts")
        ws.clear()
        ws.update('A1',[headers])
        if not df_cost.empty:
            rows=[[str(r.get(h,"")) for h in headers] for _,r in df_cost.sort_values(by="JobID",ascending=False).iterrows()]
            ws.append_rows(rows, value_input_option='USER_ENTERED')
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
    add("Users",["username","password","name","role"],[["admin","admin123","Admin","admin"],["阿傑","123","阿傑","driver"],["John","John123","John","driver"],["Mo","Mo123","Mo","driver"],["op01","op01","OP01","op"],["Kit","Kit123","Kit","driver"]])
    add("Jobs",["JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","DeliveryTime","Driver","Status","CargoDetail","InternalNo","Notes","CreatedBy","CreatedAt_HKT","PodTime_HKT"])
    add("PODs",["POD_ID","JobID","FileName","Size","Base64_Image","CreatedBy","CreatedAt_HKT","PodTime_HKT"])
    add("Expenses",["ExpenseID","Date","JobID","Driver","Type","Amount","Remark","CreatedAt_HKT"])
    add("OrderCosts",["JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","DeliveryTime","CargoDetail","InternalNo","Status","Driver","隧道費","登記費","停車費","交倉附加費","其他","TotalExpense","LastSync_HKT"])

def printable_order_html(job):
    return f"""<html><head><meta charset="utf-8"><style>body{{font-family:Arial}} table{{border-collapse:collapse}}</style></head><body>
    <div style="border:2px solid #000;max-width:800px;margin:auto;padding:15px">
    <h2 style="text-align:center">🚚 運輸訂單 {job.get('JobID','')}</h2>
    <p>Internal: {job.get('InternalNo','')} | Client: {job.get('Client','')}</p>
    <p>From: {job.get('From','')} To: {job.get('To','')}</p>
    <pre style="white-space:pre-wrap;border:1px solid #000;padding:8px">{job.get('CargoDetail','')}</pre></div>
    <div style="text-align:center;margin-top:15px"><button onclick="window.print()">🖨️ 立即打印</button></div></body></html>"""

if "user" not in st.session_state: st.session_state.user=None
if not st.session_state.user:
    st.title("TMS V44 登入 - 終極修復版")
    df_u=load_df("Users")
    u=st.text_input("用戶名"); p=st.text_input("密碼",type="password")
    if st.button("登入",type="primary",key="v44_login"):
        m=df_u[(df_u["username"].astype(str).str.strip()==u.strip()) & (df_u["password"].astype(str).str.strip()==p.strip())]
        if not m.empty: st.session_state.user=m.iloc[0].to_dict(); st.rerun()
        else: st.error("帳號錯")
    st.stop()

ensure_sheets()
user=st.session_state.user
role=str(user.get("role","driver")).lower()
operator=user.get("username","")
op_name=user.get("name","")

st.sidebar.title(f"👤 {op_name} ({role})")
st.sidebar.write("🔥 當前版本: **V44 完整版**")
st.sidebar.write("2026/10/14 12:30 Form修復")
if st.sidebar.button("登出",key="v44_logout"): st.session_state.user=None; st.cache_data.clear(); st.rerun()
if st.sidebar.button("🔄 清除Cache + 同步OrderCosts",key="v44_clear"):
    cnt=sync_order_costs_sheet()
    st.sidebar.success(f"已同步 {cnt} 單")

if role=="admin": tabs=st.tabs(["📊 Dashboard","📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費","💵 每單費用"])
elif role=="op": tabs=st.tabs(["📊 Dashboard","📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費","💵 每單費用"])
else: tabs=st.tabs(["🟢 可接任務池","🚚 我的任務","📸 上傳POD","💰 雜費","📋 重複打印"])

df_jobs=load_df("Jobs")
df_exp=load_df("Expenses")

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
                    if st.button("🟢 搶此單",key=f"v44_claim_{jid}_{idx}",type="primary"):
                        ws=get_sh().worksheet("Jobs")
                        headers=[h.strip() for h in ws.row_values(1)]
                        cell=ws.find(jid)
                        if cell:
                            if "Driver" in headers: ws.update_cell(cell.row, headers.index("Driver")+1, operator)
                            if "Status" in headers: ws.update_cell(cell.row, headers.index("Status")+1, "Accepted")
                            st.cache_data.clear()
                            send_whatsapp_to_ops(f"JobID:{jid} 司機:{operator} 已接單")
                            st.success(f"已搶 {jid}"); st.rerun()
    else:
        st.subheader("📊 Dashboard V44")
        st.dataframe(df_jobs.head(100), use_container_width=True)
        if st.button("🔄 同步到OrderCosts", key="v44_sync_dash"):
            cnt=sync_order_costs_sheet()
            st.success(f"已同步 {cnt}")

if role in ["admin","op"]:
    with tabs[1]: st.subheader("📋 任務管理"); st.dataframe(df_jobs.head(200),use_container_width=True)
    with tabs[2]:
        st.subheader("📤 上傳Order")
        with st.form("v44_order_form"):
            client=st.text_input("客戶*"); from_loc=st.text_input("起點*"); to_loc=st.text_input("終點*"); cargo=st.text_area("貨*")
            if st.form_submit_button("建立"):
                jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                append_row_by_header("Jobs",{"JobID":jid,"Date":str(today_hk()),"Client":client,"From":from_loc,"To":to_loc,"CargoDetail":cargo,"Driver":"","Status":"Pending","CreatedBy":operator,"CreatedAt_HKT":now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")})
                sync_order_costs_sheet()
                send_whatsapp_to_ops(f"新Order {jid} {client}")
                st.success(f"已建立 {jid}"); st.rerun()

pod_idx=3 if role in ["admin","op"] else 2
with tabs[pod_idx]:
    st.subheader("📸 上傳POD")
    my_jobs = get_my_jobs(df_jobs) if role=="driver" else df_jobs
    my_pool=my_jobs.sort_values(by="JobID",ascending=False) if not my_jobs.empty else df_jobs
    if not my_pool.empty:
        jid=st.selectbox("選擇JobID",my_pool["JobID"].tolist(),key="v44_pod_sel")
        pod_files=st.file_uploader("拍照",type=["jpg","png","jpeg"],accept_multiple_files=True,key="v44_pod_files")
        if pod_files and st.button(f"上傳 {len(pod_files)} 張",type="primary",key="v44_up_pod"):
            from PIL import Image
            for idx_file, pf in enumerate(pod_files):
                img=Image.open(io.BytesIO(pf.getvalue())).convert("RGB"); img.thumbnail((500,500))
                buf=io.BytesIO(); img.save(buf,format="JPEG",quality=40,optimize=True)
                fb=buf.getvalue(); b64=base64.b64encode(fb).decode(); data_url=f"data:image/jpeg;base64,{b64}"
                pod_id=f"POD-{now_hk().strftime('%m%d%H%M%S')}-{uuid.uuid4().hex[:4]}"
                append_row("PODs",[pod_id,jid,f"{jid}_POD.jpg",str(len(fb)),data_url,operator,now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"),now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")])
            send_whatsapp_to_ops(f"JobID:{jid} 司機:{operator} 已上傳{len(pod_files)}張POD")
            st.cache_data.clear(); st.success("已上傳")

# ==================== V44 雜費終極修復 ====================
expense_idx = 4 if role in ["admin","op"] else 3
with tabs[expense_idx]:
    st.subheader("💰 雜費 - V44 Form終極修復版")
    st.info("✅ V44 2026/10/14 12:30 - 用Form一次過提交，0蚊自動過濾，WhatsApp數字準確")
    my_for_expense = get_my_jobs(df_jobs) if role=="driver" else df_jobs
    if my_for_expense.empty:
        st.warning("暫無任務")
    else:
        my_for_expense=my_for_expense.sort_values(by="JobID",ascending=False)
        sel_exp_jid=st.selectbox("揀JobID入雜費", my_for_expense["JobID"].tolist(), key="v44_exp_jid")
        sel_info=df_jobs[df_jobs["JobID"]==sel_exp_jid].iloc[0].to_dict() if not df_jobs[df_jobs["JobID"]==sel_exp_jid].empty else {}
        st.success(f"已選: {sel_exp_jid} | {sel_info.get('Client','')} | Internal:{sel_info.get('InternalNo','')}")

        with st.form("v44_expense_form_final", clear_on_submit=False):
            st.markdown("#### 📝 一次過填寫所有費用 (填0=不提交)")
            exp_date = st.date_input("日期", value=today_hk(), key="v44_exp_date")
            c1,c2=st.columns(2)
            with c1:
                amt_tunnel = st.number_input("隧道費 $", min_value=0.0, max_value=10000.0, step=10.0, value=0.0, key="v44_amt_tunnel")
                amt_parking = st.number_input("停車費 $", min_value=0.0, max_value=10000.0, step=10.0, value=0.0, key="v44_amt_parking")
                amt_other = st.number_input("其他 $", min_value=0.0, max_value=10000.0, step=10.0, value=0.0, key="v44_amt_other")
            with c2:
                amt_reg = st.number_input("登記費 $", min_value=0.0, max_value=10000.0, step=10.0, value=0.0, key="v44_amt_reg")
                amt_wh = st.number_input("交倉附加費 $", min_value=0.0, max_value=10000.0, step=10.0, value=0.0, key="v44_amt_wh")

            submit_btn = st.form_submit_button("✅ 一次過提交 + 同步OrderCosts + WhatsApp", type="primary", use_container_width=True)

            if submit_btn:
                exp_dict = {
                    "隧道費": amt_tunnel,
                    "登記費": amt_reg,
                    "停車費": amt_parking,
                    "交倉附加費": amt_wh,
                    "其他": amt_other,
                }
                # DEBUG
                st.write("🔍 DEBUG 你輸入:")
                st.json(exp_dict)

                to_submit = {}
                for k,v in exp_dict.items():
                    if float(v) > 0.01:
                        to_submit[k] = float(v)

                st.write(f"🔍 DEBUG 過濾後 (>0) 要提交: {to_submit}")

                if not to_submit:
                    st.error("❌ 全部都係0，請至少填一個費用 >0，例如停車費 100")
                else:
                    total_amt = sum(to_submit.values())
                    details_list = []
                    for k,v in to_submit.items():
                        details_list.append(f"{k} ${int(v)}")
                    details_str = ", ".join(details_list)

                    st.write(f"🔍 DEBUG 準備寫入: {details_str} 總計 ${int(total_amt)}")

                    # 寫入Expenses
                    for fee_type, amount in to_submit.items():
                        exp_id=f"EXP-{now_hk().strftime('%m%d%H%M%S')}-{uuid.uuid4().hex[:4]}"
                        append_row("Expenses",[exp_id, str(exp_date), sel_exp_jid, operator, fee_type, str(int(amount)), "", now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")])

                    cnt_sync = sync_order_costs_sheet()
                    msg = f"💰{now_hk().strftime('%m/%d %H:%M')} JobID:{sel_exp_jid} 司機:{operator} 申報 {len(to_submit)}項: {details_str} 總計 ${int(total_amt)} 已同步OrderCosts"
                    ok,info=send_whatsapp_to_ops(msg)
                    st.success(f"🎉 成功提交 {len(to_submit)}項!")
                    st.success(f"📋 {details_str} 總計 ${int(total_amt)}")
                    st.success(f"📱 WhatsApp: {info}")
                    st.code(msg)

        st.divider()
        st.markdown("**📋 我嘅雜費記錄 (最新20)**")
        if not df_exp.empty and sel_exp_jid in df_exp["JobID"].astype(str).values:
            my_exp = df_exp[df_exp["JobID"]==sel_exp_jid]
            st.dataframe(my_exp.sort_values(by="ExpenseID",ascending=False).head(20), use_container_width=True)
            total_job = pd.to_numeric(my_exp["Amount"], errors="coerce").sum()
            st.metric(f"{sel_exp_jid} 總雜費", f"${total_job:.0f}")

if role=="driver":
    with tabs[1]:
        st.subheader("🚚 我的任務")
        my_jobs = get_my_jobs(df_jobs)
        if my_jobs.empty:
            if st.button("🔄 強制刷新", key="v44_force"): st.cache_data.clear(); st.rerun()
        else:
            st.dataframe(my_jobs, use_container_width=True)
    with tabs[4]:
        st.subheader("📋 重複打印")
        if not df_jobs.empty:
            opts=[f"{r['JobID']} | {r.get('Client','')}" for _,r in df_jobs.sort_values(by="JobID",ascending=False).head(100).iterrows()]
            sel=st.selectbox("揀單打印", opts, key="v44_driver_print_sel")
            if sel:
                jid=sel.split(" | ")[0]
                job=df_jobs[df_jobs["JobID"]==jid].iloc[0].to_dict()
                st.components.v1.html(printable_order_html(job), height=800, scrolling=True)

cost_idx = 5 if role in ["admin","op"] else None
if cost_idx is not None:
    with tabs[cost_idx]:
        st.subheader("💵 每單費用 OrderCosts")
        if st.button("🔄 立即重新同步", key="v44_sync_cost_btn", type="primary", use_container_width=True):
            cnt=sync_order_costs_sheet()
            st.success(f"已同步 {cnt} 單")
        df_cost=load_df("OrderCosts")
        if not df_cost.empty:
            st.dataframe(df_cost.sort_values(by="JobID",ascending=False).head(200), use_container_width=True)
