import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime
from datetime import timezone, timedelta
import uuid, io, base64, zipfile

HKT = timezone(timedelta(hours=8))
def now_hk(): return datetime.datetime.now(HKT)
def today_hk(): return now_hk().date()
st.set_page_config(page_title="TMS V34 Full Order + Print", layout="wide", page_icon="🚚")

FEE_MAP = {"油費":"隧道費","過路費":"登記費","維修":"交倉附加費","隧道":"隧道費","隧道費":"隧道費","登記費":"登記費","停車費":"停車費","交倉附加費":"交倉附加費","其他":"其他"}
NEW_FEES = ["隧道費","登記費","停車費","交倉附加費","其他"]
UOM_OPTIONS = ["袋","箱","板","件","車","托","其他"]

@st.cache_resource
def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    return gspread.authorize(creds)
def get_sh():
    client=get_client()
    return client.open_by_key(st.secrets["connections"]["gsheets"]["spreadsheet"])

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
            if clean.count(h)>0: h=f"{h}_{i}"
            clean.append(h)
        rows=[]
        for r in vals[1:]:
            if any(str(c).strip()!="" for c in r):
                if len(r)<len(clean): r=r+[""]*(len(clean)-len(r))
                rows.append(r[:len(clean)])
        df=pd.DataFrame(rows, columns=clean)
        if tab=="Expenses" and "Type" in df.columns:
            df["Type"]=df["Type"].astype(str).apply(lambda x: FEE_MAP.get(x.strip(), x.strip()) if x.strip() in FEE_MAP else x)
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
        if k in headers:
            row[headers.index(k)]=str(v)
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
    # V34 Jobs 加晒新欄
    add("Users",["username","password","name","role"],[["admin","admin123","Admin","admin"],["John","John123","John","driver"],["Mo","Mo123","Mo","driver"],["op01","op01","OP01","op"]])
    add("Jobs",["JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","Driver","Status","ProductName","QTY","UOM","CargoDetail","InternalNo","Notes","CreatedBy","CreatedAt_HKT","POD_Link","PodTime_HKT","POD_Image"])
    add("PODs",["POD_ID","JobID","FileName","Size","Base64_Image","CreatedBy","CreatedAt_HKT","PodTime_HKT"])
    add("Expenses",["ExpenseID","Date","JobID","Driver","Type","Amount","Remark","CreatedAt_HKT"])
    add("Config",["Key","Value"],[["POD_Folder_ID",""]])
    add("OrderCosts",["JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","ProductName","QTY","UOM","CargoDetail","InternalNo","Status","Driver","隧道費","登記費","停車費","交倉附加費","其他","TotalExpense","LastSync_HKT"])
    try:
        ws=sh.worksheet("Jobs")
        headers=[h.strip() for h in ws.row_values(1)]
        needed=["From_Contact","From_Phone","PickupTime","To_Contact","To_Phone","ProductName","CargoDetail","InternalNo"]
        if any(n not in headers for n in needed):
            # 重建header保持向下兼容
            new_h=["JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","Driver","Status","ProductName","QTY","UOM","CargoDetail","InternalNo","Notes","CreatedBy","CreatedAt_HKT","POD_Link","PodTime_HKT","POD_Image"]
            # 盡量保留舊數據位置
            all_vals=ws.get_all_values()
            old_data=all_vals[1:]
            ws.clear()
            ws.update('A1',[new_h])
            # 舊數據補空
            if old_data:
                fixed_rows=[]
                for r in old_data:
                    # 舊格式可能得10欄，補到新格式
                    d=dict(zip(headers, r))
                    new_r=[d.get(k,"") for k in new_h]
                    fixed_rows.append(new_r)
                ws.append_rows(fixed_rows, value_input_option='USER_ENTERED')
        ws2=sh.worksheet("OrderCosts")
        h2=[h.strip() for h in ws2.row_values(1)]
        if "From_Contact" not in h2:
            ws2.update('A1',[["JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","ProductName","QTY","UOM","CargoDetail","InternalNo","Status","Driver","隧道費","登記費","停車費","交倉附加費","其他","TotalExpense","LastSync_HKT"]])
    except: pass

def sync_order_costs_sheet():
    try:
        df_jobs=load_df("Jobs")
        df_exp=load_df("Expenses")
        if df_jobs.empty: return 0
        for c in ["Zone","UOM","From_Contact","From_Phone","PickupTime","To_Contact","To_Phone","ProductName","CargoDetail","InternalNo"]:
            if c not in df_jobs.columns: df_jobs[c]=""
        df_cost=df_jobs.copy()
        if not df_exp.empty and "JobID" in df_exp.columns and "Amount" in df_exp.columns:
            df_exp["Amount_num"]=pd.to_numeric(df_exp["Amount"], errors="coerce").fillna(0)
            exp_group=df_exp.groupby("JobID")["Amount_num"].sum().reset_index().rename(columns={"Amount_num":"TotalExpense"})
            df_cost=df_cost.merge(exp_group, on="JobID", how="left")
            if "Type" in df_exp.columns:
                pivot=df_exp.pivot_table(index="JobID", columns="Type", values="Amount_num", aggfunc="sum", fill_value=0).reset_index()
                df_cost=df_cost.merge(pivot, on="JobID", how="left")
        else:
            df_cost["TotalExpense"]=0
        df_cost["TotalExpense"]=df_cost["TotalExpense"].fillna(0)
        for t in NEW_FEES:
            if t not in df_cost.columns: df_cost[t]=0
            df_cost[t]=df_cost[t].fillna(0)
        headers=["JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","ProductName","QTY","UOM","CargoDetail","InternalNo","Status","Driver"]+NEW_FEES+["TotalExpense","LastSync_HKT"]
        df_out=df_cost.copy()
        df_out["LastSync_HKT"]=now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")
        sh=get_sh()
        ws=sh.worksheet("OrderCosts")
        ws.clear()
        ws.update('A1',[headers])
        if not df_out.empty:
            rows=[]
            for _, r in df_out.iterrows():
                row=[str(r.get(h,"")) for h in headers]
                rows.append(row)
            ws.append_rows(rows, value_input_option='USER_ENTERED')
        return len(df_out)
    except Exception as e:
        st.error(f"同步失敗: {e}")
        return 0

def find_pod_col(df):
    for col in df.columns:
        if df[col].astype(str).str.contains("data:image", na=False).any(): return col
    for c in df.columns:
        if "IMAGE" in c.upper(): return c
    return None
def show_table(df, max_rows=200):
    if df.empty: st.info("無資料"); return
    pod=find_pod_col(df)
    df2=df.drop(columns=[pod], errors="ignore") if pod else df
    if "JobID" in df2.columns:
        try: df2=df2.sort_values(by="JobID", ascending=False)
        except: pass
    st.markdown(f"<div style='overflow-x:auto'>{df2.head(max_rows).to_html(index=False, border=0)}</div>", unsafe_allow_html=True)

def printable_order_html(job):
    # 生成可打印HTML
    html = f"""
    <div id="print-area" style="font-family: Arial, sans-serif; padding:20px; border:2px solid #000; max-width:800px; background:white;">
    <h2 style="text-align:center; margin:0;">🚚 運輸訂單 Transport Order</h2>
    <p style="text-align:center;">JobID: {job.get('JobID','')} | Internal No: {job.get('InternalNo','')} | 預計送貨: {job.get('Date','')}</p>
    <hr>
    <table style="width:100%; border-collapse:collapse; font-size:14px;">
    <tr><td style="border:1px solid #000; padding:6px; width:20%; background:#eee;"><b>客戶 Client</b></td><td style="border:1px solid #000; padding:6px;">{job.get('Client','')}</td><td style="border:1px solid #000; padding:6px; width:20%; background:#eee;"><b>邊區 Zone</b></td><td style="border:1px solid #000; padding:6px;">{job.get('Zone','')}</td></tr>
    </table>
    <h3 style="margin:10px 0 5px;">📦 起點 Pickup</h3>
    <table style="width:100%; border-collapse:collapse; font-size:14px;">
    <tr><td style="border:1px solid #000; padding:6px; background:#eee; width:20%;">地址 From</td><td style="border:1px solid #000; padding:6px;" colspan="3">{job.get('From','')}</td></tr>
    <tr><td style="border:1px solid #000; padding:6px; background:#eee;">聯絡人</td><td style="border:1px solid #000; padding:6px;">{job.get('From_Contact','')}</td><td style="border:1px solid #000; padding:6px; background:#eee; width:20%;">電話</td><td style="border:1px solid #000; padding:6px;">{job.get('From_Phone','')}</td></tr>
    <tr><td style="border:1px solid #000; padding:6px; background:#eee;">提貨時間 PickupTime</td><td style="border:1px solid #000; padding:6px;" colspan="3">{job.get('PickupTime','')}</td></tr>
    </table>
    <h3 style="margin:10px 0 5px;">🏁 終點 Delivery</h3>
    <table style="width:100%; border-collapse:collapse; font-size:14px;">
    <tr><td style="border:1px solid #000; padding:6px; background:#eee; width:20%;">地址 To</td><td style="border:1px solid #000; padding:6px;" colspan="3">{job.get('To','')}</td></tr>
    <tr><td style="border:1px solid #000; padding:6px; background:#eee;">聯絡人</td><td style="border:1px solid #000; padding:6px;">{job.get('To_Contact','')}</td><td style="border:1px solid #000; padding:6px; background:#eee; width:20%;">電話</td><td style="border:1px solid #000; padding:6px;">{job.get('To_Phone','')}</td></tr>
    </table>
    <h3 style="margin:10px 0 5px;">📋 貨物資料 Cargo</h3>
    <table style="width:100%; border-collapse:collapse; font-size:14px;">
    <tr><td style="border:1px solid #000; padding:6px; background:#eee; width:20%;">貨名 Product</td><td style="border:1px solid #000; padding:6px;">{job.get('ProductName','')}</td><td style="border:1px solid #000; padding:6px; background:#eee; width:20%;">數量 QTY</td><td style="border:1px solid #000; padding:6px;">{job.get('QTY','')} {job.get('UOM','')}</td></tr>
    <tr><td style="border:1px solid #000; padding:6px; background:#eee;">詳細 Cargo Detail</td><td style="border:1px solid #000; padding:6px;" colspan="3"><pre style="margin:0; white-space:pre-wrap; font-family:inherit;">{job.get('CargoDetail','')}</pre></td></tr>
    <tr><td style="border:1px solid #000; padding:6px; background:#eee;">備註 Notes</td><td style="border:1px solid #000; padding:6px;" colspan="3">{job.get('Notes','')}</td></tr>
    </table>
    <br>
    <table style="width:100%; font-size:12px;">
    <tr><td>司機簽收: ________________</td><td>客戶簽收: ________________</td><td>日期: {job.get('Date','')}</td></tr>
    </table>
    </div>
    <script>
    function printDiv() {{
        var printContents = document.getElementById('print-area').innerHTML;
        var w = window.open('', '', 'height=800,width=900');
        w.document.write('<html><head><title>Order {job.get('JobID','')}</title></head><body>');
        w.document.write(printContents);
        w.document.write('</body></html>');
        w.document.close();
        w.print();
    }}
    </script>
    """
    return html

if "user" not in st.session_state: st.session_state.user=None
if not st.session_state.user:
    st.title("TMS V34 登入")
    df_u=load_df("Users")
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
if st.sidebar.button("登出"): st.session_state.user=None; st.cache_data.clear(); st.rerun()
if st.sidebar.button("🔄 清除Cache刷新"): st.cache_data.clear(); st.rerun()

if role=="admin": tabs=st.tabs(["📊 Dashboard","📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費","💵 每單費用","👥 用戶","🧹 維護"])
elif role=="op": tabs=st.tabs(["📊 Dashboard","📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費","💵 每單費用","🧹 維護"])
else: tabs=st.tabs(["🟢 可接任務池","🚚 我的任務","📸 上傳POD"])

df_jobs=load_df("Jobs")
df_pods=load_df("PODs")
df_exp=load_df("Expenses")
df_users=load_df("Users")

with tabs[0]:
    st.subheader("📊 Dashboard")
    if not df_jobs.empty:
        c1,c2,c3,c4=st.columns(4)
        c1.metric("總任務", len(df_jobs))
        if "Status" in df_jobs.columns: c2.metric("Delivered", len(df_jobs[df_jobs["Status"]=="Delivered"]))
        if "UOM" in df_jobs.columns: c3.metric("總QTY", pd.to_numeric(df_jobs["QTY"], errors='coerce').sum())
        if "InternalNo" in df_jobs.columns: c4.metric("有InternalNo", len(df_jobs[df_jobs["InternalNo"].astype(str).str.strip()!=""]))
    show_table(df_jobs.head(20))
    if st.button("🔄 同步到OrderCosts", type="primary"):
        cnt=sync_order_costs_sheet()
        st.success(f"已同步 {cnt}")

task_idx=1 if role in ["admin","op"] else None
if task_idx is not None:
    with tabs[task_idx]:
        st.subheader("📋 任務管理 - 可Print")
        if not df_jobs.empty:
            df_filtered=df_jobs.sort_values(by="JobID", ascending=False)
            show_table(df_filtered)
            st.divider()
            sel=st.selectbox("選JobID打印", df_jobs["JobID"].astype(str).tolist())
            if sel:
                job=df_jobs[df_jobs["JobID"]==sel].iloc[0].to_dict()
                html=printable_order_html(job)
                st.components.v1.html(html, height=800, scrolling=True)
                if st.button(f"🖨️ 打印 {sel}"):
                    st.markdown(html, unsafe_allow_html=True)
                    st.markdown('<button onclick="printDiv()" style="background:red;color:white;padding:10px 20px;border:none;border-radius:5px;">🖨️ 點擊打印</button>', unsafe_allow_html=True)

order_idx=2 if role in ["admin","op"] else None
if order_idx is not None:
    with tabs[order_idx]:
        st.subheader("📤 上傳Order - V34 起點終點聯絡人 + 貨名 + Internal No + Print")
        tab_manual, tab_excel = st.tabs(["✍️ 手動建立", "📊 Excel批量"])
        ZONES=["港島","九龍","新界","離島","機場","中港","澳門","內地","其他"]
        with tab_manual:
            with st.form("manual_v34", clear_on_submit=False):
                st.markdown("**客戶 & Internal & 預計送貨**")
                c1,c2,c3=st.columns([2,1,1])
                client=c1.text_input("客戶 Client*")
                internal_no=c2.text_input("Internal 單號# (你公司內部單)")
                zone=c3.selectbox("邊區Zone*", ZONES + ["自定"], index=1)
                zone_custom=st.text_input("自定邊區") if zone=="自定" else ""
                final_zone=zone_custom if zone=="自定" else zone
                date=st.date_input("📅 預計送貨日期*", value=today_hk())

                st.divider()
                st.markdown("**📦 起點 Pickup (新增聯絡人/電話/提貨時間)**")
                c_from,c_from_c,c_from_p,c_pick = st.columns([3,1,1,1])
                from_loc=c_from.text_input("起點地址 From*")
                from_contact=c_from_c.text_input("起點聯絡人")
                from_phone=c_from_p.text_input("起點電話")
                pickup_time=c_pick.text_input("提貨時間 (例 10:00-12:00)")

                st.markdown("**🏁 終點 Delivery (新增聯絡人/電話)**")
                c_to,c_to_c,c_to_p = st.columns([3,1,1])
                to_loc=c_to.text_input("終點地址 To*")
                to_contact=c_to_c.text_input("終點聯絡人")
                to_phone=c_to_p.text_input("終點電話")

                st.divider()
                st.markdown("**📋 貨物資料 (貨名 + QTY整數 + UOM + 詳細)**")
                c_prod,c_qty,c_uom = st.columns([2,1,1])
                product_name=c_prod.text_input("貨名 Product Name (例 食品/電子零件)")
                qty=c_qty.number_input("數量 QTY (整數)*", min_value=1, max_value=10000, value=1, step=1, format="%d")
                uom=c_uom.selectbox("單位 UOM*", UOM_OPTIONS, index=1)
                uom_custom=st.text_input("自定UOM") if uom=="其他" else ""
                final_uom=uom_custom if uom=="其他" and uom_custom.strip()!="" else uom
                cargo_detail=st.text_area("數量明細 (例: 產品A 10箱，產品B 1板) - 每行一個", placeholder="產品A 10箱\n產品B 1板\n產品C 5袋", height=100)
                notes=st.text_input("備註 Notes")

                submitted=st.form_submit_button("✅ 建立並可打印", type="primary")
                if submitted:
                    if client.strip() and from_loc.strip() and to_loc.strip():
                        jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                        append_row_by_header("Jobs",{"JobID":jid,"Date":str(date),"Client":client,"Zone":final_zone,"From":from_loc,"From_Contact":from_contact,"From_Phone":from_phone,"PickupTime":pickup_time,"To":to_loc,"To_Contact":to_contact,"To_Phone":to_phone,"Driver":"","Status":"Pending","ProductName":product_name,"QTY":str(int(qty)),"UOM":final_uom,"CargoDetail":cargo_detail,"InternalNo":internal_no,"Notes":notes,"CreatedBy":operator,"CreatedAt_HKT":now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")})
                        sync_order_costs_sheet()
                        st.success(f"✅ 已建立 {jid}")
                        # 立即顯示可打印
                        job_dict={"JobID":jid,"Date":str(date),"Client":client,"Zone":final_zone,"From":from_loc,"From_Contact":from_contact,"From_Phone":from_phone,"PickupTime":pickup_time,"To":to_loc,"To_Contact":to_contact,"To_Phone":to_phone,"ProductName":product_name,"QTY":str(int(qty)),"UOM":final_uom,"CargoDetail":cargo_detail,"InternalNo":internal_no,"Notes":notes}
                        html=printable_order_html(job_dict)
                        st.components.v1.html(html, height=900, scrolling=True)
                        st.markdown('<button onclick="printDiv()" style="background:#ff0000;color:white;padding:12px 24px;border:none;border-radius:8px;font-size:16px;">🖨️ 立即打印此Order</button>', unsafe_allow_html=True)
                        st.session_state["last_job"]=job_dict
                    else:
                        st.error("客戶、起點、終點必填")

            # 如果上次有建立，顯示打印
            if "last_job" in st.session_state:
                st.divider()
                st.subheader(f"🖨️ 上一張可打印: {st.session_state['last_job'].get('JobID')}")
                html=printable_order_html(st.session_state["last_job"])
                st.components.v1.html(html, height=800)

        with tab_excel:
            template_df=pd.DataFrame(columns=["Date(預計送貨)","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","ProductName","QTY","UOM","CargoDetail","InternalNo","Notes"])
            template_df.loc[0]=[str(today_hk()),"Sandy","九龍","HK Warehouse","John","91234567","10:00","KLN Shop","Mary","98765432","食品","10","箱","產品A 10箱\n產品B 1板","INT-001",""]
            buf_template=io.BytesIO(); template_df.to_excel(buf_template, index=False, engine='openpyxl'); buf_template.seek(0)
            st.download_button("📥 下載新Template V34 (含聯絡人+Internal+貨名)", data=buf_template, file_name="TMS_Order_Template_V34_Full.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
            excel_file=st.file_uploader("上傳Excel", type=["xlsx","xls"])
            if excel_file:
                try:
                    df_excel=pd.read_excel(excel_file)
                    st.dataframe(df_excel.head(20))
                    df_excel.columns=[c.strip() for c in df_excel.columns]
                    col_map={c.lower():c for c in df_excel.columns}
                    if st.button(f"✅ 批量建立 {len(df_excel)} 張", type="primary"):
                        created=[]
                        for _, row in df_excel.iterrows():
                            if str(row.get(col_map.get('client','Client'),"")).strip()=="": continue
                            try: q=int(float(row.get(col_map.get('qty','QTY'), 1)))
                            except: q=1
                            jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                            append_row_by_header("Jobs",{"JobID":jid,"Date":str(row.get(col_map.get('date','Date'), str(today_hk()))),"Client":str(row.get(col_map.get('client','Client'), "")),"Zone":str(row.get(col_map.get('zone','Zone'), "其他")),"From":str(row.get(col_map.get('from','From'), "")),"From_Contact":str(row.get(col_map.get('from_contact','From_Contact'), "")),"From_Phone":str(row.get(col_map.get('from_phone','From_Phone'), "")),"PickupTime":str(row.get(col_map.get('pickuptime','PickupTime'), "")),"To":str(row.get(col_map.get('to','To'), "")),"To_Contact":str(row.get(col_map.get('to_contact','To_Contact'), "")),"To_Phone":str(row.get(col_map.get('to_phone','To_Phone'), "")),"ProductName":str(row.get(col_map.get('productname','ProductName'), "")),"QTY":str(q),"UOM":str(row.get(col_map.get('uom','UOM'), "箱")),"CargoDetail":str(row.get(col_map.get('cargodetail','CargoDetail'), "")),"InternalNo":str(row.get(col_map.get('internalno','InternalNo'), "")),"Notes":str(row.get(col_map.get('notes','Notes'), "")),"Driver":"","Status":"Pending","CreatedBy":operator,"CreatedAt_HKT":now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")})
                            created.append(jid)
                        sync_order_costs_sheet()
                        st.success(f"已批量建立 {len(created)} 張: {', '.join(created[:5])}...")
                        st.rerun()
                except Exception as e: st.error(f"Excel失敗: {e}")

pod_idx=3 if role=="admin" else 3 if role=="op" else 2
with tabs[pod_idx]:
    st.subheader("📸 上傳POD")
    my_pool=df_jobs if role in ["admin","op"] else df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else df_jobs
    my_pool=my_pool.sort_values(by="JobID", ascending=False) if not my_pool.empty else my_pool
    if not my_pool.empty:
        jid=st.selectbox("選擇JobID", my_pool["JobID"].astype(str).tolist())
        sel_info=df_jobs[df_jobs["JobID"]==jid].iloc[0].to_dict() if not df_jobs[df_jobs["JobID"]==jid].empty else {}
        st.info(f"Internal: {sel_info.get('InternalNo','')} | 貨名: {sel_info.get('ProductName','')} | {sel_info.get('QTY','')} {sel_info.get('UOM','')} | 起點聯絡: {sel_info.get('From_Contact','')} {sel_info.get('From_Phone','')} | 終點聯絡: {sel_info.get('To_Contact','')} {sel_info.get('To_Phone','')}")
        pod_files=st.file_uploader("拍照 / 選圖", type=["jpg","png","jpeg"], accept_multiple_files=True)
        if pod_files and st.button(f"上傳 {len(pod_files)} 張到 {jid}", type="primary"):
            from PIL import Image
            import base64
            for idx_file, pod_file in enumerate(pod_files):
                try:
                    img=Image.open(io.BytesIO(pod_file.getvalue())).convert("RGB")
                    img.thumbnail((500,500))
                    buf=io.BytesIO(); img.save(buf, format="JPEG", quality=40, optimize=True)
                    fb=buf.getvalue()
                    if len(fb)>25000:
                        img.thumbnail((400,400))
                        buf=io.BytesIO(); img.save(buf, format="JPEG", quality=30, optimize=True)
                        fb=buf.getvalue()
                    b64=base64.b64encode(fb).decode()
                    data_url=f"data:image/jpeg;base64,{b64}"
                    if len(data_url)>=49000: continue
                    fname=f"{jid}_POD_{now_hk().strftime('%Y%m%d_%H%M%S')}_{idx_file+1}.jpg"
                    pod_id=f"POD-{now_hk().strftime('%m%d%H%M%S')}-{uuid.uuid4().hex[:4]}"
                    append_row("PODs",[pod_id, jid, fname, str(len(fb)), data_url, operator, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"), now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")])
                    ws=get_sh().worksheet("Jobs")
                    headers=[h.strip() for h in ws.row_values(1)]
                    cell=ws.find(jid)
                    if cell:
                        if "Status" in headers: ws.update_cell(cell.row, headers.index("Status")+1, "Delivered")
                        if "PodTime_HKT" in headers: ws.update_cell(cell.row, headers.index("PodTime_HKT")+1, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"))
                        if "POD_Image" in headers: ws.update_cell(cell.row, headers.index("POD_Image")+1, data_url)
                except Exception as e: st.error(f"{pod_file.name} 失敗 {e}")
            st.cache_data.clear(); st.success(f"成功上傳 {len(pod_files)} 張")

exp_idx=4 if role=="admin" else 4 if role=="op" else None
if exp_idx is not None:
    with tabs[exp_idx]:
        st.subheader("💰 雜費")
        show_table(df_exp)
        st.divider()
        with st.form("add_expense_v34", clear_on_submit=True):
            e_date=st.date_input("日期", value=today_hk())
            e_job=st.selectbox("關聯JobID", [""] + (df_jobs["JobID"].astype(str).tolist() if not df_jobs.empty else []))
            e_driver=st.text_input("司機", value=operator)
            e_type=st.selectbox("類型", NEW_FEES)
            e_amount=st.number_input("金額", min_value=0.0, step=1.0)
            e_remark=st.text_input("備註")
            if st.form_submit_button("新增雜費", type="primary"):
                eid=f"EXP-{now_hk().strftime('%m%d%H%M')}-{uuid.uuid4().hex[:4].upper()}"
                append_row("Expenses",[eid,str(e_date),e_job,e_driver,e_type,str(e_amount),e_remark,now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")])
                sync_order_costs_sheet()
                st.success(f"✅ {eid}"); st.rerun()

cost_idx=5 if role=="admin" else 5 if role=="op" else None
if cost_idx is not None:
    with tabs[cost_idx]:
        st.subheader("💵 每單費用 - V34 含聯絡人+Internal+貨名")
        if df_jobs.empty:
            st.info("無Jobs")
        else:
            df_cost=df_jobs.copy()
            for c in ["Zone","UOM","From_Contact","From_Phone","PickupTime","To_Contact","To_Phone","ProductName","CargoDetail","InternalNo"]:
                if c not in df_cost.columns: df_cost[c]=""
            if not df_exp.empty and "JobID" in df_exp.columns and "Amount" in df_exp.columns:
                df_exp["Amount_num"]=pd.to_numeric(df_exp["Amount"], errors="coerce").fillna(0)
                exp_group=df_exp.groupby("JobID")["Amount_num"].sum().reset_index().rename(columns={"Amount_num":"TotalExpense"})
                df_cost=df_cost.merge(exp_group, on="JobID", how="left")
                if "Type" in df_exp.columns:
                    pivot=df_exp.pivot_table(index="JobID", columns="Type", values="Amount_num", aggfunc="sum", fill_value=0).reset_index()
                    df_cost=df_cost.merge(pivot, on="JobID", how="left")
            else:
                df_cost["TotalExpense"]=0
            df_cost["TotalExpense"]=df_cost["TotalExpense"].fillna(0)
            for t in NEW_FEES:
                if t not in df_cost.columns: df_cost[t]=0
                df_cost[t]=df_cost[t].fillna(0)
            display_cols=["JobID","Date","InternalNo","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","ProductName","QTY","UOM","CargoDetail","Status","Driver"]+NEW_FEES+["TotalExpense"]
            display_cols=[c for c in display_cols if c in df_cost.columns]
            df_display=df_cost[display_cols].sort_values(by="JobID", ascending=False) if "JobID" in df_cost.columns else df_cost[display_cols]
            st.dataframe(df_display, use_container_width=True, height=500)
            col_sync,col_dl=st.columns(2)
            with col_sync:
                if st.button("🔄 同步到OrderCosts", type="primary", use_container_width=True):
                    cnt=sync_order_costs_sheet()
                    st.success(f"✅ 已同步 {cnt}")
            with col_dl:
                buf=io.BytesIO()
                df_display.to_excel(buf, index=False, engine='openpyxl')
                buf.seek(0)
                st.download_button("📥 下載Excel", data=buf, file_name=f"Cost_Per_Order_{now_hk().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)

user_idx=6 if role=="admin" else None
if user_idx is not None:
    with tabs[user_idx]:
        st.subheader("👥 用戶管理")
        df_show=df_users.copy()
        if "password" in df_show.columns: df_show["password"]="******"
        st.markdown(f"<div style='overflow-x:auto'>{df_show.to_html(index=False, border=0)}</div>", unsafe_allow_html=True)

maint_idx=7 if role=="admin" else 6 if role=="op" else None
if maint_idx is not None:
    with tabs[maint_idx]:
        st.subheader("🧹 維護")
        if st.button("🔄 同步到OrderCosts", type="primary", use_container_width=True):
            cnt=sync_order_costs_sheet()
            st.success(f"已同步 {cnt}")

if role=="driver":
    with tabs[0]:
        st.subheader("🟢 可接任務池")
        if not df_jobs.empty:
            mask=df_jobs["Driver"].astype(str).str.strip()==""
            avail=df_jobs[mask | (df_jobs["Status"].astype(str).str.lower()=="pending")]
            show_table(avail)
            for _, row in avail.iterrows():
                jid=row["JobID"]
                if st.button(f"搶 {jid}", key=f"claim_{jid}_v34", type="primary"):
                    ws=get_sh().worksheet("Jobs")
                    headers=[h.strip() for h in ws.row_values(1)]
                    cell=ws.find(jid)
                    if cell:
                        if "Driver" in headers: ws.update_cell(cell.row, headers.index("Driver")+1, operator)
                        if "Status" in headers: ws.update_cell(cell.row, headers.index("Status")+1, "Accepted")
                        st.cache_data.clear(); st.rerun()
    with tabs[1]:
        st.subheader("🚚 我的任務")
        my=df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else pd.DataFrame()
        show_table(my)
        if not my.empty:
            sel=st.selectbox("選我的任務打印", my["JobID"].astype(str).tolist())
            if sel:
                job=my[my["JobID"]==sel].iloc[0].to_dict()
                html=printable_order_html(job)
                st.components.v1.html(html, height=700)
