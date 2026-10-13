import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime
from datetime import timezone, timedelta
import uuid, io, base64

HKT = timezone(timedelta(hours=8))
def now_hk(): return datetime.datetime.now(HKT)
def today_hk(): return now_hk().date()
st.set_page_config(page_title="TMS V36 Fixed Only CargoDetail", layout="wide", page_icon="🚚")

FEE_MAP = {"油費":"隧道費","過路費":"登記費","維修":"交倉附加費","隧道":"隧道費"}
NEW_FEES = ["隧道費","登記費","停車費","交倉附加費","其他"]
UOM_OPTIONS = ["袋","箱","板","件","車","托","罐","樽","包","其他"]

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
            if h in clean: h=f"{h}_{i}"
            clean.append(h)
        rows=[]
        for r in vals[1:]:
            if any(str(c).strip()!="" for c in r):
                if len(r)<len(clean): r=r+[""]*(len(clean)-len(r))
                rows.append(r[:len(clean)])
        df=pd.DataFrame(rows, columns=clean)
        if tab=="Expenses" and "Type" in df.columns:
            df["Type"]=df["Type"].astype(str).apply(lambda x: FEE_MAP.get(x.strip(), x) if x.strip() in FEE_MAP else x)
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
    add("Users",["username","password","name","role"],[["admin","admin123","Admin","admin"],["John","John123","John","driver"],["Mo","Mo123","Mo","driver"],["op01","op01","OP01","op"]])
    add("Jobs",["JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","DeliveryTime","Driver","Status","ProductName","QTY","UOM","CargoDetail","InternalNo","Notes","CreatedBy","CreatedAt_HKT","POD_Link","PodTime_HKT","POD_Image"])
    add("PODs",["POD_ID","JobID","FileName","Size","Base64_Image","CreatedBy","CreatedAt_HKT","PodTime_HKT"])
    add("Expenses",["ExpenseID","Date","JobID","Driver","Type","Amount","Remark","CreatedAt_HKT"])
    add("Config",["Key","Value"],[["POD_Folder_ID",""]])
    add("OrderCosts",["JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","DeliveryTime","CargoDetail","InternalNo","Status","Driver","隧道費","登記費","停車費","交倉附加費","其他","TotalExpense","LastSync_HKT"])

def sync_order_costs_sheet():
    try:
        df_jobs=load_df("Jobs")
        df_exp=load_df("Expenses")
        if df_jobs.empty: return 0
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
        headers=["JobID","Date","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","DeliveryTime","CargoDetail","InternalNo","Status","Driver"]+NEW_FEES+["TotalExpense","LastSync_HKT"]
        df_cost["LastSync_HKT"]=now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")
        sh=get_sh()
        ws=sh.worksheet("OrderCosts")
        ws.clear()
        ws.update('A1',[headers])
        if not df_cost.empty:
            rows=[[str(r.get(h,"")) for h in headers] for _,r in df_cost.iterrows()]
            ws.append_rows(rows, value_input_option='USER_ENTERED')
        return len(df_cost)
    except Exception as e:
        st.error(f"同步失敗: {e}")
        return 0

def find_pod_col(df):
    for col in df.columns:
        if df[col].astype(str).str.contains("data:image", na=False).any(): return col
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
    html = f"""
    <div id="print-area" style="font-family: Arial, sans-serif; padding:20px; border:2px solid #000; max-width:800px; background:white;">
    <h2 style="text-align:center; margin:0;">🚚 運輸訂單 Transport Order</h2>
    <p style="text-align:center; font-size:13px;">JobID: {job.get('JobID','')} | Internal: {job.get('InternalNo','')} | 預計: {job.get('Date','')} {job.get('DeliveryTime','')}</p>
    <hr>
    <table style="width:100%; border-collapse:collapse; font-size:14px;">
    <tr><td style="border:1px solid #000; padding:6px; width:20%; background:#eee;"><b>客戶</b></td><td style="border:1px solid #000; padding:6px;">{job.get('Client','')}</td><td style="border:1px solid #000; padding:6px; width:20%; background:#eee;"><b>Zone</b></td><td style="border:1px solid #000; padding:6px;">{job.get('Zone','')}</td></tr>
    </table>
    <h3 style="margin:10px 0 5px;">📦 起點 Pickup</h3>
    <table style="width:100%; border-collapse:collapse; font-size:14px;">
    <tr><td style="border:1px solid #000; padding:6px; background:#eee; width:20%;">地址 From</td><td style="border:1px solid #000; padding:6px;" colspan="3">{job.get('From','')}</td></tr>
    <tr><td style="border:1px solid #000; padding:6px; background:#eee;">聯絡人</td><td style="border:1px solid #000; padding:6px;">{job.get('From_Contact','')}</td><td style="border:1px solid #000; padding:6px; background:#eee; width:20%;">電話</td><td style="border:1px solid #000; padding:6px;">{job.get('From_Phone','')}</td></tr>
    <tr><td style="border:1px solid #000; padding:6px; background:#eee;">提貨時間</td><td style="border:1px solid #000; padding:6px;" colspan="3">{job.get('PickupTime','')}</td></tr>
    </table>
    <h3 style="margin:10px 0 5px;">🏁 終點 Delivery</h3>
    <table style="width:100%; border-collapse:collapse; font-size:14px;">
    <tr><td style="border:1px solid #000; padding:6px; background:#eee; width:20%;">地址 To</td><td style="border:1px solid #000; padding:6px;" colspan="3">{job.get('To','')}</td></tr>
    <tr><td style="border:1px solid #000; padding:6px; background:#eee;">聯絡人</td><td style="border:1px solid #000; padding:6px;">{job.get('To_Contact','')}</td><td style="border:1px solid #000; padding:6px; background:#eee; width:20%;">電話</td><td style="border:1px solid #000; padding:6px;">{job.get('To_Phone','')}</td></tr>
    <tr><td style="border:1px solid #000; padding:6px; background:#eee;">送貨時間</td><td style="border:1px solid #000; padding:6px;" colspan="3">{job.get('DeliveryTime','')}</td></tr>
    </table>
    <h3 style="margin:10px 0 5px;">📋 貨物資料 Cargo</h3>
    <table style="width:100%; border-collapse:collapse; font-size:14px;">
    <tr><td style="border:1px solid #000; padding:6px; background:#eee; width:25%;">詳細明細</td><td style="border:1px solid #000; padding:6px;"><pre style="margin:0; white-space:pre-wrap; font-family:inherit; font-size:14px;">{job.get('CargoDetail','')}</pre></td></tr>
    <tr><td style="border:1px solid #000; padding:6px; background:#eee;">備註</td><td style="border:1px solid #000; padding:6px;">{job.get('Notes','')}</td></tr>
    </table>
    <br><table style="width:100%; font-size:12px;"><tr><td>司機簽收: ________________</td><td>客戶簽收: ________________</td><td>日期: {job.get('Date','')}</td></tr></table>
    </div>
    <script>function printDiv(){{var c=document.getElementById('print-area').innerHTML;var w=window.open('','','height=800,width=900');w.document.write('<html><head><title>Order {job.get('JobID','')}</title></head><body>');w.document.write(c);w.document.write('</body></html>');w.document.close();w.print();}}</script>
    """
    return html

if "user" not in st.session_state: st.session_state.user=None
if not st.session_state.user:
    st.title("TMS V36 Fixed 登入")
    df_u=load_df("Users")
    u=st.text_input("用戶名"); p=st.text_input("密碼",type="password")
    if st.button("登入",type="primary", use_container_width=True, key="v36f_login"):
        m=df_u[(df_u["username"].astype(str).str.strip()==str(u).strip()) & (df_u["password"].astype(str).str.strip()==str(p).strip())]
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
if st.sidebar.button("登出", key="v36f_logout"): st.session_state.user=None; st.cache_data.clear(); st.rerun()
if st.sidebar.button("🔄 清除Cache", key="v36f_clear_cache"): st.cache_data.clear(); st.rerun()

if role=="admin": tabs=st.tabs(["📊 Dashboard","📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費","💵 每單費用","👥 用戶","🧹 維護"])
elif role=="op": tabs=st.tabs(["📊 Dashboard","📋 任務管理","📤 上傳Order","📸 上傳POD","💰 雜費","💵 每單費用","🧹 維護"])
else: tabs=st.tabs(["🟢 可接任務池","🚚 我的任務","📸 上傳POD"])

df_jobs=load_df("Jobs")
df_exp=load_df("Expenses")
df_users=load_df("Users")

with tabs[0]:
    st.subheader("📊 Dashboard")
    show_table(df_jobs.head(20))
    if st.button("🔄 同步到OrderCosts", key="v36f_sync_dash", type="primary"):
        cnt=sync_order_costs_sheet()
        st.success(f"已同步 {cnt}")

task_idx=1 if role in ["admin","op"] else None
if task_idx is not None:
    with tabs[task_idx]:
        st.subheader("📋 任務管理")
        if not df_jobs.empty:
            show_table(df_jobs.sort_values(by="JobID", ascending=False))
            sel=st.selectbox("選JobID打印", df_jobs["JobID"].astype(str).tolist(), key="v36f_task_print_sel")
            if sel:
                job=df_jobs[df_jobs["JobID"]==sel].iloc[0].to_dict()
                html=printable_order_html(job)
                st.components.v1.html(html, height=700, scrolling=True)
                st.markdown('<button onclick="printDiv()" style="background:red;color:white;padding:10px 20px;border:none;border-radius:5px;">🖨️ 點擊打印</button>', unsafe_allow_html=True)

order_idx=2 if role in ["admin","op"] else None
if order_idx is not None:
    with tabs[order_idx]:
        st.subheader("📤 上傳Order V36-Fix - 只保留詳細明細")
        tab_manual, tab_excel = st.tabs(["✍️ 手動建立", "📊 Excel批量"])
        ZONES=["港島","九龍","新界","離島","機場","中港","澳門","內地","其他"]
        with tab_manual:
            st.markdown("**📋 貨物建造器 - 入貨名+數量+UOM -> 添加到明細**")
            cb1,cb2,cb3=st.columns([3,1,1])
            cb_product=cb1.text_input("貨名 Product Code", placeholder="FPCQHKGBP003", key="v36f_cb_prod")
            cb_qty=cb2.number_input("數量", min_value=1, value=40, step=1, key="v36f_cb_qty")
            cb_uom=cb3.selectbox("UOM", UOM_OPTIONS, index=6, key="v36f_cb_uom")
            c_btn1,c_btn2,c_btn3=st.columns([1,1,2])
            with c_btn1:
                if st.button("➕ 添加", key="v36f_cb_add", type="primary", use_container_width=True):
                    if cb_product.strip():
                        line=f"{cb_product.strip()} {int(cb_qty)}{cb_uom}"
                        st.session_state.cargo_builder.append(line)
                        st.toast(f"已添加 {line}")
                    else:
                        st.error("貨名不能空")
            with c_btn2:
                if st.button("🗑️ 清空", key="v36f_cb_clear", use_container_width=True):
                    st.session_state.cargo_builder=[]
                    st.toast("已清空")
            with c_btn3:
                if st.session_state.cargo_builder:
                    st.success(" | ".join(st.session_state.cargo_builder))
                else:
                    st.write("明細: (空) - 例如 FPCQHKGBP003 40罐")

            st.divider()
            with st.form("v36f_manual_form", clear_on_submit=False):
                st.markdown("**客戶 & Internal**")
                c1,c2,c3=st.columns([2,1,1])
                client=c1.text_input("客戶 Client*")
                internal_no=c2.text_input("Internal 單號#")
                zone=c3.selectbox("邊區Zone*", ZONES + ["自定"], index=1, key="v36f_zone")
                zone_custom=st.text_input("自定邊區") if zone=="自定" else ""
                final_zone=zone_custom if zone=="自定" else zone

                st.markdown("**時間**")
                c_date,c_pick,c_del=st.columns(3)
                date=c_date.date_input("📅 預計送貨日期*", value=today_hk())
                pickup_time=c_pick.text_input("提貨時間 PickupTime")
                delivery_time=c_del.text_input("送貨時間 DeliveryTime*", placeholder="14:00-18:00")

                st.markdown("**📦 起點**")
                c_from,c_from_c,c_from_p = st.columns([3,1,1])
                from_loc=c_from.text_input("起點地址 From*")
                from_contact=c_from_c.text_input("起點聯絡人")
                from_phone=c_from_p.text_input("起點電話")

                st.markdown("**🏁 終點**")
                c_to,c_to_c,c_to_p = st.columns([3,1,1])
                to_loc=c_to.text_input("終點地址 To*")
                to_contact=c_to_c.text_input("終點聯絡人")
                to_phone=c_to_p.text_input("終點電話")

                st.markdown("**📋 最終貨物明細**")
                auto_detail = "\n".join(st.session_state.cargo_builder)
                cargo_detail=st.text_area("詳細 Cargo Detail*", value=auto_detail, height=120, placeholder="FPCQHKGBP003 40罐\nFPCQHKGBP001 20罐\nFPCQHKBA001 15袋")
                notes=st.text_input("備註 Notes")

                submitted=st.form_submit_button("✅ 建立並可打印", type="primary")
                if submitted:
                    if client.strip() and from_loc.strip() and to_loc.strip() and cargo_detail.strip():
                        jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                        append_row_by_header("Jobs",{"JobID":jid,"Date":str(date),"Client":client,"Zone":final_zone,"From":from_loc,"From_Contact":from_contact,"From_Phone":from_phone,"PickupTime":pickup_time,"To":to_loc,"To_Contact":to_contact,"To_Phone":to_phone,"DeliveryTime":delivery_time,"Driver":"","Status":"Pending","ProductName":"","QTY":"0","UOM":"","CargoDetail":cargo_detail,"InternalNo":internal_no,"Notes":notes,"CreatedBy":operator,"CreatedAt_HKT":now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")})
                        sync_order_costs_sheet()
                        job_dict={"JobID":jid,"Date":str(date),"Client":client,"Zone":final_zone,"From":from_loc,"From_Contact":from_contact,"From_Phone":from_phone,"PickupTime":pickup_time,"To":to_loc,"To_Contact":to_contact,"To_Phone":to_phone,"DeliveryTime":delivery_time,"CargoDetail":cargo_detail,"InternalNo":internal_no,"Notes":notes}
                        st.session_state.last_job=job_dict
                        st.session_state.cargo_builder=[]
                        st.success(f"✅ 已建立 {jid}")
                        st.rerun()
                    else:
                        st.error("客戶、起點、終點、詳細明細必填")

            if st.session_state.last_job:
                st.divider()
                st.subheader(f"🖨️ 上一張: {st.session_state.last_job.get('JobID')}")
                html=printable_order_html(st.session_state.last_job)
                st.components.v1.html(html, height=650, scrolling=True)
                st.markdown('<button onclick="printDiv()" style="background:#ff0000;color:white;padding:12px 24px;border:none;border-radius:8px;font-size:16px;">🖨️ 立即打印</button>', unsafe_allow_html=True)

        with tab_excel:
            template_df=pd.DataFrame(columns=["Date(預計送貨)","Client","Zone","From","From_Contact","From_Phone","PickupTime","To","To_Contact","To_Phone","DeliveryTime","CargoDetail","InternalNo","Notes"])
            template_df.loc[0]=[str(today_hk()),"Sandy","九龍","HK Warehouse","John","91234567","10:00","KLN Shop","Mary","98765432","14:00-16:00","FPCQHKGBP003 40罐\nFPCQHKGBP001 20罐","INT-001",""]
            buf_template=io.BytesIO(); template_df.to_excel(buf_template, index=False, engine='openpyxl'); buf_template.seek(0)
            st.download_button("📥 下載Template V36", data=buf_template, file_name="TMS_Order_Template_V36_OnlyDetail.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", key="v36f_dl_template")
            excel_file=st.file_uploader("上傳Excel", type=["xlsx","xls"], key="v36f_up_excel")
            if excel_file:
                try:
                    df_excel=pd.read_excel(excel_file)
                    st.dataframe(df_excel.head(20))
                    df_excel.columns=[c.strip() for c in df_excel.columns]
                    col_map={c.lower():c for c in df_excel.columns}
                    if st.button(f"✅ 批量建立 {len(df_excel)} 張", type="primary", key="v36f_batch"):
                        for _, row in df_excel.iterrows():
                            if str(row.get(col_map.get('client','Client'),"")).strip()=="": continue
                            jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                            append_row_by_header("Jobs",{"JobID":jid,"Date":str(row.get(col_map.get('date','Date'), str(today_hk()))),"Client":str(row.get(col_map.get('client','Client'), "")),"Zone":str(row.get(col_map.get('zone','Zone'), "其他")),"From":str(row.get(col_map.get('from','From'), "")),"From_Contact":str(row.get(col_map.get('from_contact','From_Contact'), "")),"From_Phone":str(row.get(col_map.get('from_phone','From_Phone'), "")),"PickupTime":str(row.get(col_map.get('pickuptime','PickupTime'), "")),"To":str(row.get(col_map.get('to','To'), "")),"To_Contact":str(row.get(col_map.get('to_contact','To_Contact'), "")),"To_Phone":str(row.get(col_map.get('to_phone','To_Phone'), "")),"DeliveryTime":str(row.get(col_map.get('deliverytime','DeliveryTime'), "")),"CargoDetail":str(row.get(col_map.get('cargodetail','CargoDetail'), "")),"InternalNo":str(row.get(col_map.get('internalno','InternalNo'), "")),"Notes":str(row.get(col_map.get('notes','Notes'), "")),"Driver":"","Status":"Pending","ProductName":"","QTY":"0","UOM":"","CreatedBy":operator,"CreatedAt_HKT":now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")})
                        sync_order_costs_sheet()
                        st.success("已批量建立"); st.rerun()
                except Exception as e: st.error(f"Excel失敗: {e}")

pod_idx=3 if role=="admin" else 3 if role=="op" else 2
with tabs[pod_idx]:
    st.subheader("📸 上傳POD")
    my_pool=df_jobs if role in ["admin","op"] else df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else df_jobs
    my_pool=my_pool.sort_values(by="JobID", ascending=False) if not my_pool.empty else my_pool
    if not my_pool.empty:
        jid=st.selectbox("選擇JobID", my_pool["JobID"].astype(str).tolist(), key="v36f_pod_sel")
        sel_info=df_jobs[df_jobs["JobID"]==jid].iloc[0].to_dict() if not df_jobs[df_jobs["JobID"]==jid].empty else {}
        st.info(f"Internal: {sel_info.get('InternalNo','')} | 明細: {sel_info.get('CargoDetail','')} | 送貨: {sel_info.get('DeliveryTime','')}")
        pod_files=st.file_uploader("拍照", type=["jpg","png","jpeg"], accept_multiple_files=True, key="v36f_pod_files")
        if pod_files and st.button(f"上傳 {len(pod_files)} 張", type="primary", key="v36f_up_pod"):
            from PIL import Image
            for idx_file, pod_file in enumerate(pod_files):
                try:
                    img=Image.open(io.BytesIO(pod_file.getvalue())).convert("RGB")
                    img.thumbnail((500,500))
                    buf=io.BytesIO(); img.save(buf, format="JPEG", quality=40, optimize=True)
                    fb=buf.getvalue()
                    b64=base64.b64encode(fb).decode()
                    data_url=f"data:image/jpeg;base64,{b64}"
                    fname=f"{jid}_POD_{now_hk().strftime('%Y%m%d_%H%M%S')}_{idx_file+1}.jpg"
                    pod_id=f"POD-{now_hk().strftime('%m%d%H%M%S')}-{uuid.uuid4().hex[:4]}"
                    append_row("PODs",[pod_id, jid, fname, str(len(fb)), data_url, operator, now_hk().strftime("%Y-%m-%d %H:%M:%S HKT"), now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")])
                except Exception as e: st.error(f"{e}")
            st.cache_data.clear(); st.success("已上傳")

cost_idx=5 if role=="admin" else 5 if role=="op" else None
if cost_idx is not None:
    with tabs[cost_idx]:
        st.subheader("💵 每單費用")
        if not df_jobs.empty:
            st.dataframe(df_jobs[["JobID","Date","DeliveryTime","InternalNo","Client","Zone","CargoDetail","Status"]].sort_values(by="JobID", ascending=False), use_container_width=True, height=400)
        if st.button("🔄 同步到OrderCosts", key="v36f_sync_cost", type="primary", use_container_width=True):
            cnt=sync_order_costs_sheet()
            st.success(f"✅ 已同步 {cnt}")

if role=="driver":
    with tabs[0]:
        st.subheader("🟢 可接任務池")
        if not df_jobs.empty:
            show_table(df_jobs[df_jobs["Driver"].astype(str).str.strip()==""].sort_values(by="JobID", ascending=False))
    with tabs[1]:
        st.subheader("🚚 我的任務")
        my=df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else pd.DataFrame()
        show_table(my)
        if not my.empty:
            sel=st.selectbox("選任務打印", my["JobID"].astype(str).tolist(), key="v36f_driver_print")
            if sel:
                job=my[my["JobID"]==sel].iloc[0].to_dict()
                html=printable_order_html(job)
                st.components.v1.html(html, height=600)
