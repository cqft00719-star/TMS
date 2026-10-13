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
st.set_page_config(page_title="TMS V33 Date=EDD QTY+UOM", layout="wide", page_icon="🚚")

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
            ws=sh.add_worksheet(title=t, rows=2000, cols=35)
            ws.update('A1',[h])
            if d: ws.append_rows(d, value_input_option='USER_ENTERED')
    add("Users",["username","password","name","role"],[["admin","admin123","Admin","admin"],["John","John123","John","driver"],["Mo","Mo123","Mo","driver"],["op01","op01","OP01","op"]])
    # V33 Jobs 加 UOM，Date 備註係預計送貨日期
    add("Jobs",["JobID","Date","Client","Zone","From","To","Driver","Status","QTY","UOM","Notes","CreatedBy","CreatedAt_HKT","POD_Link","PodTime_HKT","POD_Image"])
    add("PODs",["POD_ID","JobID","FileName","Size","Base64_Image","CreatedBy","CreatedAt_HKT","PodTime_HKT"])
    add("Expenses",["ExpenseID","Date","JobID","Driver","Type","Amount","Remark","CreatedAt_HKT"])
    add("Config",["Key","Value"],[["POD_Folder_ID",""]])
    add("OrderCosts",["JobID","Date","Client","Zone","From","To","Driver","Status","QTY","UOM","隧道費","登記費","停車費","交倉附加費","其他","TotalExpense","LastSync_HKT"])
    try:
        ws=sh.worksheet("Jobs")
        headers=[h.strip() for h in ws.row_values(1)]
        # 自動加 UOM 欄如果冇
        if "UOM" not in headers:
            # 喺 QTY 後面加
            new_headers=[]
            for h in headers:
                new_headers.append(h)
                if h=="QTY" and "UOM" not in headers:
                    new_headers.append("UOM")
            if "UOM" not in new_headers:
                new_headers.append("UOM")
            # 如果舊表只有11欄，補齊到新格式
            if len(new_headers)<16:
                new_headers=["JobID","Date","Client","Zone","From","To","Driver","Status","QTY","UOM","Notes","CreatedBy","CreatedAt_HKT","POD_Link","PodTime_HKT","POD_Image"]
            ws.update('A1',[new_headers])
        # OrderCosts 都加 UOM
        ws2=sh.worksheet("OrderCosts")
        h2=[h.strip() for h in ws2.row_values(1)]
        if "UOM" not in h2:
            ws2.update('A1',[["JobID","Date","Client","Zone","From","To","Driver","Status","QTY","UOM","隧道費","登記費","停車費","交倉附加費","其他","TotalExpense","LastSync_HKT"]])
    except: pass

def sync_order_costs_sheet():
    try:
        df_jobs=load_df("Jobs")
        df_exp=load_df("Expenses")
        if df_jobs.empty: return 0
        if "Zone" not in df_jobs.columns: df_jobs["Zone"]=""
        if "UOM" not in df_jobs.columns: df_jobs["UOM"]=""
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
        display_cols=["JobID","Date","Client","Zone","From","To","Driver","Status","QTY","UOM"]+NEW_FEES+["TotalExpense"]
        display_cols=[c for c in display_cols if c in df_cost.columns]
        df_out=df_cost[display_cols].sort_values(by="JobID", ascending=False) if "JobID" in df_cost.columns else df_cost[display_cols]
        df_out["LastSync_HKT"]=now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")
        sh=get_sh()
        ws=sh.worksheet("OrderCosts")
        ws.clear()
        headers=["JobID","Date","Client","Zone","From","To","Driver","Status","QTY","UOM"]+NEW_FEES+["TotalExpense","LastSync_HKT"]
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

def auto_fix_zone():
    try:
        ws=get_sh().worksheet("Jobs")
        all_vals=ws.get_all_values()
        if len(all_vals)<2: return
        headers=all_vals[0]
        new_rows=[]; fixed=0
        for r in all_vals[1:]:
            if len(r)>=7 and (r[3] in ["2","21C"] or (len(r)>5 and r[5] in ["Mo","Kit","John"] and r[3] not in ["港島","九龍","新界","離島","機場","中港","澳門","內地","其他"])):
                old_from=r[3]; old_to=r[4]; old_driver=r[5]; old_status=r[6]; old_qty=r[7] if len(r)>7 else ""; old_notes=r[8] if len(r)>8 else ""
                new_r=[r[0],r[1],r[2],"",old_from,old_to,old_driver,old_status,old_qty,""] + [old_notes] + r[9:]
                new_rows.append(new_r); fixed+=1
            else:
                new_rows.append(r)
        if fixed>0:
            ws.clear(); ws.update('A1',[headers]); ws.append_rows(new_rows, value_input_option='USER_ENTERED')
            st.cache_data.clear()
    except: pass

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

if "user" not in st.session_state: st.session_state.user=None
if not st.session_state.user:
    st.title("TMS V33 登入")
    df_u=load_df("Users")
    u=st.text_input("用戶名"); p=st.text_input("密碼",type="password")
    if st.button("登入",type="primary", use_container_width=True):
        m=df_u[(df_u["username"].astype(str).str.strip()==str(u).strip()) & (df_u["password"].astype(str).str.strip()==str(p).strip())]
        if not m.empty: st.session_state.user=m.iloc[0].to_dict(); st.rerun()
        else: st.error("帳號錯")
    st.stop()

ensure_sheets()
auto_fix_zone()
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
        if "Zone" in df_jobs.columns: c3.metric("有Zone", len(df_jobs[df_jobs["Zone"].astype(str).str.strip()!=""]))
        if "UOM" in df_jobs.columns: c4.metric("總數量", pd.to_numeric(df_jobs["QTY"], errors='coerce').sum())
        if "Zone" in df_jobs.columns and df_jobs["Zone"].astype(str).str.strip().ne("").any():
            st.bar_chart(df_jobs["Zone"].astype(str).value_counts())
    show_table(df_jobs.head(20))
    if st.button("🔄 同步每單費用到OrderCosts", type="primary"):
        cnt=sync_order_costs_sheet()
        st.success(f"已同步 {cnt} 單")

task_idx=1 if role in ["admin","op"] else None
if task_idx is not None:
    with tabs[task_idx]:
        st.subheader("📋 任務管理")
        if not df_jobs.empty:
            with st.container(border=True):
                f1,f2,f3,f4=st.columns(4)
                all_status=df_jobs["Status"].astype(str).unique().tolist() if "Status" in df_jobs.columns else []
                sel_status=f1.multiselect("Status", all_status, default=all_status)
                all_zone=df_jobs["Zone"].astype(str).unique().tolist() if "Zone" in df_jobs.columns else []
                sel_zone=f2.multiselect("Zone", [z for z in all_zone if z.strip()!=""])
                all_clients=df_jobs["Client"].astype(str).unique().tolist() if "Client" in df_jobs.columns else []
                sel_clients=f3.multiselect("Client", all_clients)
                all_uom=df_jobs["UOM"].astype(str).unique().tolist() if "UOM" in df_jobs.columns else []
                sel_uom=f4.multiselect("UOM", [u for u in all_uom if u.strip()!=""])
                df_filtered=df_jobs.copy()
                if sel_status: df_filtered=df_filtered[df_filtered["Status"].astype(str).isin(sel_status)]
                if sel_zone: df_filtered=df_filtered[df_filtered["Zone"].astype(str).isin(sel_zone)]
                if sel_clients: df_filtered=df_filtered[df_filtered["Client"].astype(str).isin(sel_clients)]
                if sel_uom: df_filtered=df_filtered[df_filtered["UOM"].astype(str).isin(sel_uom)]
                df_filtered=df_filtered.sort_values(by="JobID", ascending=False) if "JobID" in df_filtered.columns else df_filtered
            show_table(df_filtered)

order_idx=2 if role in ["admin","op"] else None
if order_idx is not None:
    with tabs[order_idx]:
        st.subheader("📤 上傳Order - 日期=預計送貨日期 | 數量=整數 | 有UOM")
        tab_manual, tab_excel = st.tabs(["✍️ 手動", "📊 Excel批量"])
        ZONES=["港島","九龍","新界","離島","機場","中港","澳門","內地","其他"]
        with tab_manual:
            with st.form("manual_v33", clear_on_submit=True):
                c1,c2=st.columns(2)
                client=c1.text_input("客戶*")
                zone=c2.selectbox("邊區Zone*", ZONES + ["自定"], index=1)
                zone_custom=st.text_input("自定邊區 (選自定才填)") if zone=="自定" else ""
                final_zone=zone_custom if zone=="自定" else zone
                # V33 明確label係預計送貨日期
                date=st.date_input("📅 預計送貨日期 (Expected Delivery Date)*", value=today_hk())
                c3,c4=st.columns(2)
                from_loc=c3.text_input("起點"); to_loc=c4.text_input("終點")
                c5,c6=st.columns(2)
                qty=c5.number_input("數量 QTY (整數)*", min_value=1, max_value=10000, value=1, step=1, format="%d")
                uom=c6.selectbox("單位 UOM*", UOM_OPTIONS, index=1)
                uom_custom=st.text_input("自定UOM (選其他才填)") if uom=="其他" else ""
                final_uom=uom_custom if uom=="其他" and uom_custom.strip()!="" else uom
                notes=st.text_input("備註")
                if st.form_submit_button("建立", type="primary"):
                    if client.strip():
                        jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                        append_row_by_header("Jobs",{"JobID":jid,"Date":str(date),"Client":client,"Zone":final_zone,"From":from_loc,"To":to_loc,"Driver":"","Status":"Pending","QTY":str(int(qty)),"UOM":final_uom,"Notes":notes,"CreatedBy":operator,"CreatedAt_HKT":now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")})
                        sync_order_costs_sheet()
                        st.success(f"✅ {jid} 預計 {date} {int(qty)}{final_uom} Zone:{final_zone}"); st.rerun()
        with tab_excel:
            template_df=pd.DataFrame(columns=["Date(預計送貨日期)","Client","Zone","From","To","QTY(整數)","UOM(袋/箱/板)","Notes"])
            template_df.loc[0]=[str(today_hk()),"Sandy","九龍","HK","KLN",1,"箱","Sample"]
            template_df.loc[1]=[str(today_hk()),"ABC Ltd","港島","Central","TST",2,"板","Urgent"]
            buf_template=io.BytesIO(); template_df.to_excel(buf_template, index=False, engine='openpyxl'); buf_template.seek(0)
            st.download_button("📥 下載Excel Template (有UOM+整數QTY)", data=buf_template, file_name="TMS_Order_Template_V33_QTY_UOM.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
            excel_file=st.file_uploader("上傳已填好嘅Excel", type=["xlsx","xls"])
            if excel_file:
                try:
                    df_excel=pd.read_excel(excel_file)
                    st.dataframe(df_excel.head(20))
                    df_excel.columns=[c.strip() for c in df_excel.columns]
                    col_map={c.lower():c for c in df_excel.columns}
                    # 自動搵QTY同UOM欄
                    qty_col = next((col_map[k] for k in col_map if "qty" in k), "QTY")
                    uom_col = next((col_map[k] for k in col_map if "uom" in k), "UOM")
                    if st.button(f"✅ 確認批量建立 {len(df_excel)} 張Order (QTY整數+UOM)", type="primary"):
                        for _, row in df_excel.iterrows():
                            if str(row.get(col_map.get('client','Client'),"")).strip()=="": continue
                            try: q=int(float(row.get(qty_col, 1)))
                            except: q=1
                            jid=f"JOB-{now_hk().strftime('%m%d%H%M')}-{str(uuid.uuid4())[:4].upper()}"
                            append_row_by_header("Jobs",{"JobID":jid,"Date":str(row.get(col_map.get('date','Date'), str(today_hk()))),"Client":str(row.get(col_map.get('client','Client'), "")),"Zone":str(row.get(col_map.get('zone','Zone'), "其他")),"From":str(row.get(col_map.get('from','From'), "")),"To":str(row.get(col_map.get('to','To'), "")),"QTY":str(q),"UOM":str(row.get(uom_col, "箱")),"Notes":str(row.get(col_map.get('notes','Notes'), "")),"Driver":"","Status":"Pending","CreatedBy":operator,"CreatedAt_HKT":now_hk().strftime("%Y-%m-%d %H:%M:%S HKT")})
                        sync_order_costs_sheet()
                        st.success("已批量建立並同步到OrderCosts"); st.rerun()
                except Exception as e: st.error(f"Excel失敗: {e}")

pod_idx=3 if role=="admin" else 3 if role=="op" else 2
with tabs[pod_idx]:
    st.subheader("📸 上傳POD")
    my_pool=df_jobs if role in ["admin","op"] else df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else df_jobs
    my_pool=my_pool.sort_values(by="JobID", ascending=False) if not my_pool.empty else my_pool
    if not my_pool.empty:
        jid=st.selectbox("選擇JobID", my_pool["JobID"].astype(str).tolist())
        sel_info=df_jobs[df_jobs["JobID"]==jid].iloc[0] if not df_jobs[df_jobs["JobID"]==jid].empty else None
        if sel_info is not None:
            st.info(f"預計送貨: {sel_info.get('Date','')} | {sel_info.get('QTY','')} {sel_info.get('UOM','')} | Zone: {sel_info.get('Zone','')} | Client: {sel_info.get('Client','')}")
        pod_files=st.file_uploader("拍照 / 選圖 (可一次選多張)", type=["jpg","png","jpeg"], accept_multiple_files=True)
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
        st.subheader("💰 雜費管理")
        show_table(df_exp)
        st.divider()
        with st.form("add_expense_v33", clear_on_submit=True):
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
        st.subheader("💵 每一個Order的費用 - V33 含QTY整數+UOM")
        if df_jobs.empty:
            st.info("無Jobs")
        else:
            df_cost=df_jobs.copy()
            if "Zone" not in df_cost.columns: df_cost["Zone"]=""
            if "UOM" not in df_cost.columns: df_cost["UOM"]=""
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
            with st.container(border=True):
                c1,c2,c3,c4=st.columns(4)
                all_zone=df_cost["Zone"].astype(str).unique().tolist() if "Zone" in df_cost.columns else []
                sel_zone=c1.multiselect("篩選Zone", [z for z in all_zone if z.strip()!=""])
                all_client=df_cost["Client"].astype(str).unique().tolist() if "Client" in df_cost.columns else []
                sel_client=c2.multiselect("篩選Client", all_client)
                all_uom=df_cost["UOM"].astype(str).unique().tolist() if "UOM" in df_cost.columns else []
                sel_uom=c3.multiselect("篩選UOM", [u for u in all_uom if u.strip()!=""])
                only_with_exp=c4.checkbox("只顯示有雜費")
                df_show=df_cost.copy()
                if sel_zone: df_show=df_show[df_show["Zone"].astype(str).isin(sel_zone)]
                if sel_client: df_show=df_show[df_show["Client"].astype(str).isin(sel_client)]
                if sel_uom: df_show=df_show[df_show["UOM"].astype(str).isin(sel_uom)]
                if only_with_exp: df_show=df_show[df_show["TotalExpense"]>0]
            display_cols=["JobID","Date","Client","Zone","From","To","Driver","Status","QTY","UOM"]+NEW_FEES+["TotalExpense"]
            display_cols=[c for c in display_cols if c in df_show.columns]
            df_display=df_show[display_cols].sort_values(by="JobID", ascending=False) if "JobID" in df_show.columns else df_show[display_cols]
            st.write(f"**共 {len(df_display)} 單，總雜費 {df_display['TotalExpense'].sum():.2f}**")
            st.dataframe(df_display, use_container_width=True, height=500)
            col_sync,col_dl=st.columns(2)
            with col_sync:
                if st.button("🔄 同步到Google Sheet OrderCosts", type="primary", use_container_width=True):
                    cnt=sync_order_costs_sheet()
                    st.success(f"✅ 已同步 {cnt} 單")
            with col_dl:
                buf=io.BytesIO()
                df_display.to_excel(buf, index=False, engine='openpyxl')
                buf.seek(0)
                st.download_button("📥 下載Excel (無亂碼)", data=buf, file_name=f"Cost_Per_Order_{now_hk().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)

user_idx=6 if role=="admin" else None
if user_idx is not None:
    with tabs[user_idx]:
        st.subheader("👥 用戶管理")
        df_show=df_users.copy()
        if "password" in df_show.columns: df_show["password"]="******"
        st.markdown(f"<div style='overflow-x:auto'>{df_show.to_html(index=False, border=0)}</div>", unsafe_allow_html=True)
        st.divider()
        with st.form("add_user_v33", clear_on_submit=True):
            new_u=st.text_input("用戶名*"); new_p=st.text_input("密碼*")
            new_name=st.text_input("姓名"); new_role=st.selectbox("角色", ["admin","op","driver"])
            if st.form_submit_button("建立用戶", type="primary"):
                if new_u.strip() and new_p.strip():
                    append_row("Users",[new_u.strip(), new_p.strip(), new_name.strip() or new_u.strip(), new_role])
                    st.success(f"✅ {new_u}"); st.rerun()

maint_idx=7 if role=="admin" else 6 if role=="op" else None
if maint_idx is not None:
    with tabs[maint_idx]:
        st.subheader("🧹 維護")
        if st.button("🔄 同步每單費用到 OrderCosts Sheet", type="primary", use_container_width=True):
            cnt=sync_order_costs_sheet()
            st.success(f"已同步 {cnt} 單")

if role=="driver":
    with tabs[0]:
        st.subheader("🟢 可接任務池")
        if not df_jobs.empty:
            mask=df_jobs["Driver"].astype(str).str.strip()==""
            avail=df_jobs[mask | (df_jobs["Status"].astype(str).str.lower()=="pending")]
            show_table(avail)
            for _, row in avail.iterrows():
                jid=row["JobID"]
                if st.button(f"搶 {jid}", key=f"claim_{jid}_v33", type="primary"):
                    ws=get_sh().worksheet("Jobs")
                    headers=[h.strip() for h in ws.row_values(1)]
                    cell=ws.find(jid)
                    if cell:
                        if "Driver" in headers: ws.update_cell(cell.row, headers.index("Driver")+1, operator)
                        if "Status" in headers: ws.update_cell(cell.row, headers.index("Status")+1, "Accepted")
                        st.cache_data.clear(); st.rerun()
    with tabs[1]:
        st.subheader("🚚 我的任務"); my=df_jobs[df_jobs["Driver"].astype(str).str.strip()==operator] if not df_jobs.empty else pd.DataFrame(); show_table(my)
