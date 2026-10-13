import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import datetime, io

st.set_page_config(page_title="TMS 多人運輸系統", layout="wide", initial_sidebar_state="collapsed")

@st.cache_resource
def get_clients():
    scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    gc = gspread.authorize(creds)
    drive = build('drive','v3', credentials=creds)
    return gc, drive

def get_sh():
    gc,_ = get_clients()
    return gc.open_by_key(st.secrets["connections"]["gsheets"]["spreadsheet"])

@st.cache_data(ttl=10)
def load_df(tab):
    try:
        ws = get_sh().worksheet(tab)
        return pd.DataFrame(ws.get_all_records())
    except Exception as e:
        return pd.DataFrame()

def append_row(tab,row):
    get_sh().worksheet(tab).append_row(row, value_input_option='USER_ENTERED')
    st.cache_data.clear()

def append_rows(tab,rows):
    if rows:
        get_sh().worksheet(tab).append_rows(rows, value_input_option='USER_ENTERED')
        st.cache_data.clear()

def update_order(order_id, status=None, driver=None, pod_link=None):
    ws = get_sh().worksheet("Orders")
    vals = ws.get_all_values()
    h = vals[0]
    id_i = h.index("OrderID")
    for i, r in enumerate(vals[1:], start=2):
        if len(r)>id_i and r[id_i]==str(order_id):
            if status: ws.update_cell(i, h.index("Status")+1, status)
            if driver: ws.update_cell(i, h.index("Driver")+1, driver)
            if pod_link: ws.update_cell(i, h.index("POD_Link")+1, pod_link)
            break
    st.cache_data.clear()

def upload_pod(file):
    _, drive = get_clients()
    folder_id = st.secrets.get("drive_folder_id", "")
    meta = {"name": f"{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_{file.name}"}
    if folder_id: meta["parents"]=[folder_id]
    media = MediaIoBaseUpload(io.BytesIO(file.getbuffer()), mimetype=file.type)
    f = drive.files().create(body=meta, media_body=media, fields='id, webViewLink').execute()
    drive.permissions().create(fileId=f['id'], body={'type':'anyone','role':'reader'}).execute()
    return f.get('webViewLink')

# --- LOGIN ---
if "user" not in st.session_state: st.session_state.user=None
if not st.session_state.user:
    st.title("🚚 TMS 登入")
    df_u = load_df("Users")
    u = st.text_input("用戶名 (admin/op01/driver01)")
    p = st.text_input("密碼", type="password")
    if st.button("登入", use_container_width=True, type="primary"):
        if df_u.empty:
            if u=="admin" and p=="admin123":
                st.session_state.user={"username":"admin","name":"管理員","role":"admin"}
                st.rerun()
        m = df_u[(df_u["username"].astype(str).str.strip()==u.strip()) & (df_u["password"].astype(str).str.strip()==p.strip())]
        if not m.empty:
            st.session_state.user=m.iloc[0].to_dict()
            st.rerun()
        else: st.error("帳號或密碼錯誤")
    st.stop()

user=st.session_state.user
role=str(user.get("role","driver")).lower()
st.sidebar.success(f"👤 {user.get('name')} ({user.get('username')}) | {role}")
if st.sidebar.button("登出"): st.session_state.user=None; st.cache_data.clear(); st.rerun()

# --- MAIN ---
st.title("🚚 TMS 運輸管理系統")

df_orders = load_df("Orders")
df_users = load_df("Users")
df_fees = load_df("Fees")
driver_list = df_users[df_users["role"].astype(str).str.lower()=="driver"]["username"].tolist() if not df_users.empty else ["driver01"]

# 司機界面 - 手機友好
if role=="driver":
    st.subheader(f"我的運單 - {user['username']}")
    my = df_orders[df_orders["Driver"].astype(str)==user["username"]] if not df_orders.empty else pd.DataFrame()
    # 狀態篩選
    f_status = st.segmented_control("篩選", ["全部","未接單","已接單","已完成"], default="全部")
    if f_status!="全部" and not my.empty:
        my = my[my["Status"]==f_status]

    if my.empty:
        st.info("暫無派單給你")
    else:
        for _, r in my.iterrows():
            with st.container(border=True):
                st.markdown(f"**{r['OrderID']}** | {r.get('Customer','')} | {r.get('From','')} ➡️ {r.get('To','')}")
                st.caption(f"日期: {r.get('Date','')} | 狀態: {r.get('Status','')} | 備註: {r.get('Remark','')}")
                if r.get("POD_Link"): st.link_button("查看簽收相", r["POD_Link"])
                c1,c2 = st.columns(2)
                if c1.button("✅ 已接單", key=f"a_{r['OrderID']}", use_container_width=True):
                    update_order(r['OrderID'], status="已接單"); st.toast("已接單"); st.rerun()
                # 簽收上傳
                up = c2.file_uploader("簽收相上傳", type=["jpg","jpeg","png"], key=f"up_{r['OrderID']}", label_visibility="collapsed")
                if up:
                    with st.spinner("上傳中..."):
                        link = upload_pod(up)
                        update_order(r['OrderID'], status="已完成", pod_link=link)
                        st.success("已完成並上傳"); st.rerun()

    st.divider()
    st.subheader("💰 雜費申報")
    with st.form("fee_form"):
        oid = st.selectbox("OrderID", my["OrderID"].tolist() if not my.empty else [])
        t = st.selectbox("類型", ["停車場","隧道費","登記費","油費","其他"])
        amt = st.number_input("金額 $", min_value=0.0, step=1.0)
        remark = st.text_input("備註 (例如：西隧)")
        if st.form_submit_button("提交雜費", use_container_width=True):
            append_row("Fees", [str(datetime.date.today()), oid, user["username"], t, amt, remark])
            st.success(f"已提交 {t} ${amt}"); st.balloons()

else: # Admin / Operator
    tab1, tab2, tab3, tab4 = st.tabs(["📋 訂單總覽","📤 訂單上傳","💰 雜費管理","👥 用戶"])

    with tab1:
        col_f1, col_f2 = st.columns(2)
        s_filter = col_f1.selectbox("狀態篩選", ["全部","未接單","已接單","已完成"])
        d_filter = col_f2.selectbox("司機篩選", ["全部"]+driver_list)
        view_df = df_orders.copy()
        if s_filter!="全部" and not view_df.empty: view_df=view_df[view_df["Status"]==s_filter]
        if d_filter!="全部" and not view_df.empty: view_df=view_df[view_df["Driver"]==d_filter]
        st.dataframe(view_df, use_container_width=True, height=400)

        if not df_orders.empty:
            st.divider()
            st.write("**快速改狀態 / 派單**")
            c1,c2,c3,c4 = st.columns(4)
            sel_id = c1.selectbox("OrderID", df_orders["OrderID"].tolist())
            new_s = c2.selectbox("狀態", ["未接單","已接單","已完成"])
            new_d = c3.selectbox("司機", [""]+driver_list)
            if c4.button("更新", type="primary", use_container_width=True):
                update_order(sel_id, status=new_s, driver=new_d if new_d else None)
                st.success(f"{sel_id} 已更新為 {new_s}"); st.rerun()

    with tab2:
        st.subheader("批量上傳訂單")
        st.download_button("下載Excel範本", data="OrderID,Date,Customer,From,To,Remark\nORD001,2026-05-13,客戶A,荃灣,觀塘,急件\nORD002,2026-05-13,客戶B,葵涌,中環,,\n", file_name="TMS_Orders_Template.csv")
        f = st.file_uploader("上傳 Excel / CSV", type=["xlsx","csv"])
        if f:
            df_up = pd.read_csv(f) if f.name.endswith(".csv") else pd.read_excel(f)
            st.dataframe(df_up.head(), use_container_width=True)
            if st.button("確認上傳", type="primary"):
                rows=[]
                for _,r in df_up.iterrows():
                    rows.append([str(r.get("OrderID","")), str(r.get("Date",datetime.date.today())), str(r.get("Customer","")), str(r.get("From","")), str(r.get("To","")), "未接單", "", "", str(r.get("Remark",""))])
                append_rows("Orders", rows)
                st.success(f"已上傳 {len(rows)} 張訂單"); st.balloons(); st.rerun()

    with tab3:
        st.dataframe(df_fees, use_container_width=True)
        if not df_fees.empty:
            df_fees["Amount"]=pd.to_numeric(df_fees["Amount"], errors='coerce').fillna(0)
            c1,c2,c3 = st.columns(3)
            c1.metric("總雜費", f"${df_fees['Amount'].sum():.2f}")
            c2.metric("停車場", f"${df_fees[df_fees['Type']=='停車場']['Amount'].sum():.2f}")
            c3.metric("隧道費", f"${df_fees[df_fees['Type']=='隧道費']['Amount'].sum():.2f}")
        # 手動輸入雜費
        with st.form("fee_admin"):
            st.write("代司機輸入雜費")
            oid = st.selectbox("OrderID", df_orders["OrderID"].tolist() if not df_orders.empty else [])
            t = st.selectbox("類型", ["停車場","隧道費","登記費","其他"])
            amt = st.number_input("金額", min_value=0.0)
            drv = st.selectbox("司機", driver_list)
            if st.form_submit_button("新增雜費"):
                append_row("Fees", [str(datetime.date.today()), oid, drv, t, amt, "Admin代入"])
                st.success("已新增"); st.rerun()

    with tab4:
        if role=="admin":
            st.dataframe(df_users, use_container_width=True)
            with st.form("add_user"):
                c1,c2,c3,c4 = st.columns(4)
                nu=c1.text_input("新用戶名"); np=c2.text_input("密碼"); nn=c3.text_input("姓名"); nr=c4.selectbox("角色",["operator","driver","admin"])
                if st.form_submit_button("新增用戶"):
                    append_row("Users", [nu,np,nn,nr]); st.success(f"已新增 {nu}"); st.rerun()
        else:
            st.info("只有Admin可以管理用戶")