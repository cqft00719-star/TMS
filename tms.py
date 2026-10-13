# --- LOGIN - 修復版 ---
if "user" not in st.session_state:
    st.session_state.user = None

if not st.session_state.user:
    st.title("TMS 多人運輸系統 - 登入")
    st.caption("用戶名 (admin/op01/driver01)")

    df_u = load_df("Users")
    # 防呆：如果Users表空咗或者欄名唔啱，自動用預設帳號
    if df_u.empty or "username" not in df_u.columns or "password" not in df_u.columns:
        df_u = pd.DataFrame([
            {"username":"admin","password":"admin123","name":"Admin","role":"admin"},
            {"username":"op01","password":"op01","name":"OP01","role":"op"},
            {"username":"driver01","password":"driver01","name":"John","role":"driver"},
            {"username":"John","password":"John123","name":"John","role":"driver"},
        ])

    # 確保係字串，唔會strip出錯
    df_u["username"] = df_u["username"].astype(str).str.strip()
    df_u["password"] = df_u["password"].astype(str).str.strip()

    u = st.text_input("用戶名 (admin/op01/driver01)", key="tms_user")
    p = st.text_input("密碼", type="password", key="tms_pass")

    if st.button("登入", type="primary", use_container_width=True, key="tms_login"):
        if not u or not p:
            st.error("請輸入用戶名同密碼")
        else:
            m = df_u[(df_u["username"]==str(u).strip()) & (df_u["password"]==str(p).strip())]
            if not m.empty:
                st.session_state.user = m.iloc[0].to_dict()
                st.success(f"歡迎 {m.iloc[0]['username']}")
                st.rerun()
            else:
                st.error(f"帳號或密碼錯誤，查到嘅用戶有: {', '.join(df_u['username'].tolist())}")
    st.stop()
