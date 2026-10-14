import streamlit as st, pandas as pd, gspread, datetime, uuid, io, base64, requests, urllib.parse
from google.oauth2.service_account import Credentials
from datetime import timezone, timedelta
HKT=timezone(timedelta(hours=8))
def now_hk(): return datetime.datetime.now(HKT)
def today_hk(): return now_hk().date()
st.set_page_config(page_title="TMS V39 CallMeBot OK", layout="wide", page_icon="🚚")

def send_whatsapp_to_ops(message):
    try:
        wa_cfg=st.secrets.get("whatsapp",{})
        if not wa_cfg.get("enabled", False): return False,"未啟用"
        phones=wa_cfg.get("callmebot_phones",[])
        keys=wa_cfg.get("callmebot_apikeys",[])
        ok=0
        for i,ph in enumerate(phones):
            key=keys[i] if i < len(keys) else keys[0] if keys else ""
            url=f"https://api.callmebot.com/whatsapp.php?phone={ph}&text={urllib.parse.quote(message)}&apikey={key}"
            try:
                r=requests.get(url,timeout=15)
                if "Message sent" in r.text or r.status_code==200: ok+=1
            except Exception as e: st.error(f"{e}")
        return True,f"已發送 {ok}/{len(phones)} 個Op"
    except Exception as e: return False,str(e)

def notify_new_order(job):
    msg=f"🆕新Order {now_hk().strftime('%m/%d %H:%M')}\nJobID:{job.get('JobID')}\nClient:{job.get('Client')} Internal:{job.get('InternalNo')}\nZone:{job.get('Zone')}\n起點:{job.get('From')} {job.get('From_Contact')} {job.get('From_Phone')} 提貨:{job.get('PickupTime')}\n終點:{job.get('To')} {job.get('To_Contact')} {job.get('To_Phone')} 送貨:{job.get('DeliveryTime')}\n貨:{job.get('CargoDetail')}\n"
    return send_whatsapp_to_ops(msg)
def notify_status(jid,driver,status,client=""):
    return send_whatsapp_to_ops(f"🔄{now_hk().strftime('%m/%d %H:%M')}\nJobID:{jid}\nClient:{client}\n司機:{driver} 已{status}")
def notify_pod(jid,driver,client="",cnt=1):
    return send_whatsapp_to_ops(f"✅{now_hk().strftime('%m/%d %H:%M')}\nJobID:{jid}\nClient:{client}\n司機:{driver} 已上傳{cnt}張POD Delivered")

#...下面你V38嘅get_client, get_sh, load_df, append_row_by_header, ensure_sheets, sync_order_costs_sheet, printable_order_html 全部keep...
# 為咗短，我唔重貼，你直接用我上一版V39個完整file，只係將Secrets改做你而家呢個就得
