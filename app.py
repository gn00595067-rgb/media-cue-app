import streamlit as st
import pandas as pd
from datetime import timedelta, date
import io
import requests
import json

# ==========================================
# 0. Ragic 設定區 (請填入真實 ID)
# ==========================================
RAGIC_CONFIG = {
    "client_name": 10012,
    "start_date": 10013,
    "end_date": 10014,
    "region": 10015,
    "total_budget": 10016,
    
    # 子表格欄位 ID (請確認這些都填對)
    "sub_station": 10017,
    "sub_sec": 10018,
    "sub_rate": 10019,
    "sub_cost": 10020,
    "sub_spots": 10021,
}

# ==========================================
# 1. 費率與基礎函式
# ==========================================
RATE_CARD = {
    "全家便利商店": {
        "全省": {"10s": 150, "15s": 200, "20s": 260},
        "北部": {"10s": 180, "15s": 240, "20s": 310},
        "中部": {"10s": 150, "15s": 200, "20s": 260},
        "南部": {"10s": 150, "15s": 200, "20s": 260},
    },
    "全家新鮮視": {
        "全省": {"10s": 400, "15s": 500, "20s": 600},
        "北部": {"10s": 450, "15s": 550, "20s": 650},
        "中部": {"10s": 400, "15s": 500, "20s": 600},
        "南部": {"10s": 400, "15s": 500, "20s": 600},
    },
    "家樂福": {
        "全省": {"10s": 130, "15s": 180, "20s": 230},
        "北部": {"10s": 160, "15s": 210, "20s": 260},
        "中部": {"10s": 130, "15s": 180, "20s": 230},
        "南部": {"10s": 130, "15s": 180, "20s": 230},
    }
}

WEEKDAY_MAP = {0: "一", 1: "二", 2: "三", 3: "四", 4: "五", 5: "六", 6: "日"}

def calculate_single_schedule(channel, region, sec, budget, s_date, e_date, program_name):
    if budget <= 0: return None
    try:
        rate = RATE_CARD.get(channel, {}).get(region, {}).get(sec, 0)
    except:
        rate = 0
    if rate == 0: return None
    
    total_spots = int(budget / rate)
    days = (e_date - s_date).days + 1
    base = total_spots // days
    remainder = total_spots % days
    
    schedule = []
    current = s_date
    for i in range(days):
        spots = base + (1 if i < remainder else 0)
        schedule.append(spots)
        current += timedelta(days=1)
        
    return {
        "Station": channel,
        "Location": region,
        "Program": program_name,
        "Day-part": "06-24",
        "Size": sec,
        "Rate (Net)": rate,
        "Package Cost": int(budget),
        "Schedule": schedule,
        "Total Spots": total_spots
    }

def render_mix_ui_v2(channel_name, key_id, budget, region, start_date, end_date, program_name):
    if budget <= 0: return []
    st.markdown("---")
    st.caption(f"🔻 {channel_name} 秒數與佔比")

    cols_chk = st.columns(3)
    with cols_chk[0]: use_10 = st.checkbox("10s", value=False, key=f"c10_{key_id}")
    with cols_chk[1]: use_15 = st.checkbox("15s", value=True, key=f"c15_{key_id}")
    with cols_chk[2]: use_20 = st.checkbox("20s", value=False, key=f"c20_{key_id}")
    
    active_secs = []
    if use_10: active_secs.append("10s")
    if use_15: active_secs.append("15s")
    if use_20: active_secs.append("20s")
    
    count = len(active_secs)
    pcts = {}

    if count == 0:
        st.warning("請至少勾選一種秒數")
        return []
    elif count == 1:
        sec = active_secs[0]
        pcts[sec] = 100
        st.info(f"✅ {sec} 佔比: 100%")
    elif count == 2:
        sec_a, sec_b = active_secs[0], active_secs[1]
        val_a = st.slider(f"{sec_a} 佔比", 0, 100, 50, key=f"sl2_{key_id}")
        val_b = 100 - val_a
        pcts[sec_a] = val_a
        pcts[sec_b] = val_b
        st.write(f"{sec_b} 自動連動: **{val_b}%**")
        st.progress(val_b/100)
    elif count == 3:
        st.caption("手動分配 (需等於 100%)")
        c1, c2, c3 = st.columns(3)
        with c1: val_10 = st.number_input("10s %", 0, 100, 33, key=f"ni3_10_{key_id}")
        with c2: val_15 = st.number_input("15s %", 0, 100, 33, key=f"ni3_15_{key_id}")
        with c3: val_20 = st.number_input("20s %", 0, 100, 34, key=f"ni3_20_{key_id}")
        total_p = val_10 + val_15 + val_20
        if total_p != 100:
            st.error(f"合計 {total_p}% (請調整至 100%)")
        else:
            st.success("合計 100%")
        pcts["10s"] = val_10
        pcts["15s"] = val_15
        pcts["20s"] = val_20

    result_rows = []
    for sec, pct in pcts.items():
        if pct > 0:
            sub_budget = budget * (pct / 100)
            r = calculate_single_schedule(channel_name, region, sec, sub_budget, start_date, end_date, program_name)
            if r: result_rows.append(r)
    return result_rows

# ==========================================
# 2. UI 頁面開始
# ==========================================
st.set_page_config(page_title="媒體排程系統 v7.0", layout="wide")
st.title("📱 媒體報價系統 v7.0")

# 步驟 1 (新增客戶名稱)
with st.expander("🛠️ 步驟 1：基礎資訊", expanded=True):
    client_name = st.text_input("客戶名稱", placeholder="例如：台灣讀廣-南港lalaport")
    
    st.divider()
    
    c1, c2, c3 = st.columns(3)
    with c1: start_date = st.date_input("開始日期", value=date.today())
    with c2: end_date = st.date_input("結束日期", value=date.today() + timedelta(days=29))
    with c3: region = st.selectbox("投放區域", ["全省", "北部", "中部", "南部"])
    total_budget = st.number_input("總預算 (未稅)", value=500000, step=10000)
    total_days = (end_date - start_date).days + 1

# 步驟 2
st.divider()
st.subheader("🛠️ 步驟 2：通路配置")
sel_c1, sel_c2, sel_c3 = st.columns(3)
with sel_c1: enable_fm = st.checkbox("全家便利商店", value=True)
with sel_c2: enable_fv = st.checkbox("全家新鮮視", value=True)
with sel_c3: enable_cf = st.checkbox("家樂福", value=False)

active_channels = []
if enable_fm: active_channels.append("FM")
if enable_fv: active_channels.append("FV")
if enable_cf: active_channels.append("CF")
active_count = len(active_channels)

all_schedule_rows = []
layout_c1, layout_c2, layout_c3 = st.columns(3)
pct_fm, pct_fv, pct_cf = 0, 0, 0

with layout_c1:
    if enable_fm:
        st.info("🏪 全家便利商店")
        if active_count == 2 and active_channels[0] != "FM":
             st.caption("自動連動計算中...")
             pct_fm = 0 
        else:
             pct_fm = st.slider("預算佔比 %", 0, 100, 50 if active_count==1 else 33, key="sl_fm")
        budget_fm_placeholder = st.empty()

with layout_c2:
    if enable_fv:
        st
