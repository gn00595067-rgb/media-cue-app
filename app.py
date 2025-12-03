import streamlit as st
import pandas as pd
from datetime import timedelta, date
import io

# ==========================================
# 1. 系統設定 (費率卡 Rate Card)
# ==========================================
RATE_CARD = {
    "全家便利商店": { # 通路廣播
        "全省": {"10s": 150, "15s": 200, "20s": 260},
        "北部": {"10s": 180, "15s": 240, "20s": 310},
        "中部": {"10s": 150, "15s": 200, "20s": 260},
        "南部": {"10s": 150, "15s": 200, "20s": 260},
    },
    "全家新鮮視": { # TV 螢幕
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

st.set_page_config(page_title="媒體排程報價系統 v3.1", layout="wide")

# ==========================================
# 2. 核心計算函式
# ==========================================
def calculate_single_schedule(channel, region, sec, budget, s_date, e_date, program_name):
    """計算單一條目(Row)的排程"""
    if budget <= 0: return None
    
    # 查價
    try:
        rate = RATE_CARD.get(channel, {}).get(region, {}).get(sec, 0)
    except:
        rate = 0
    if rate == 0: return None # 防呆
    
    # 算檔次
    total_spots = int(budget / rate)
    
    # 每日分配
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

# ==========================================
# 3. UI: 基礎設定
# ==========================================
st.title("📱 媒體報價系統 v3.1")

with st.expander("🛠️ 步驟 1：設定走期與總預算", expanded=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        start_date = st.date_input("開始日期", value=date.today())
    with col2:
        end_date = st.date_input("結束日期", value=date.today() + timedelta(days=29))
    with col3:
        region = st.selectbox("投放區域", ["全省", "北部", "中部", "南部"])
        
    total_budget = st.number_input("總預算 (未稅)", min_value=10000, value=500000, step=10000)
    total_days = (end_date - start_date).days + 1
    st.caption(f"📅 總走期：{total_days} 天")

# ==========================================
# 4. UI: 通路配置 (智慧連動邏輯)
# ==========================================
st.divider()
st.subheader("🛠️ 步驟 2：通路配置與配比")

# 先讓用戶選擇要啟用哪些通路
c_sel1, c_sel2, c_sel3 = st.columns(3)
with c_sel1: enable_fm = st.checkbox("全家便利商店", value=True)
with c_sel2: enable_fv = st.checkbox("全家新鮮視", value=True)
with c_sel3: enable_cf = st.checkbox("家樂福", value=False)

active_channels = []
if enable_fm: active_channels.append("FM")
if enable_fv: active_channels.append("FV")
if enable_cf: active_channels.append("CF")
active_count = len(active_channels)

# 初始化變數
pct_fm = 0
pct_fv = 0
pct_cf = 0
budget_fm = 0
budget_fv = 0
budget_cf = 0

# --- 智慧連動邏輯區 ---
st.markdown("---")
col_ui1, col_ui2, col_ui3 = st.columns(3)

# 情境 A: 剛好 2 個通路 (啟動 100% 連動)
if active_count == 2:
    first = active_channels[0]
    
    # 通路 1: 全家便利商店
    with col_ui1:
        if enable_fm:
            st.markdown("### 🏪 全家便利商店")
            if first == "FM":
                pct_fm = st.slider("預算佔比 %", 0, 100, 50, key="slider_fm_link")
