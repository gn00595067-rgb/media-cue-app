import streamlit as st
import pandas as pd
from datetime import timedelta, date
import io

# ==========================================
# 0. 基礎函式定義 (放在最前面確保不報錯)
# ==========================================

# 費率卡
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
    """計算單一條目(Row)的排程"""
    if budget <= 0: return None
    
    # 查價
    try:
        rate = RATE_CARD.get(channel, {}).get(region, {}).get(sec, 0)
    except:
        rate = 0
    if rate == 0: return None
    
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

def render_mix_ui(channel_name, key_id, budget, region, start_date, end_date, program_name):
    """渲染秒數混搭介面並回傳計算結果"""
    if budget <= 0: return []
    
    result_rows = []
    st.markdown("---")
    st.caption(f"🔻 {channel_name} 秒數配置")
    
    is_mix = st.checkbox(f"開啟混搭 ({channel_name})", key=f"mix_{key_id}")
    
    if not is_mix:
