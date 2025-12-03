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
    "全家新鮮視": { # TV 螢幕 (已更正名稱)
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

st.set_page_config(page_title="媒體排程報價系統 v3.0", layout="wide")

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
st.title("📱 媒體報價系統 v3.0 (Pro)")

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
with c_sel3: enable_cf = st.checkbox("家樂福", value=False) # 預設關閉一個，測試連動

active_channels = []
if enable_fm: active_channels.append("FM")
if enable_fv: active_channels.append("FV")
if enable_cf: active_channels.append("CF")
active_count = len(active_channels)

# 初始化變數
pct_fm = 0
pct_fv = 0
pct_cf = 0

# --- 智慧連動邏輯區 ---
st.markdown("---")
col_ui1, col_ui2, col_ui3 = st.columns(3)

# 用於存放每個通路的預算結果，供下方秒數分配使用
budget_fm = 0
budget_fv = 0
budget_cf = 0

# 情境 A: 剛好 2 個通路 (啟動 100% 連動)
if active_count == 2:
    # 找出哪兩個是啟用的，將第一個設為 Slider，第二個自動計算
    first = active_channels[0]
    second = active_channels[1]
    
    # 為了 UI 美觀，我們還是渲染三個欄位，但只在對應的欄位顯示內容
    
    # 通路 1: 全家便利商店
    with col_ui1:
        if enable_fm:
            st.markdown("### 🏪 全家便利商店")
            if first == "FM":
                pct_fm = st.slider("預算佔比 %", 0, 100, 50, key="slider_fm_link")
            else:
                # 這是第二順位，自動計算
                pct_fm = 100 - (pct_fv if first == "FV" else pct_cf)
                st.progress(pct_fm / 100)
                st.write(f"自動連動佔比: **{pct_fm}%**")
            
            budget_fm = total_budget * (pct_fm / 100)
            st.info(f"預算: ${int(budget_fm):,}")

    # 通路 2: 全家新鮮視
    with col_ui2:
        if enable_fv:
            st.markdown("### 📺 全家新鮮視")
            if first == "FV":
                pct_fv = st.slider("預算佔比 %", 0, 100, 50, key="slider_fv_link")
            else:
                # 自動計算
                pct_fv = 100 - (pct_fm if first == "FM" else pct_cf)
                st.progress(pct_fv / 100)
                st.write(f"自動連動佔比: **{pct_fv}%**")
                
            budget_fv = total_budget * (pct_fv / 100)
            st.info(f"預算: ${int(budget_fv):,}")

    # 通路 3: 家樂福
    with col_ui3:
        if enable_cf:
            st.markdown("### 🛒 家樂福")
            if first == "CF":
                pct_cf = st.slider("預算佔比 %", 0, 100, 50, key="slider_cf_link")
            else:
                # 自動計算
                pct_cf = 100 - (pct_fm if first == "FM" else pct_fv)
                st.progress(pct_cf / 100)
                st.write(f"自動連動佔比: **{pct_cf}%**")
                
            budget_cf = total_budget * (pct_cf / 100)
            st.info(f"預算: ${int(budget_cf):,}")

# 情境 B: 1 個或 3 個通路 (手動模式 + 警示)
else:
    # 全家便利商店
    with col_ui1:
        if enable_fm:
            st.markdown("### 🏪 全家便利商店")
            pct_fm = st.slider("預算佔比 %", 0, 100, 50 if active_count==1 else 33, key="slider_fm_manual")
            budget_fm = total_budget * (pct_fm / 100)
            st.info(f"預算: ${int(budget_fm):,}")
    
    # 全家新鮮視
    with col_ui2:
        if enable_fv:
            st.markdown("### 📺 全家新鮮視")
            pct_fv = st.slider("預算佔比 %", 0, 100, 50 if active_count==1 else 33, key="slider_fv_manual")
            budget_fv = total_budget * (pct_fv / 100)
            st.info(f"預算: ${int(budget_fv):,}")

    # 家樂福
    with col_ui3:
        if enable_cf:
            st.markdown("### 🛒 家樂福")
            pct_cf = st.slider("預算佔比 %", 0, 100, 50 if active_count==1 else 34, key="slider_cf_manual")
            budget_cf = total_budget * (pct_cf / 100)
            st.info(f"預算: ${int(budget_cf):,}")

    # 檢查總和
    total_pct = pct_fm + pct_fv + pct_cf
    if total_pct != 100:
        if total_pct > 100:
            st.error(f"⚠️ 總佔比 {total_pct}% (超過 100%)，請調整。")
        else:
            st.warning(f"⚠️ 總佔比 {total_pct}% (不足 100%)，剩餘 {100-total_pct}% 未分配。")


# ==========================================
# 5. UI: 秒數混搭細節 (每個通路都可以任意混搭)
# ==========================================
rows = [] # 收集最後要產出的資料

def render_duration_mix_ui(channel_name, channel_key, total_channel_budget, program_name):
    """
    渲染通用的秒數混合介面
    channel_key: 用來區分不同通路的元件 ID
    """
    if total_channel_budget <= 0:
        return []
    
    generated_rows = []
    
    # 混搭開關
    is_mix = st.checkbox(f"開啟 {channel_name} 秒數混搭", key=f"mix_{channel_key}")
    
    if not is_mix:
        # 單一秒數模式
        sec = st.selectbox("選擇秒數", ["10s", "15s", "20s"], index=1, key=f"sec_s_{channel_key}")
        # 產生 1 筆資料 (100% 預算)
        r = calculate_single_schedule(channel_name, region, sec, total_channel_budget, start_date, end_date, program_name)
        if r: generated_rows.append(r)
        
    else:
        # 混搭模式 (Slots)
        st.markdown(f"**{channel_name} 組合配置:**")
        
        # Slot 1
        c_mix1, c_mix2 = st.columns([1, 1])
        with c_mix1:
            sec_1 = st.selectbox(f"組合 1 秒數", ["10s", "15s", "20s"], index=0, key=f"sec_m1_{channel_key}")
        with c_mix2:
            pct_1 = st.number_input(f"組合 1 佔該通路 %", 0, 100, 50, key=f"pct_m1_{channel_key}")
            
        # Slot 2 (自動計算剩餘)
        pct_2 = 100 - pct_1
        c_mix3, c_mix4 = st.columns([1, 1])
        with c_mix3:
            sec_2 = st.selectbox(f"組合 2 秒數", ["10s", "15s", "20s"], index=2, key=f"sec_m2_{channel_key}")
        with c_mix4:
            st.write(f"組合 2 佔該通路 %")
            st.caption(f"**{pct_2}%** (自動計算)")
        
        # 計算 Slot 1
        budget_1 = total_channel_budget * (pct_1 / 100)
        r1 = calculate_single_schedule(channel_name, region, sec_1, budget_1, start_date, end_date, program_name)
        if r1: generated_rows.append(r1)
        
        # 計算 Slot 2
        budget_2 = total_channel_budget * (pct_2 / 100)
        r2 = calculate_single_schedule(channel_name, region, sec_2, budget_2, start_date, end_date, program_name)
        if r2: generated_rows.append(r2)
        
    return generated_rows


# 依序渲染下方的詳細設定區
st.markdown("---")

# 全家便利商店 詳細設定
if enable_fm:
    with col_ui1:
        st.caption("🔻 秒數設定")
        new_rows = render_duration_mix_ui("全家便利商店", "fm", budget_fm, "通路廣播")
        rows.extend(new_rows)

# 全家新鮮視 詳細設定
if enable_fv:
    with col_ui2:
        st.caption("🔻 秒數
