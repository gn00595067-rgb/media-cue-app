import streamlit as st
import pandas as pd
from datetime import timedelta, date
import io

# ==========================================
# 0. 基礎函式定義
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
        # 單一模式 (這裡之前縮排錯了，現在修復)
        sec = st.selectbox("選擇秒數", ["10s", "15s", "20s"], index=1, key=f"s_{key_id}")
        r = calculate_single_schedule(channel_name, region, sec, budget, start_date, end_date, program_name)
        if r: result_rows.append(r)
    else:
        # 混搭模式
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            sec_1 = st.selectbox(f"組合1", ["10s", "15s", "20s"], index=0, key=f"sm1_{key_id}")
            pct_1 = st.number_input(f"組合1佔比%", 0, 100, 50, key=f"pm1_{key_id}")
        with col_m2:
            sec_2 = st.selectbox(f"組合2", ["10s", "15s", "20s"], index=2, key=f"sm2_{key_id}")
            pct_2 = 100 - pct_1
            st.write("組合2佔比%")
            st.info(f"{pct_2}%")
            
        # 計算
        b1 = budget * (pct_1 / 100)
        r1 = calculate_single_schedule(channel_name, region, sec_1, b1, start_date, end_date, program_name)
        if r1: result_rows.append(r1)
        
        b2 = budget * (pct_2 / 100)
        r2 = calculate_single_schedule(channel_name, region, sec_2, b2, start_date, end_date, program_name)
        if r2: result_rows.append(r2)
        
    return result_rows

# ==========================================
# 1. 頁面開始
# ==========================================
st.set_page_config(page_title="媒體排程系統 v4.0", layout="wide")
st.title("📱 媒體報價系統 v4.0")

# 步驟 1: 全域設定
with st.expander("🛠️ 步驟 1：基礎設定 (日期/預算)", expanded=True):
    c1, c2, c3 = st.columns(3)
    with c1: start_date = st.date_input("開始日期", value=date.today())
    with c2: end_date = st.date_input("結束日期", value=date.today() + timedelta(days=29))
    with c3: region = st.selectbox("投放區域", ["全省", "北部", "中部", "南部"])
    
    total_budget = st.number_input("總預算 (未稅)", value=500000, step=10000)
    total_days = (end_date - start_date).days + 1

# ==========================================
# 2. 通路與預算配置 (垂直分組佈局)
# ==========================================
st.divider()
st.subheader("🛠️ 步驟 2：通路配置")

# 通路啟用開關
sel_c1, sel_c2, sel_c3 = st.columns(3)
with sel_c1: enable_fm = st.checkbox("全家便利商店", value=True)
with sel_c2: enable_fv = st.checkbox("全家新鮮視", value=True)
with sel_c3: enable_cf = st.checkbox("家樂福", value=False)

# 偵測啟用數量以判斷是否連動
active_channels = []
if enable_fm: active_channels.append("FM")
if enable_fv: active_channels.append("FV")
if enable_cf: active_channels.append("CF")
active_count = len(active_channels)

# 準備容器收集結果
all_schedule_rows = []

# 建立三欄佈局 (所有設定都在各自的欄位內完成)
layout_c1, layout_c2, layout_c3 = st.columns(3)

# 初始化變數，確保不為空
pct_fm = 0
pct_fv = 0
pct_cf = 0

# ----------------------------
# 第一欄：全家便利商店 (FM)
# ----------------------------
with layout_c1:
    if enable_fm:
        st.info("🏪 全家便利商店")
        # 預算邏輯
        if active_count == 2 and active_channels[0] != "FM":
             st.caption("自動連動計算中...")
             pct_fm = 0 
        else:
             pct_fm = st.slider("預算佔比 %", 0, 100, 50 if active_count==1 else 33, key="sl_fm")
        
        budget_fm_placeholder = st.empty()

# ----------------------------
# 第二欄：全家新鮮視 (FV)
# ----------------------------
with layout_c2:
    if enable_fv:
        st.info("📺 全家新鮮視")
        if active_count == 2 and active_channels[0] == "FM" and active_channels[1] == "FV":
             pct_fv = 100 - pct_fm
             st.progress(pct_fv/100)
             st.write(f"連動佔比: **{pct_fv}%**")
        elif active_count == 2 and active_channels[0] != "FV":
             pct_fv = 0 
        else:
             pct_fv = st.slider("預算佔比 %", 0, 100, 50 if active_count==1 else 33, key="sl_fv")

        budget_fv_placeholder = st.empty()

# ----------------------------
# 第三欄：家樂福 (CF)
# ----------------------------
with layout_c3:
    if enable_cf:
        st.info("🛒 家樂福")
        if active_count == 2:
             leader_pct = pct_fm if enable_fm else pct_fv
             pct_cf = 100 - leader_pct
             st.progress(pct_cf/100)
             st.write(f"連動佔比: **{pct_cf}%**")
        else:
             pct_cf = st.slider("預算佔比 %", 0, 100, 50 if active_count==1 else 34, key="sl_cf")
        
        budget_cf_placeholder = st.empty()

# ----------------------------
# 計算並渲染
# ----------------------------
budget_fm = total_budget * (pct_fm / 100) if enable_fm else 0
budget_fv = total_budget * (pct_fv / 100) if enable_fv else 0
budget_cf = total_budget * (pct_cf / 100) if enable_cf else 0

if enable_fm: 
    with layout_c1: 
        budget_fm_placeholder.markdown(f"**${int(budget_fm):,}**")
        rows = render_mix_ui("全家便利商店", "fm", budget_fm, region, start_date, end_date, "通路廣播")
        all_schedule_rows.extend(rows)

if enable_fv: 
    with layout_c2: 
        budget_fv_placeholder.markdown(f"**${int(budget_fv):,}**")
        rows = render_mix_ui("全家新鮮視", "fv", budget_fv, region, start_date, end_date, "新鮮視TV")
        all_schedule_rows.extend(rows)

if enable_cf: 
    with layout_c3: 
        budget_cf_placeholder.markdown(f"**${int(budget_cf):,}**")
        rows = render_mix_ui("家樂福", "cf", budget_cf, region, start_date, end_date, "家樂福聯播")
        all_schedule_rows.extend(rows)

# 總和檢查
current_total = 0
if enable_fm: current_total += pct_fm
if enable_fv: current_total += pct_fv
if enable_cf: current_total += pct_cf

if active_count != 2 and current_total != 100:
    st.warning(f"⚠️ 目前總佔比 {current_total}% (建議調整為 100%)")

# ==========================================
# 3. 產出報表
# ==========================================
if not all_schedule_rows:
    st.divider()
    st.warning("⚠️ 請至少啟用一個通路")
else:
    # 建立 DataFrame
    date_headers = []
    curr = start_date
    for _ in range(total_days):
        wd = WEEKDAY_MAP[curr.weekday()]
        date_headers.append(f"{curr.strftime('%m/%d')} ({wd})")
        curr += timedelta(days=1)

    final_data = []
    for r in all_schedule_rows:
        base_info = {
            "Station": r["Station"],
            "Location": r["Location"],
            "Program": r["Program"],
            "Day-part": r["Day-part"],
            "Size": r["Size"],
            "Rate (Net)": r["Rate (Net)"],
            "Package Cost": r["Package Cost"],
        }
        for idx, spots in enumerate(r["Schedule"]):
            base_info[date_headers[idx]] = spots
        base_info["總檔次"] = r["Total Spots"]
        final_data.append(base_info)

    df = pd.DataFrame(final_data)

    # Total Row
    sum_row = df.sum(numeric_only=True)
    sum_row["Station"] = "Total"
    sum_row["Rate (Net)"] = ""
    df_display = pd.concat([df, pd.DataFrame([sum_row])], ignore_index=True).fillna("")

    st.divider()
    st.subheader("📊 試算結果 Cue 表")
    st.dataframe(df_display, use_container_width=True)

    # 下載按鈕
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_display.to_excel(writer, sheet_name='Cue表', index=False)
        wb = writer.book
        ws = writer.sheets['Cue表']
        fmt = wb.add_format({'bold': True, 'fg_color': '#4F81BD', 'font_color': 'white', 'border': 1})
        for c, val in enumerate(df_display.columns.values):
            ws.write(0, c, val, fmt)
        ws.set_column(0, 0, 15)
        ws.set_column(1, 6, 10)
        ws.set_column(7, len(df_display.columns)-1, 6)
    
    output.seek(0)
    st.download_button("📥 下載 Excel", data=output, file_name=f"Schedule_{start_date}.xlsx")
