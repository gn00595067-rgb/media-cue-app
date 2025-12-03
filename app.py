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

def render_mix_ui_v2(channel_name, key_id, budget, region, start_date, end_date, program_name):
    """
    新版混搭介面：支援 3 種秒數自由選
    """
    if budget <= 0: return []
    
    st.markdown("---")
    st.caption(f"🔻 {channel_name} 秒數與佔比")

    # 1. 讓用戶勾選要用的秒數
    cols_chk = st.columns(3)
    with cols_chk[0]: use_10 = st.checkbox("10s", value=False, key=f"c10_{key_id}")
    with cols_chk[1]: use_15 = st.checkbox("15s", value=True, key=f"c15_{key_id}")
    with cols_chk[2]: use_20 = st.checkbox("20s", value=False, key=f"c20_{key_id}")
    
    # 建立選取的秒數清單
    active_secs = []
    if use_10: active_secs.append("10s")
    if use_15: active_secs.append("15s")
    if use_20: active_secs.append("20s")
    
    count = len(active_secs)
    pcts = {} # 存放結果 { "10s": 50, "20s": 50 }

    # 2. 根據勾選數量決定介面邏輯
    if count == 0:
        st.warning("請至少勾選一種秒數")
        return []
        
    elif count == 1:
        # 單一秒數 -> 自動 100%
        sec = active_secs[0]
        pcts[sec] = 100
        st.info(f"✅ {sec} 佔比: 100%")
        
    elif count == 2:
        # 兩個秒數 -> 自動連動
        sec_a, sec_b = active_secs[0], active_secs[1]
        val_a = st.slider(f"{sec_a} 佔比", 0, 100, 50, key=f"sl2_{key_id}")
        val_b = 100 - val_a
        
        pcts[sec_a] = val_a
        pcts[sec_b] = val_b
        
        # 顯示連動結果
        st.write(f"{sec_b} 自動連動: **{val_b}%**")
        st.progress(val_b/100)
        
    elif count == 3:
        # 三個秒數 -> 手動輸入 + 警示
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

    # 3. 根據佔比計算結果
    result_rows = []
    for sec, pct in pcts.items():
        if pct > 0:
            sub_budget = budget * (pct / 100)
            r = calculate_single_schedule(channel_name, region, sec, sub_budget, start_date, end_date, program_name)
            if r: result_rows.append(r)
            
    return result_rows

# ==========================================
# 1. 頁面開始
# ==========================================
st.set_page_config(page_title="媒體排程系統 v5.0", layout="wide")
st.title("📱 媒體報價系統 v5.0")

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

pct_fm = 0
pct_fv = 0
pct_cf = 0

# --- 第一欄：全家便利商店 ---
with layout_c1:
    if enable_fm:
        st.info("🏪 全家便利商店")
        if active_count == 2 and active_channels[0] != "FM":
             st.caption("自動連動計算中...")
             pct_fm = 0 
        else:
             pct_fm = st.slider("預算佔比 %", 0, 100, 50 if active_count==1 else 33, key="sl_fm")
        budget_fm_placeholder = st.empty()

# --- 第二欄：全家新鮮視 ---
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

# --- 第三欄：家樂福 ---
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

# --- 計算與渲染 ---
budget_fm = total_budget * (pct_fm / 100) if enable_fm else 0
budget_fv = total_budget * (pct_fv / 100) if enable_fv else 0
budget_cf = total_budget * (pct_cf / 100) if enable_cf else 0

if enable_fm: 
    with layout_c1: 
        budget_fm_placeholder.markdown(f"**${int(budget_fm):,}**")
        # 使用新版 V2 混搭介面
        rows = render_mix_ui_v2("全家便利商店", "fm", budget_fm, region, start_date, end_date, "通路廣播")
        all_schedule_rows.extend(rows)

if enable_fv: 
    with layout_c2: 
        budget_fv_placeholder.markdown(f"**${int(budget_fv):,}**")
        rows = render_mix_ui_v2("全家新鮮視", "fv", budget_fv, region, start_date, end_date, "新鮮視TV")
        all_schedule_rows.extend(rows)

if enable_cf: 
    with layout_c3: 
        budget_cf_placeholder.markdown(f"**${int(budget_cf):,}**")
        rows = render_mix_ui_v2("家樂福", "cf", budget_cf, region, start_date, end_date, "家樂福聯播")
        all_schedule_rows.extend(rows)

# 總和檢查
current_total = 0
if enable_fm: current_total += pct_fm
if enable_fv: current_total += pct_fv
if enable_cf: current_total += pct_cf
if active_count != 2 and current_total != 100:
    st.warning(f"⚠️ 目前通路總佔比 {current_total}% (建議調整為 100%)")

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

    # ==========================================
    # Excel 下載邏輯優化
    # ==========================================
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
    
    # 檔名轉換為字串，確保相容性
    filename = f"Schedule_{start_date}.xlsx"
    
    # 下載按鈕
    st.download_button(
        label="📥 下載 Excel 報表",
        data=output,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="btn_download_excel" # 加入 key 確保狀態唯一
    )
    
    # 針對手機用戶的提示
    st.caption("ℹ️ 手機用戶請注意：若點擊下載無反應，請嘗試點選右上角選單 > 「以瀏覽器開啟」(Open in Browser)，避免使用 Line 內建瀏覽器。")
