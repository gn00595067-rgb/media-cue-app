import streamlit as st
import pandas as pd
from datetime import timedelta, date
import io
import requests

# ==========================================
# 1. 系統設定 (費率卡 Rate Card)
# ==========================================
# 請根據實際情況調整價格
RATE_CARD = {
    "全家便利商店": { # 通常指音訊廣播
        "全省": {"10s": 150, "15s": 200, "20s": 260},
        "北部": {"10s": 180, "15s": 240, "20s": 310},
        "中部": {"10s": 150, "15s": 200, "20s": 260},
        "南部": {"10s": 150, "15s": 200, "20s": 260},
    },
    "全家新鮮市": { # 通常指 TV 螢幕
        "全省": {"10s": 400, "15s": 500, "20s": 600}, # 模擬價格
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

# 星期幾的中文對照
WEEKDAY_MAP = {0: "一", 1: "二", 2: "三", 3: "四", 4: "五", 5: "六", 6: "日"}

st.set_page_config(page_title="媒體排程報價系統", layout="wide") # 改成寬版顯示

# ==========================================
# 2. 側邊欄與上方設定 (UI)
# ==========================================
st.title("📱 媒體報價系統 v2.0")

# 放在 Expander 讓手機畫面不要太長
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
    st.caption(f"📅 總走期：{total_days} 天 ({start_date} ~ {end_date})")

# ==========================================
# 3. 三大通路配置 (UI)
# ==========================================
st.divider()
st.subheader("🛠️ 步驟 2：通路配置")

# 用 Tabs 或是 Columns 來分開設定，這裡用 Columns 比較直觀
c1, c2, c3 = st.columns(3)

# --- 通路 1: 全家便利商店 ---
with c1:
    st.markdown("### 🏪 全家便利商店")
    enable_fm_store = st.checkbox("啟用", value=True, key="cb_fm_s")
    if enable_fm_store:
        pct_fm_store = st.slider("預算佔比 %", 0, 100, 30, key="sl_fm_s")
        sec_fm_store = st.selectbox("廣告秒數", ["10s", "15s", "20s"], index=1, key="sb_fm_s")
        cost_fm_store = total_budget * (pct_fm_store / 100)
        st.info(f"預算: ${int(cost_fm_store):,}")
    else:
        cost_fm_store = 0
        pct_fm_store = 0
        sec_fm_store = "15s"

# --- 通路 2: 全家新鮮市 ---
with c2:
    st.markdown("### 📺 全家新鮮市")
    enable_fm_fresh = st.checkbox("啟用", value=True, key="cb_fm_f")
    if enable_fm_fresh:
        pct_fm_fresh = st.slider("預算佔比 %", 0, 100, 30, key="sl_fm_f")
        sec_fm_fresh = st.selectbox("廣告秒數", ["10s", "15s", "20s"], index=0, key="sb_fm_f")
        cost_fm_fresh = total_budget * (pct_fm_fresh / 100)
        st.info(f"預算: ${int(cost_fm_fresh):,}")
    else:
        cost_fm_fresh = 0
        pct_fm_fresh = 0
        sec_fm_fresh = "10s"

# --- 通路 3: 家樂福 ---
with c3:
    st.markdown("### 🛒 家樂福")
    enable_carrefour = st.checkbox("啟用", value=True, key="cb_cf")
    if enable_carrefour:
        # 自動計算剩餘建議值，但不強制
        remain_pct = max(0, 100 - pct_fm_store - pct_fm_fresh)
        pct_carrefour = st.slider("預算佔比 %", 0, 100, remain_pct, key="sl_cf")
        sec_carrefour = st.selectbox("廣告秒數", ["10s", "15s", "20s"], index=1, key="sb_cf")
        cost_carrefour = total_budget * (pct_carrefour / 100)
        
        # 檢查總和
        total_pct = pct_fm_store + pct_fm_fresh + pct_carrefour
        if total_pct > 100:
            st.warning(f"⚠️ 注意：總佔比已達 {total_pct}%，超過 100%")
        st.info(f"預算: ${int(cost_carrefour):,}")
    else:
        cost_carrefour = 0
        sec_carrefour = "15s"

# ==========================================
# 4. 核心運算邏輯
# ==========================================
def calculate_row(channel, region, sec, budget, s_date, e_date, program_name):
    if budget <= 0:
        return None
        
    # 1. 查價
    try:
        rate = RATE_CARD.get(channel, {}).get(region, {}).get(sec, 0)
    except:
        rate = 0
    
    # 防呆
    if rate == 0:
        return None
    
    # 2. 算總檔次
    total_spots = int(budget / rate)
    
    # 3. 每日分配
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
        "Day-part": "06-24", # 預設全天
        "Size": sec,
        "Rate (Net)": rate,
        "Package Cost": int(budget),
        "Schedule": schedule,
        "Total Spots": total_spots
    }

# 收集資料
rows = []

if enable_fm_store:
    r = calculate_row("全家便利商店", region, sec_fm_store, cost_fm_store, start_date, end_date, "通路廣播")
    if r: rows.append(r)

if enable_fm_fresh:
    r = calculate_row("全家新鮮市", region, sec_fm_fresh, cost_fm_fresh, start_date, end_date, "新鮮視TV")
    if r: rows.append(r)

if enable_carrefour:
    r = calculate_row("家樂福", region, sec_carrefour, cost_carrefour, start_date, end_date, "家樂福聯播")
    if r: rows.append(r)

# ==========================================
# 5. 建立表格與顯示
# ==========================================
if not rows:
    st.warning("尚未配置任何預算，請開啟上方通路開關。")
else:
    # 產生日期標頭 (含星期) e.g., "10/01 (三)"
    date_headers = []
    curr = start_date
    for _ in range(total_days):
        wd = WEEKDAY_MAP[curr.weekday()]
        date_str = f"{curr.strftime('%m/%d')} ({wd})"
        date_headers.append(date_str)
        curr += timedelta(days=1)

    # 轉成 DataFrame
    final_data = []
    for r in rows:
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
            col_name = date_headers[idx]
            base_info[col_name] = spots
        base_info["總檔次"] = r["Total Spots"]
        final_data.append(base_info)

    df = pd.DataFrame(final_data)

    # 計算 Total
    sum_row = df.sum(numeric_only=True)
    sum_row["Station"] = "Total"
    sum_row["Rate (Net)"] = ""
    sum_df = pd.DataFrame([sum_row])
    df_display = pd.concat([df, sum_df], ignore_index=True)
    df_display = df_display.fillna("")

    # 顯示
    st.divider()
    st.subheader("📊 試算結果 Cue 表")
    st.dataframe(df_display, use_container_width=True)

    # ==========================================
    # 6. Excel 下載 (修復版)
    # ==========================================
    # 使用 BytesIO 確保記憶體寫入
    output = io.BytesIO()
    
    # 建立 Excel Writer
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_display.to_excel(writer, sheet_name='Cue表', index=False)
        
        # 取得 workbook 和 worksheet 物件來進行格式設定
        workbook = writer.book
        worksheet = writer.sheets['Cue表']
        
        # 設定格式
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        
        # 設定欄寬
        worksheet.set_column(0, 0, 15) # Station
        worksheet.set_column(1, 4, 10) # Info columns
        worksheet.set_column(5, 6, 12) # Price columns
        worksheet.set_column(7, len(df_display.columns)-1, 5) # Date columns 窄一點
        
    writer.close()
    
    # 重要的修復：將指標移回開頭，不然下載的檔案會是空的
    output.seek(0)
    
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        st.download_button(
            label="📥 下載 Excel 報表",
            data=output,
            file_name=f"MediaSchedule_{start_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    # ==========================================
    # 7. 上傳 Ragic
    # ==========================================
    with col_d2:
        with st.popover("☁️ 上傳至 Ragic"):
            ragic_url = st.text_input("API URL", placeholder="https://www.ragic.com/...")
            ragic_key = st.text_input("API Key", type="password")
            if st.button("確認上傳"):
                st.success("資料已送出 (模擬)")
