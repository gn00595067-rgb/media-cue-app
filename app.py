import streamlit as st
import pandas as pd
from datetime import timedelta, date
import io
import requests

# ==========================================
# 1. 系統設定 (費率卡 Rate Card)
# ==========================================
# 為了方便維護，您可以隨時回來修改這裡的價格
RATE_CARD = {
    "全家": {
        "全省": {"10s": 150, "15s": 200, "20s": 260},
        "北部": {"10s": 180, "15s": 240, "20s": 310},
        "中部": {"10s": 150, "15s": 200, "20s": 260}, # 假設
        "南部": {"10s": 150, "15s": 200, "20s": 260}, # 假設
    },
    "家樂福": {
        "全省": {"10s": 130, "15s": 180, "20s": 230},
        "北部": {"10s": 160, "15s": 210, "20s": 260}, # 假設
        "中部": {"10s": 130, "15s": 180, "20s": 230},
        "南部": {"10s": 130, "15s": 180, "20s": 230},
    }
}

# 設定頁面 (手機版面會自動適應)
st.set_page_config(page_title="媒體排程報價系統", layout="centered")

# ==========================================
# 2. 業務輸入介面 (UI)
# ==========================================
st.title("📱 媒體報價系統")
st.info("請輸入條件，下方自動生成 Cue 表")

# A. 基礎條件
with st.expander("1. 基礎設定 (日期/預算)", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("開始日期", value=date.today())
    with col2:
        end_date = st.date_input("結束日期", value=date.today() + timedelta(days=29))
    
    total_days = (end_date - start_date).days + 1
    st.caption(f"📅 總走期天數：{total_days} 天")

    total_budget = st.number_input("總預算 (未稅)", min_value=10000, value=500000, step=10000, format="%d")
    region = st.selectbox("投放區域", ["全省", "北部", "中部", "南部"])

# B. 複雜配置
with st.expander("2. 通路配置 (全家/家樂福)", expanded=True):
    st.subheader("🏪 全家便利商店")
    fm_ratio = st.slider("全家佔總預算 %", 0, 100, 50)
    
    c1, c2 = st.columns(2)
    with c1:
        fm_sec_1 = st.selectbox("組合1 秒數", ["10s", "15s", "20s"], index=0)
        fm_ratio_1 = st.number_input("組合1 佔全家 %", 0, 100, 20)
    with c2:
        fm_sec_2 = st.selectbox("組合2 秒數", ["10s", "15s", "20s"], index=2)
        st.write(f"組合2 佔全家 % : **{100 - fm_ratio_1}%**")
    
    st.divider()
    
    st.subheader("🛒 家樂福")
    carrefour_ratio = 100 - fm_ratio
    st.write(f"家樂福佔總預算 % : **{carrefour_ratio}%**")
    cf_sec = st.selectbox("家樂福 秒數", ["10s", "15s", "20s"], index=1)

# ==========================================
# 3. 核心運算邏輯
# ==========================================
def calculate_row(channel, region, sec, budget, s_date, e_date):
    # 1. 查價
    try:
        rate = RATE_CARD[channel][region][sec]
    except:
        rate = 200 # 預設防呆
    
    # 2. 算總檔次
    total_spots = int(budget / rate) if rate > 0 else 0
    
    # 3. 每日分配 (平均分配 + 餘數填補)
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
        "Program": f"{channel}聯播網",
        "Day-part": "06-24",
        "Size": sec,
        "Rate (Net)": rate,
        "Package Cost": int(budget),
        "Schedule": schedule, # List of daily spots
        "Total Spots": total_spots
    }

# 開始計算三筆資料
# 1. 全家 A
budget_fm_total = total_budget * (fm_ratio / 100)
budget_fm_1 = budget_fm_total * (fm_ratio_1 / 100)
row1 = calculate_row("全家", region, fm_sec_1, budget_fm_1, start_date, end_date)

# 2. 全家 B
budget_fm_2 = budget_fm_total * ((100 - fm_ratio_1) / 100)
row2 = calculate_row("全家", region, fm_sec_2, budget_fm_2, start_date, end_date)

# 3. 家樂福
budget_cf = total_budget * (carrefour_ratio / 100)
row3 = calculate_row("家樂福", region, cf_sec, budget_cf, start_date, end_date)

# ==========================================
# 4. 建立 DataFrame 表格
# ==========================================
# 產生日期標頭
date_headers = []
curr = start_date
for _ in range(total_days):
    date_headers.append(curr.strftime("%m/%d"))
    curr += timedelta(days=1)

# 組合資料
data_rows = [row1, row2, row3]
final_data = []

for r in data_rows:
    base_info = {
        "Station": r["Station"],
        "Location": r["Location"],
        "Program": r["Program"],
        "Day-part": r["Day-part"],
        "Size": r["Size"],
        "Rate (Net)": r["Rate (Net)"],
        "Package Cost": r["Package Cost"],
    }
    # 把每天的檔次攤平成欄位
    for idx, spots in enumerate(r["Schedule"]):
        col_name = date_headers[idx]
        base_info[col_name] = spots
        
    base_info["總檔次"] = r["Total Spots"]
    final_data.append(base_info)

# 轉成 Pandas DataFrame
df = pd.DataFrame(final_data)

# 計算 Total Row
sum_row = df.sum(numeric_only=True)
sum_row["Station"] = "Total"
# 修正 Rate 等不需要加總的欄位
sum_row["Rate (Net)"] = "" 
sum_df = pd.DataFrame([sum_row])
df_display = pd.concat([df, sum_df], ignore_index=True)
df_display = df_display.fillna("") # 把 NaN 補空值

# ==========================================
# 5. 顯示結果與下載
# ==========================================
st.divider()
st.subheader("📊 試算結果 Cue 表")
st.dataframe(df_display, use_container_width=True)

# --- 功能 A: 下載 Excel ---
# 寫入 BytesIO 緩衝區
buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
    df_display.to_excel(writer, sheet_name='Cue表', index=False)
    
    # 簡單美化 Excel 寬度
    worksheet = writer.sheets['Cue表']
    for i, col in enumerate(df_display.columns):
        worksheet.set_column(i, i, 12) # 設定欄寬

    writer.close()

st.download_button(
    label="📥 下載 Excel 報表 (傳給客戶)",
    data=buffer,
    file_name=f"報價單_{start_date}_{total_budget}.xlsx",
    mime="application/vnd.ms-excel"
)

# --- 功能 B: 上傳 Ragic ---
st.divider()
with st.expander("☁️ 進階：上傳至 Ragic"):
    st.write("確認無誤後，點擊按鈕直接存入系統。")
    
    # 這裡請換成您的 Ragic API URL
    # 格式通常是: https://www.ragic.com/你的帳號/你的頁籤/表單ID?api=true
    ragic_url = st.text_input("Ragic API URL", "https://www.ragic.com/demo/sales/1?api=true")
    ragic_key = st.text_input("API Key", type="password")
    
    if st.button("🚀 確認開單並上傳"):
        if not ragic_key:
            st.error("請輸入 API Key")
        else:
            # 整理 payload
            payload = {
                "10001": str(start_date),       # 對應您的 Ragic 欄位 ID
                "10002": str(total_budget),     # 對應您的 Ragic 欄位 ID
                # 子表格資料通常比較複雜，這裡僅做簡單示範
                # 實務上要根據您的 Ragic 子表格 ID 結構來組 JSON
            }
            
            # 模擬發送 (您可以解開下面註解來真正發送)
            # resp = requests.post(ragic_url, json=payload, headers={"Authorization": "Basic " + ragic_key})
            
            st.success("✅ 已發送資料至 Ragic！(模擬)")
            st.json(payload)
