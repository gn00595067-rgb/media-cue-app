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
    "start_date": 10003,
    "end_date": 10004,
    "region": 10005,
    "total_budget": 10006,
    # 子表格 ID
    "sub_station": 20001,
    "sub_sec": 20002,
    "sub_rate": 20003,
    "sub_cost": 20004,
    "sub_spots": 20005,
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
    # 新增客戶名稱欄位
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

budget_fm = total_budget * (pct_fm / 100) if enable_fm else 0
budget_fv = total_budget * (pct_fv / 100) if enable_fv else 0
budget_cf = total_budget * (pct_cf / 100) if enable_cf else 0

if enable_fm: 
    with layout_c1: 
        budget_fm_placeholder.markdown(f"**${int(budget_fm):,}**")
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

current_total = 0
if enable_fm: current_total += pct_fm
if enable_fv: current_total += pct_fv
if enable_cf: current_total += pct_cf
if active_count != 2 and current_total != 100:
    st.warning(f"⚠️ 目前通路總佔比 {current_total}% (建議調整為 100%)")

# ==========================================
# 3. 產出與上傳區
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
    # 收集使用到的秒數和通路，用於 Excel 表頭
    used_secs = set()
    used_channels = set()

    for r in all_schedule_rows:
        used_secs.add(r["Size"])
        used_channels.add(r["Station"])
        
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
    sum_row = df.sum(numeric_only=True)
    sum_row["Station"] = "Total"
    sum_row["Rate (Net)"] = ""
    df_display = pd.concat([df, pd.DataFrame([sum_row])], ignore_index=True).fillna("")

    st.divider()
    st.subheader("📊 試算結果 Cue 表")
    st.dataframe(df_display, use_container_width=True)

    # ---------------------------
    # 下載區塊 (Excel 排版優化)
    # ---------------------------
    st.markdown("### 📥 匯出資料")
    col_dl, col_ragic = st.columns([1, 1])
    
    # 準備 Excel 表頭文字
    str_product = " ".join(sorted(used_secs)) # e.g., "10s 15s"
    str_period = f"{start_date.strftime('%Y.%m.%d')} - {end_date.strftime('%Y.%m.%d')}"
    str_medium = " ".join(sorted(used_channels)) # e.g., "全家便利商店 家樂福"
    if not client_name: client_name = ""

    with col_dl:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # 我們將表格寫入，但從第 8 列開始 (startrow=7)，留位置給表頭
            df_display.to_excel(writer, sheet_name='Cue表', index=False, startrow=7)
            
            wb = writer.book
            ws = writer.sheets['Cue表']
            
            # --- 樣式設定 ---
            fmt_title = wb.add_format({'bold': True, 'font_size': 14, 'align': 'left'})
            fmt_label = wb.add_format({'bold': True, 'align': 'right'})
            fmt_text = wb.add_format({'align': 'left'})
            fmt_header = wb.add_format({'bold': True, 'fg_color': '#4F81BD', 'font_color': 'white', 'border': 1})
            
            # --- 寫入上方資訊 (Row 1-6) ---
            ws.write('A1', 'Media Schedule', fmt_title)
            
            # 客戶名稱
            ws.write('A3', '客戶名稱：', fmt_label)
            ws.write('B3', client_name, fmt_text)
            
            # Product
            ws.write('A4', 'Product：', fmt_label)
            ws.write('B4', str_product, fmt_text)
            
            # Period
            ws.write('A5', 'Period：', fmt_label)
            ws.write('B5', str_period, fmt_text)
            
            # Medium
            ws.write('A6', 'Medium：', fmt_label)
            ws.write('B6', str_medium, fmt_text)
            
            # --- 寫入表格 Header 樣式 (Row 8) ---
            # 因為 to_excel 已經寫了資料，我們覆蓋 Header 的格式
            for c, val in enumerate(df_display.columns.values):
                ws.write(7, c, val, fmt_header)
            
            # --- 欄寬調整 ---
            ws.set_column(0, 0, 18) # Station
            ws.set_column(1, 6, 11) # Info
            ws.set_column(7, len(df_display.columns)-1, 6) # Dates
            
        output.seek(0)
        
        st.download_button(
            label="下載 Excel 報表",
            data=output,
            file_name=f"Schedule_{client_name}_{start_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.caption("手機請用瀏覽器開啟 (Chrome/Safari) 以確保下載成功")

    # ---------------------------
    # Ragic 上傳區塊
    # ---------------------------
    with col_ragic:
        with st.popover("☁️ 上傳至 Ragic"):
            st.markdown("#### 系統連線設定")
            ragic_url = st.text_input("API URL", placeholder="https://www.ragic.com/...")
            ragic_key = st.text_input("API Key", type="password")
            
            if st.button("確認上傳", type="primary", use_container_width=True):
                if not ragic_url or not ragic_key:
                    st.error("請輸入 API URL 與 Key")
                else:
                    # 1. 表頭
                    payload = {
                        RAGIC_CONFIG["client_name"]: client_name,
                        RAGIC_CONFIG["start_date"]: str(start_date),
                        RAGIC_CONFIG["end_date"]: str(end_date),
                        RAGIC_CONFIG["region"]: region,
                        RAGIC_CONFIG["total_budget"]: total_budget,
                        # 建議您在 Ragic 新增這三個欄位來接收這些資訊
                        # "100xx": client_name, 
                        # "100xx": str_product,
                        # "100xx": str_medium,
                    }
                    
                    # 2. 子表格
                    subtable_data = {}
                    for idx, r in enumerate(all_schedule_rows):
                        row_key = str((idx + 1) * -1)
                        subtable_data[row_key] = {
                            RAGIC_CONFIG["sub_station"]: r["Station"],
                            RAGIC_CONFIG["sub_sec"]: r["Size"],
                            RAGIC_CONFIG["sub_rate"]: r["Rate (Net)"],
                            RAGIC_CONFIG["sub_cost"]: r["Package Cost"],
                            RAGIC_CONFIG["sub_spots"]: r["Total Spots"]
                        }
                    
                    subtable_key = f"_subtable_{RAGIC_CONFIG['sub_station']}" 
                    payload[subtable_key] = subtable_data
                    
                    st.info("正在連線...")
                    try:
                        if "?api" not in ragic_url:
                            ragic_url += "?api=true" if "?" not in ragic_url else "&api=true"
                            
                        resp = requests.post(
                            ragic_url, 
                            json=payload, 
                            headers={"Authorization": "Basic " + ragic_key}
                        )
                        
                        if resp.status_code == 200:
                            res_json = resp.json()
                            if res_json.get("status") == "SUCCESS":
                                st.success(f"✅ 上傳成功！ID: {res_json.get('ragicId')}")
                            else:
                                st.error(f"上傳失敗: {res_json.get('msg')}")
                                st.json(res_json)
                        else:
                            st.error(f"連線錯誤: {resp.status_code}")
                            st.write(resp.text)
                    except Exception as e:
                        st.error(f"發生錯誤: {e}")
