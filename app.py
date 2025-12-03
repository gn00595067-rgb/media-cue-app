import streamlit as st
import pandas as pd
from datetime import timedelta, date
import io
import requests
import json

# ==========================================
# 0. 系統預設值 (已內建您的 Ragic 資訊)
# ==========================================
DEFAULT_RAGIC_URL = "https://ap15.ragic.com/liuskyo/media-quotation/2"
DEFAULT_API_KEY = "L04zZGhrVmtTV3pqN1VLbUpnOFZMZ0NvaEJ6RlRUd1pjOEtDZ3lmSXl1RW8wcUJPZ2pSbWdZYmpHQUp2R1VJOA=="

# Ragic 欄位 ID 設定 (請確認這些 ID 與您 Ragic 表單一致)
RAGIC_CONFIG = {
    "client_name": 10012,
    "start_date": 10013,
    "end_date": 10014,
    "region": 10015,
    "total_budget": 10016,
    "file_upload": 10022, # 檔案上傳欄位 ID
    
    # 子表格欄位
    "sub_station": 10017,
    "sub_sec": 10018,
    "sub_rate": 10019,
    "sub_cost": 10020,
    "sub_spots": 10021,
}

# 子表格 Root ID
SUBTABLE_ROOT_ID = "1000076"

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
st.set_page_config(page_title="媒體排程系統 v9.0", layout="wide")
st.title("📱 媒體報價系統 v9.0")

with st.expander("🛠️ 步驟 1：基礎資訊", expanded=True):
    client_name = st.text_input("客戶名稱", placeholder="例如：台灣讀廣")
    st.divider()
    c1, c2, c3 = st.columns(3)
    with c1: start_date = st.date_input("開始日期", value=date.today())
    with c2: end_date = st.date_input("結束日期", value=date.today() + timedelta(days=29))
    with c3: region = st.selectbox("投放區域", ["全省", "北部", "中部", "南部"])
    total_budget = st.number_input("總預算 (未稅)", value=500000, step=10000)
    total_days = (end_date - start_date).days + 1

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

# STEP 3: OUTPUT
if not all_schedule_rows:
    st.divider()
    st.warning("⚠️ 請至少啟用一個通路")
else:
    # Build DataFrame
    date_headers = []
    curr = start_date
    for _ in range(total_days):
        wd = WEEKDAY_MAP[curr.weekday()]
        date_headers.append(f"{curr.strftime('%m/%d')} ({wd})")
        curr += timedelta(days=1)

    final_data = []
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

    # ==========================
    # 新增：網頁版資訊顯示區
    # ==========================
    str_product = " ".join(sorted(used_secs))
    str_period = f"{start_date.strftime('%Y.%m.%d')} - {end_date.strftime('%Y.%m.%d')}"
    str_medium = " ".join(sorted(used_channels))
    if not client_name: client_name = "(未填寫)"

    st.info(f"""
    **客戶名稱：** {client_name}  
    **Product：** {str_product}  
    **Period：** {str_period}  
    **Medium：** {str_medium}
    """)
    # ==========================

    st.dataframe(df_display, use_container_width=True)

    # ---------------------------
    # 準備 Excel 資料 (共用)
    # ---------------------------
    def generate_excel_bytes():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_display.to_excel(writer, sheet_name='Cue表', index=False, startrow=7)
            wb = writer.book
            ws = writer.sheets['Cue表']
            
            fmt_title = wb.add_format({'bold': True, 'font_size': 14, 'align': 'left'})
            fmt_label = wb.add_format({'bold': True, 'align': 'right'})
            fmt_text = wb.add_format({'align': 'left'})
            fmt_header = wb.add_format({'bold': True, 'fg_color': '#4F81BD', 'font_color': 'white', 'border': 1})
            
            ws.write('A1', 'Media Schedule', fmt_title)
            ws.write('A3', '客戶名稱：', fmt_label)
            ws.write('B3', client_name, fmt_text)
            ws.write('A4', 'Product：', fmt_label)
            ws.write('B4', str_product, fmt_text)
            ws.write('A5', 'Period：', fmt_label)
            ws.write('B5', str_period, fmt_text)
            ws.write('A6', 'Medium：', fmt_label)
            ws.write('B6', str_medium, fmt_text)
            
            for c, val in enumerate(df_display.columns.values):
                ws.write(7, c, val, fmt_header)
            
            ws.set_column(0, 0, 18)
            ws.set_column(1, 6, 11)
            ws.set_column(7, len(df_display.columns)-1, 6)
        output.seek(0)
        return output

    # ---------------------------
    # 下載區塊
    # ---------------------------
    st.markdown("### 📥 匯出資料")
    col_dl, col_ragic = st.columns([1, 1])
    
    with col_dl:
        excel_file = generate_excel_bytes()
        st.download_button(
            label="下載 Excel 報表",
            data=excel_file,
            file_name=f"Schedule_{client_name}_{start_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.caption("手機請用瀏覽器開啟以確保下載成功")

    # ---------------------------
    # Ragic 上傳區塊
    # ---------------------------
    with col_ragic:
        with st.popover("☁️ 上傳至 Ragic"):
            st.markdown("#### 系統連線設定")
            
            # 使用預設值填入
            ragic_url = st.text_input("API URL", value=DEFAULT_RAGIC_URL)
            ragic_key = st.text_input("API Key", value=DEFAULT_API_KEY, type="password")
            
            if st.button("確認上傳", type="primary", use_container_width=True):
                if not ragic_url or not ragic_key:
                    st.error("請輸入 API URL 與 Key")
                else:
                    payload = {
                        RAGIC_CONFIG["client_name"]: client_name,
                        RAGIC_CONFIG["start_date"]: str(start_date),
                        RAGIC_CONFIG["end_date"]: str(end_date),
                        RAGIC_CONFIG["region"]: region,
                        RAGIC_CONFIG["total_budget"]: total_budget,
                    }
                    
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
                    
                    subtable_key = f"_subtable_{SUBTABLE_ROOT_ID}"
                    payload[subtable_key] = subtable_data
                    
                    st.info("步驟 1/2: 建立報價單資料...")
                    try:
                        api_url = ragic_url
                        if "?api" not in api_url:
                            api_url += "?api=true" if "?" not in api_url else "&api=true"
                            
                        resp = requests.post(
                            api_url, 
                            json=payload, 
                            headers={"Authorization": "Basic " + ragic_key}
                        )
                        
                        if resp.status_code == 200:
                            res_json = resp.json()
                            if res_json.get("status") == "SUCCESS":
                                ragic_id = res_json.get('ragicId')
                                st.success(f"✅ 資料建立成功！(ID: {ragic_id})")
                                
                                st.info("步驟 2/2: 上傳 Excel 附件...")
                                upload_file_bytes = generate_excel_bytes()
                                upload_filename = f"CueSheet_{client_name}.xlsx"
                                
                                base_url = ragic_url.split('?')[0]
                                if base_url.endswith('/'): base_url = base_url[:-1]
                                upload_url = f"{base_url}/{ragic_id}?api=true"
                                
                                files = {
                                    str(RAGIC_CONFIG["file_upload"]): (upload_filename, upload_file_bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                                }
                                
                                resp_upload = requests.post(
                                    upload_url,
                                    files=files,
                                    headers={"Authorization": "Basic " + ragic_key}
                                )
                                
                                if resp_upload.status_code == 200:
                                    st.success("🎉 附件上傳成功！全數完成。")
                                else:
                                    st.warning(f"附件上傳失敗: {resp_upload.text}")
                                
                            else:
                                st.error(f"上傳失敗: {res_json.get('msg')}")
                        else:
                            st.error(f"連線錯誤: {resp.status_code}")
                            st.write(resp.text)
                    except Exception as e:
                        st.error(f"發生錯誤: {e}")
