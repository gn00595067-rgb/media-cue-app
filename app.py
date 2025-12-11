import streamlit as st
import pandas as pd
import math
import io
import xlsxwriter
from datetime import timedelta, datetime

# ==========================================
# 1. 基礎資料與設定 (Configuration)
# ==========================================

# 區域與店數對照 (Program 欄位內容)
STORE_COUNTS = {
    # 廣播
    "全省": "4,437店", # 假設全省總數
    "北區": "北北基 1,649店",
    "桃竹苗": "桃竹苗 779店",
    "中區": "中彰投 839店",
    "雲嘉南": "雲嘉南 499店",
    "高屏": "高高屏 490店",
    "東區": "宜花東 181店",
    # 新鮮視 (前面加前綴以區分)
    "新鮮視_全省": "3,124面",
    "新鮮視_北區": "北北基 1,127面",
    "新鮮視_桃竹苗": "桃竹苗 616面",
    "新鮮視_中區": "中彰投 528面",
    "新鮮視_雲嘉南": "雲嘉南 365面",
    "新鮮視_高屏": "高高屏 405面",
    "新鮮視_東區": "宜花東 83面",
}

# 區域排序
REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]

# --- 報價資料庫 [List_Price(定價), Net_Price(實收價)] ---
# 為了讓計算符合邏輯，這裡使用 "Net_Price" 作為計算檔次的基準
# "List_Price" 用於顯示在 Rate 欄位 (通常 Rate = List / 720)
PRICING_DB = {
    "全家廣播": {
        "Std_Spots": 480, # 月標準檔次
        "全省": [400000, 320000], 
        "北區": [250000, 200000], "桃竹苗": [150000, 120000],
        "中區": [150000, 120000], "雲嘉南": [100000, 80000], 
        "高屏": [100000, 80000], "東區": [62500, 50000]
    },
    "新鮮視": {
        "Std_Spots": 504,
        "全省": [150000, 120000], 
        "北區": [150000, 120000], "桃竹苗": [120000, 96000],
        "中區": [90000, 72000], "雲嘉南": [75000, 60000], 
        "高屏": [75000, 60000], "東區": [45000, 36000]
    },
    "家樂福": {
        # 家樂福比較特殊，這裡設定兩條線的基準價
        # 假設預算分配給家樂福後，會自動拆成 量販 與 超市
        # 這裡的 Net 是指單檔成本估算
        "量販_全省": {"List": 310000, "Net_Unit": 595}, # 範例推算: 250000/420 approx
        "超市_全省": {"List": 100000, "Net_Unit": 111}  # 範例推算: 80000/720 approx
    }
}

# 秒數折扣 (影響價格)
DISCOUNT_TABLE = {5: 0.5, 10: 0.6, 15: 0.7, 20: 0.8, 25: 0.9, 30: 1.0, 35: 1.15, 40: 1.3, 45: 1.5, 60: 2.0}

def get_discount(seconds):
    if seconds in DISCOUNT_TABLE: return DISCOUNT_TABLE[seconds]
    for s in sorted(DISCOUNT_TABLE.keys()):
        if s >= seconds: return DISCOUNT_TABLE[s]
    return 1.0

def calculate_schedule(total_spots, days):
    """分配檔次：平均、偶數優先、前多後少"""
    if days == 0: return []
    schedule = [0] * days
    remaining = total_spots
    
    # 基礎平均
    base = remaining // days
    for i in range(days): schedule[i] = base
    remaining -= (base * days)
    
    # 餘數分配 (由前向後)
    idx = 0
    while remaining > 0:
        schedule[idx] += 1
        remaining -= 1
        idx = (idx + 1) % days
        
    # 偶數優化 (嘗試讓奇數變偶數)
    for i in range(days - 1):
        if schedule[i] % 2 != 0:
            if schedule[i+1] > 0:
                schedule[i] += 1; schedule[i+1] -= 1
            elif schedule[i] > 0:
                schedule[i] -= 1; schedule[i+1] += 1
    return schedule

# ==========================================
# 2. UI 介面 (Streamlit)
# ==========================================

st.set_page_config(layout="wide", page_title="Cue Sheet Generator Final")
st.markdown("""<style>.reportview-container { margin-top: -2em; } #MainMenu {visibility: hidden;} footer {visibility: hidden;} .stProgress > div > div > div > div { background-color: #ff4b4b; }</style>""", unsafe_allow_html=True)

st.title("媒體 Cue 表生成器")

# --- Sidebar ---
with st.sidebar:
    st.header("1. 基本資料")
    client_name = st.text_input("客戶名稱", "萬國通路")
    c1, c2 = st.columns(2)
    start_date = c1.date_input("開始日", datetime(2025, 1, 1))
    end_date = c2.date_input("結束日", datetime(2025, 1, 31))
    days_count = (end_date - start_date).days + 1
    total_budget_input = st.number_input("總預算 (未稅)", value=1140000, step=10000)

# --- Main Configuration (Waterfall Logic) ---
config_media = {}
st.subheader("2. 媒體投放設定 (連動總和 100%)")

col_m1, col_m2, col_m3 = st.columns(3)
remaining_global_share = 100 

# 1. 全家廣播 (Priority 1)
with col_m1:
    fm_act = st.checkbox("開啟全家廣播", value=True, key="fm_act")
    fm_data = None
    if fm_act:
        st.markdown("---")
        is_nat = st.checkbox("全省聯播", value=True, key="fm_nat")
        regs = ["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, key="fm_reg")
        
        # 秒數排序
        _secs_input = st.multiselect("秒數", DURATIONS, default=[20], key="fm_sec")
        secs = sorted(_secs_input)
        
        # 預算滑桿
        share = st.slider("廣播-預算佔比%", 0, remaining_global_share, min(70, remaining_global_share), key="fm_share")
        remaining_global_share -= share
        
        # 秒數佔比連動
        sec_shares = {}
        if len(secs) > 1:
            st.caption("各秒數佔比")
            ls = 100
            for i, s in enumerate(secs[:-1]):
                v = st.slider(f"{s}秒佔比", 0, ls, int(ls/2), key=f"fm_s_{s}")
                sec_shares[s] = v; ls -= v
            sec_shares[secs[-1]] = ls
            st.info(f"🔹 {secs[-1]}秒: {ls}% (餘額)")
        elif secs: sec_shares[secs[0]] = 100
            
        fm_data = {"is_national": is_nat, "regions": regs, "seconds": secs, "share": share, "sec_shares": sec_shares}

# 2. 新鮮視 (Priority 2)
with col_m2:
    fv_act = st.checkbox("開啟新鮮視", value=True, key="fv_act")
    fv_data = None
    if fv_act:
        st.markdown("---")
        is_nat = st.checkbox("全省聯播", value=False, key="fv_nat") # 預設非全省以符合範例
        regs = ["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, default=["北區", "桃竹苗"], key="fv_reg")
        
        _secs_input = st.multiselect("秒數", DURATIONS, default=[5], key="fv_sec")
        secs = sorted(_secs_input)
        
        # 預算滑桿 (上限為剩餘)
        limit = remaining_global_share
        default_val = min(20, limit)
        share = st.slider("新鮮視-預算佔比%", 0, limit, default_val, key="fv_share")
        remaining_global_share -= share
        
        sec_shares = {}
        if len(secs) > 1:
            st.caption("各秒數佔比")
            ls = 100
            for i, s in enumerate(secs[:-1]):
                v = st.slider(f"{s}秒佔比", 0, ls, int(ls/2), key=f"fv_s_{s}")
                sec_shares[s] = v; ls -= v
            sec_shares[secs[-1]] = ls
            st.info(f"🔹 {secs[-1]}秒: {ls}% (餘額)")
        elif secs: sec_shares[secs[0]] = 100
        
        fv_data = {"is_national": is_nat, "regions": regs, "seconds": secs, "share": share, "sec_shares": sec_shares}

# 3. 家樂福 (Priority 3 - Auto Fill)
with col_m3:
    cf_act = st.checkbox("開啟家樂福", value=True, key="cf_act")
    cf_data = None
    if cf_act:
        st.markdown("---")
        st.write("區域：全省")
        _secs_input = st.multiselect("秒數", DURATIONS, default=[20], key="cf_sec")
        secs = sorted(_secs_input)
        
        # 自動填滿
        share = remaining_global_share
        st.caption(f"家樂福-預算佔比: {share}% (自動填滿)")
        st.progress(share / 100.0 if share <= 100 else 1.0)
        
        sec_shares = {}
        if len(secs) > 1:
            st.caption("各秒數佔比")
            ls = 100
            for i, s in enumerate(secs[:-1]):
                v = st.slider(f"{s}秒佔比", 0, ls, int(ls/2), key=f"cf_s_{s}")
                sec_shares[s] = v; ls -= v
            sec_shares[secs[-1]] = ls
            st.info(f"🔹 {secs[-1]}秒: {ls}% (餘額)")
        elif secs: sec_shares[secs[0]] = 100
            
        cf_data = {"regions": ["全省"], "seconds": secs, "share": share, "sec_shares": sec_shares}

if fm_data: config_media["全家廣播"] = fm_data
if fv_data: config_media["新鮮視"] = fv_data
if cf_data: config_media["家樂福"] = cf_data

# ==========================================
# 3. 計算邏輯 (Calculator)
# ==========================================

final_rows = []
all_secs = set()
all_media = set()

# 總佔比檢查
total_share_sum = sum(m["share"] for m in config_media.values())

if total_share_sum > 0:
    for m_type, cfg in config_media.items():
        # 分配給該媒體的總預算
        media_budget = total_budget_input * (cfg["share"] / 100.0)
        all_media.add(m_type)
        
        # 針對該媒體下的每個秒數
        for sec, sec_share in cfg["sec_shares"].items():
            all_secs.add(f"{sec}秒")
            # 分配給該秒數的預算
            sec_budget = media_budget * (sec_share / 100.0)
            if sec_budget <= 0: continue
            
            discount = get_discount(sec)
            
            # --- 全家廣播 / 新鮮視 ---
            if m_type in ["全家廣播", "新鮮視"]:
                db = PRICING_DB[m_type]
                # 決定要計算的區域 (如果是全省，計算邏輯雖是一包，但顯示要展開)
                calc_regions = ["全省"] if cfg["is_national"] else cfg["regions"]
                display_regions = REGIONS_ORDER if cfg["is_national"] else cfg["regions"]
                
                # 計算組合單價 (Net)
                combined_unit_net = 0
                for reg in calc_regions:
                    # 實收價基準 (用 Std_Spots 換算單檔單價)
                    net_price_total = db[reg][1]
                    unit_net = (net_price_total / db["Std_Spots"]) * discount
                    combined_unit_net += unit_net
                
                if combined_unit_net == 0: continue
                
                # 逆推檔次 (Ceil: 確保金額 > 預算)
                target_spots = math.ceil(sec_budget / combined_unit_net)
                if target_spots == 0: target_spots = 1
                
                # 產生排程
                daily_sch = calculate_schedule(target_spots, days_count)
                
                # 計算 Package Cost (僅全省時)
                pkg_cost_total = 0
                if cfg["is_national"]:
                    nat_list = db["全省"][0]
                    # 公式: 定價/720 * 檔次 * 折扣 * 1.1 (if <720)
                    mult = 1.1 if target_spots < 720 else 1.0
                    pkg_cost_total = (nat_list / 720.0) * target_spots * discount * mult

                # 生成資料列
                for reg in display_regions:
                    # 定價 (Rate Net 的分子)
                    list_price = db.get(reg, [0,0])[0] if cfg["is_national"] else db[reg][0]
                    
                    # Rate (Net) 欄位顯示值 (List / 720 * Spots * Disc)
                    rate_val = (list_price / 720.0) * target_spots * discount
                    
                    # 真實成本 (用於 Grand Total)
                    # 全省: 成本算在第一筆(北區)以免重複加總; 區域: 各自算
                    real_c = (combined_unit_net * target_spots) if (not cfg["is_national"] or reg == "北區") else 0
                    
                    # Package Cost 顯示 (只在第一筆標記)
                    pkg_val = pkg_cost_total if (cfg["is_national"] and reg == "北區") else 0
                    
                    # 店數顯示
                    prog_name = STORE_COUNTS.get(reg, reg)
                    if m_type == "新鮮視":
                        prog_name = STORE_COUNTS.get(f"新鮮視_{reg}", reg)
                    
                    final_rows.append({
                        "media": m_type, 
                        "region": reg, 
                        "location": f"{reg.replace('區', '')}區-{reg}" if m_type=="全家廣播" else f"{reg.replace('區', '')}區-{reg}", # 模擬excel location格式
                        "program": prog_name,
                        "daypart": "00:00-24:00" if m_type=="全家廣播" else "07:00-22:00",
                        "seconds": sec, 
                        "schedule": daily_sch, 
                        "spots": target_spots,
                        "rate_net": rate_val, 
                        "pkg_cost": pkg_val, 
                        "is_pkg_start": (cfg["is_national"] and reg == "北區"),
                        "is_pkg_member": cfg["is_national"], 
                        "real_cost": real_c
                    })

            # --- 家樂福 ---
            elif m_type == "家樂福":
                db = PRICING_DB["家樂福"]
                # 家樂福固定產生 量販 + 超市
                unit_hyp = db["量販_全省"]["Net_Unit"] * discount
                unit_sup = db["超市_全省"]["Net_Unit"] * discount
                combined = unit_hyp + unit_sup
                
                target_spots = math.ceil(sec_budget / combined)
                if target_spots == 0: target_spots = 1
                
                sch = calculate_schedule(target_spots, days_count)
                
                # 量販 Row
                final_rows.append({
                    "media": "家樂福", "region": "全省量販", "location": "全省量販", "program": "67店",
                    "daypart": "09:00-23:00", "seconds": sec, "schedule": sch, "spots": target_spots,
                    "rate_net": (db["量販_全省"]["List"]/720.0)*target_spots*discount,
                    "pkg_cost": 0, "is_pkg_start": False, "is_pkg_member": False, 
                    "real_cost": unit_hyp * target_spots
                })
                # 超市 Row
                final_rows.append({
                    "media": "家樂福", "region": "全省超市", "location": "全省超市", "program": "250店",
                    "daypart": "00:00-24:00", "seconds": sec, "schedule": sch, "spots": target_spots,
                    "rate_net": (db["超市_全省"]["List"]/720.0)*target_spots*discount,
                    "pkg_cost": 0, "is_pkg_start": False, "is_pkg_member": False, 
                    "real_cost": unit_sup * target_spots
                })

# 計算總金額
media_total = sum(r["real_cost"] for r in final_rows)
prod_cost = 10000
vat = (media_total + prod_cost) * 0.05
grand_total = media_total + prod_cost + vat

# 折扣率
discount_ratio_str = "N/A"
if grand_total > 0:
    ratio = (total_budget_input / grand_total) * 100
    discount_ratio_str = f"{ratio:.1f}%"

# ==========================================
# 4. Excel 生成 (XlsxWriter)
# ==========================================

def generate_excel(rows, days_cnt, start_dt, c_name, products, mediums, totals_data):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Media Schedule")

    # 樣式
    fmt_title = workbook.add_format({'font_size': 18, 'bold': True, 'align': 'center'})
    fmt_header_left = workbook.add_format({'align': 'left', 'valign': 'top', 'bold': True, 'border': 0})
    # 藍色表頭
    fmt_col_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#DDEBF7', 'text_wrap': True, 'font_size': 10})
    # 日期 (週末黃底)
    fmt_date_wk = workbook.add_format({'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#DDEBF7'})
    fmt_date_we = workbook.add_format({'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFF2CC'}) # 黃底
    
    fmt_cell = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 10})
    fmt_cell_left = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'font_size': 10, 'text_wrap': True})
    fmt_num = workbook.add_format({'align': 'right', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0', 'font_size': 10})
    # 檔次黃底
    fmt_spots = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bold': True, 'bg_color': '#FFF2CC', 'font_size': 10})
    
    # 寫入標題與 Info
    worksheet.merge_range('A1:AJ1', "Media Schedule", fmt_title)
    
    info = [
        ("客戶名稱：", c_name),
        ("Product：", products),
        ("Period :", f"{start_dt.strftime('%Y. %m. %d')} - {end_date.strftime('%Y. %m. %d')}"),
        ("Medium :", mediums)
    ]
    for i, (label, val) in enumerate(info):
        worksheet.write(2+i, 0, label, fmt_header_left)
        worksheet.write(2+i, 1, val, fmt_header_left)

    # 寫入月曆 Header
    # Row 6: 月份
    worksheet.write(6, 6, f"{start_dt.month}月", fmt_cell)
    # Row 7: 日期 (1, 2...)
    # Row 8: 星期 (三, 四...)
    weekdays = ["一", "二", "三", "四", "五", "六", "日"]
    curr = start_dt
    for i in range(days_cnt):
        col_idx = 7 + i
        wd = curr.weekday()
        # 樣式：週末黃底
        fmt = fmt_date_we if wd >= 5 else fmt_date_wk
        worksheet.write(7, col_idx, curr.day, fmt)
        worksheet.write(8, col_idx, weekdays[wd], fmt)
        curr += timedelta(days=1)

    # 主表頭 (Row 8)
    headers = ["Station", "Location", "Program", "Day-part", "Size", "rate (Net)", "Package-cost\n(Net)"]
    for i, h in enumerate(headers):
        worksheet.write(8, i, h, fmt_col_header)
    
    # "檔次" 在最後
    last_col = 7 + days_cnt
    worksheet.write(8, last_col, "檔次", fmt_col_header)

    # 寫入資料
    current_row = 9
    i = 0
    while i < len(rows):
        row = rows[i]
        # 尋找 Group (同媒體同秒數)
        j = i + 1
        while j < len(rows) and rows[j]['media'] == row['media'] and rows[j]['seconds'] == row['seconds']:
            j += 1
        group_size = j - i
        
        # 寫入 Station (合併)
        m_name = row['media']
        if "全家廣播" in m_name: m_name = "全家便利商店\n通路廣播廣告"
        if "新鮮視" in m_name: m_name = "全家便利商店\n新鮮視廣告"
        
        if group_size > 1:
            worksheet.merge_range(current_row, 0, current_row + group_size - 1, 0, m_name, fmt_cell_left)
        else:
            worksheet.write(current_row, 0, m_name, fmt_cell_left)
            
        # 寫入各列資料
        for k in range(group_size):
            r_data = rows[i + k]
            r_idx = current_row + k
            
            # Location 對應 Excel 範例
            loc_txt = r_data['location']
            if "北北基" in loc_txt and "廣播" in r_data['media']: loc_txt = "北區-北北基+東" # 特例處理
            
            worksheet.write(r_idx, 1, loc_txt, fmt_cell)
            worksheet.write(r_idx, 2, r_data['program'], fmt_cell)
            worksheet.write(r_idx, 3, r_data['daypart'], fmt_cell)
            worksheet.write(r_idx, 4, f"{r_data['seconds']}秒", fmt_cell)
            worksheet.write(r_idx, 5, r_data['rate_net'], fmt_num)
            
            # Schedule
            for d_idx, s_val in enumerate(r_data['schedule']):
                worksheet.write(r_idx, 7 + d_idx, s_val, fmt_cell)
                
            # Total Spots
            worksheet.write(r_idx, last_col, r_data['spots'], fmt_spots)

        # Package Cost 合併
        if row['is_pkg_start']:
            if group_size > 1:
                worksheet.merge_range(current_row, 6, current_row + group_size - 1, 6, row['pkg_cost'], fmt_num)
            else:
                worksheet.write(current_row, 6, row['pkg_cost'], fmt_num)
        elif not row['is_pkg_member']:
            # 非 Package，每格填空或個別值(家樂福)
             for k in range(group_size):
                 val = rows[i+k]['rate_net'] if "家樂福" in rows[i+k]['media'] else ""
                 # 家樂福範例中 Package-cost 欄位是空的或填特定值? 範例圖中量販有值
                 if "家樂福" in rows[i+k]['media']:
                     # 簡單邏輯: 家樂福實收價填在 Package Cost 欄位? 
                     # 依照範例圖: Rate(Net)=310,000, Package-cost=258,333
                     # 我們這裡已經算出 Rate(Net), Package Cost 欄位若無全省包則留白
                     # 依照截圖，家樂福的實收顯示在 Package-cost 欄位
                     worksheet.write(r_idx, 6, rows[i+k]['real_cost'], fmt_num) 
                 else:
                     worksheet.write(current_row + k, 6, "", fmt_num)

        current_row += group_size
        i = j

    # Footer Totals
    worksheet.write(current_row, 2, "Total", fmt_cell)
    worksheet.write(current_row, 5, sum(r['rate_net'] for r in rows), fmt_num)
    worksheet.write(current_row, 6, totals_data['media_total'], fmt_num) # 實收總計
    
    # 檔次總計
    total_spots_daily = [0] * days_cnt
    for r in rows:
        for idx, val in enumerate(r['schedule']):
            total_spots_daily[idx] += val
    for idx, val in enumerate(total_spots_daily):
        worksheet.write(current_row, 7+idx, val, fmt_cell)
    worksheet.write(current_row, last_col, sum(r['spots'] for r in rows), fmt_cell)
    
    current_row += 1
    worksheet.write(current_row, 6, "製作", fmt_cell)
    worksheet.write(current_row, 7, totals_data['prod_cost'], fmt_num)
    current_row += 1
    worksheet.write(current_row, 6, "5% VAT", fmt_cell)
    worksheet.write(current_row, 7, totals_data['vat'], fmt_num)
    current_row += 1
    worksheet.write(current_row, 6, "Grand Total", fmt_cell)
    worksheet.write(current_row, 7, totals_data['grand_total'], fmt_num)

    # 調整欄寬
    worksheet.set_column('A:A', 20)
    worksheet.set_column('B:B', 15)
    worksheet.set_column('C:E', 12)
    worksheet.set_column('F:G', 12)
    worksheet.set_column(7, last_col, 4)

    workbook.close()
    return output

# ==========================================
# 5. Output Display
# ==========================================

st.markdown("### 計算結果摘要")
m1, m2, m3 = st.columns(3)
m1.metric("客戶預算", f"{total_budget_input:,}")
m2.metric("Cue表總金額 (含稅)", f"{int(grand_total):,}", delta=f"差異 +{int(grand_total - total_budget_input):,}")
m3.metric("預算/表價比 (折扣率)", discount_ratio_str)

# HTML Preview (Simplified)
if final_rows:
    df_preview = pd.DataFrame(final_rows)
    # 簡單呈現關鍵欄位
    st.dataframe(df_preview[['media', 'location', 'seconds', 'spots', 'rate_net', 'real_cost']])

    # Download Button
    product_str = "、".join(sorted(list(all_secs)))
    medium_str = "、".join(list(all_media))
    
    xlsx_data = generate_excel(
        final_rows, days_count, start_date, client_name, 
        product_str, medium_str,
        {"media_total": media_total, "prod_cost": prod_cost, "vat": vat, "grand_total": grand_total}
    )
    
    st.download_button(
        label="📥 下載 Excel Cue表 (.xlsx)",
        data=xlsx_data.getvalue(),
        file_name=f"CueSheet_{client_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
