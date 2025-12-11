import streamlit as st
import pandas as pd
import math
import io
import xlsxwriter
from datetime import timedelta, datetime

# ==========================================
# 1. 基礎資料與設定
# ==========================================

# 區域與店數對照
STORE_COUNTS = {
    "北區": "北北基 1649店",
    "桃竹苗": "桃竹苗 779店",
    "中區": "中彰投 839店",
    "雲嘉南": "雲嘉南 499店",
    "高屏": "高高屏 490店",
    "東區": "宜花東 181店",
    "新鮮視_北區": "北北基 1127店",
    "新鮮視_桃竹苗": "桃竹苗 616店",
    "新鮮視_中區": "中彰投 528店",
    "新鮮視_雲嘉南": "雲嘉南 365店",
    "新鮮視_高屏": "高高屏 405店",
    "新鮮視_東區": "宜花東 83店",
}

REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]

# 報價資料庫 [List_Price(定價), Net_Price(實收價)]
PRICING_DB = {
    "全家廣播": {
        "Base_Sec": 30, "Std_Spots": 480,
        "全省": [400000, 320000], "北區": [250000, 200000], "桃竹苗": [150000, 120000],
        "中區": [150000, 120000], "雲嘉南": [100000, 80000], "高屏": [100000, 80000],
        "東區": [62500, 50000]
    },
    "新鮮視": {
        "Base_Sec": 10, "Std_Spots": 504,
        "全省": [150000, 120000], "北區": [150000, 120000], "桃竹苗": [120000, 96000],
        "中區": [90000, 72000], "雲嘉南": [75000, 60000], "高屏": [75000, 60000],
        "東區": [45000, 36000]
    },
    "家樂福": {
        "量販_全省": {"List": 300000, "Net": 250000, "Std_Spots": 420},
        "超市_全省": {"List": 100000, "Net": 80000, "Std_Spots": 720} 
    }
}

DISCOUNT_TABLE = {5: 0.5, 10: 0.6, 15: 0.7, 20: 0.8, 25: 0.9, 30: 1.0, 35: 1.15, 40: 1.3, 45: 1.5, 60: 2.0}

def get_discount(seconds):
    if seconds in DISCOUNT_TABLE: return DISCOUNT_TABLE[seconds]
    for s in sorted(DISCOUNT_TABLE.keys()):
        if s >= seconds: return DISCOUNT_TABLE[s]
    return 1.0

def calculate_schedule(total_spots, days):
    if days == 0: return []
    schedule = [0] * days
    remaining = total_spots
    base = remaining // days
    for i in range(days): schedule[i] = base
    remaining -= (base * days)
    idx = 0
    while remaining > 0:
        schedule[idx] += 1
        remaining -= 1
        idx = (idx + 1) % days
    for i in range(days - 1):
        if schedule[i] % 2 != 0:
            if schedule[i+1] > 0:
                schedule[i] += 1; schedule[i+1] -= 1
            elif schedule[i] > 0:
                schedule[i] -= 1; schedule[i+1] += 1
    return schedule

# ==========================================
# 2. UI 設定
# ==========================================

st.set_page_config(layout="wide", page_title="Cue Sheet Generator v4")
st.markdown("""<style>.reportview-container { margin-top: -2em; } #MainMenu {visibility: hidden;} footer {visibility: hidden;}</style>""", unsafe_allow_html=True)

st.title("媒體 Cue 表生成器 (瀑布式預算連動版)")

with st.sidebar:
    st.header("1. 基本資料")
    client_name = st.text_input("客戶名稱", "範例客戶")
    c1, c2 = st.columns(2)
    start_date = c1.date_input("開始日", datetime.today())
    end_date = c2.date_input("結束日", datetime.today() + timedelta(days=13))
    days_count = (end_date - start_date).days + 1
    total_budget_input = st.number_input("總預算 (未稅)", value=500000, step=10000)

config_media = {}
st.subheader("2. 媒體投放設定 (自動連動總和 100%)")

# --------------------------------------------------------
# Step 1: 先定義使用者想開啟哪些媒體
# (為了讓版面整齊，我們把「是否開啟」的開關先放在上面，或直接在 Column 裡處理)
# 但為了做「連動計算」，我們必須在 render slider 之前知道有哪些媒體是 active 的
# --------------------------------------------------------

# 這裡使用 Session State 或簡單的變數流
# 為了 UX，我們還是保持三欄佈局，但邏輯上我們依序扣除剩餘預算

col_m1, col_m2, col_m3 = st.columns(3)

# 變數：剩餘可分配的媒體總預算 (Global Share)
remaining_global_share = 100 

# --- 全家廣播 (優先權 1) ---
with col_m1:
    fm_act = st.checkbox("開啟全家廣播", value=True, key="fm_act")
    fm_data = None
    
    if fm_act:
        st.markdown("---")
        is_nat = st.checkbox("全省聯播", value=True, key="fm_nat")
        regs = ["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, key="fm_reg")
        
        # [修改] 秒數由小排到大
        _secs_input = st.multiselect("秒數", DURATIONS, default=[20], key="fm_sec")
        secs = sorted(_secs_input)
        
        # [修改] 媒體預算佔比 - 瀑布式邏輯
        # 廣播是第一個，所以它可以選 0 到 100
        share = st.slider("廣播-總預算佔比%", 0, remaining_global_share, min(40, remaining_global_share), key="fm_share")
        remaining_global_share -= share # 扣除廣播用掉的
        
        # [修改] 秒數預算佔比
        sec_shares = {}
        if len(secs) > 1:
            st.caption("各秒數佔比 (自動連動)")
            ls = 100
            for i, s in enumerate(secs[:-1]):
                v = st.slider(f"{s}秒佔比", 0, ls, int(ls/2), key=f"fm_s_{s}")
                sec_shares[s] = v; ls -= v
            # 最後一個
            sec_shares[secs[-1]] = ls
            st.info(f"🔹 {secs[-1]}秒: {ls}% (餘額)")
        elif secs: 
            sec_shares[secs[0]] = 100
            
        fm_data = {"is_national": is_nat, "regions": regs, "seconds": secs, "share": share, "sec_shares": sec_shares}

# --- 新鮮視 (優先權 2) ---
with col_m2:
    fv_act = st.checkbox("開啟新鮮視", value=True, key="fv_act")
    fv_data = None
    
    if fv_act:
        st.markdown("---")
        is_nat = st.checkbox("全省聯播", value=True, key="fv_nat")
        regs = ["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, key="fv_reg")
        
        _secs_input = st.multiselect("秒數", DURATIONS, default=[10], key="fv_sec")
        secs = sorted(_secs_input)
        
        # [修改] 媒體預算佔比 - 瀑布式邏輯
        # 新鮮視只能選「剩下」的
        # 如果剩下 0，就強制 0
        if remaining_global_share > 0:
            share = st.slider("新鮮視-總預算佔比%", 0, remaining_global_share, min(30, remaining_global_share), key="fv_share")
        else:
            share = 0
            st.warning("預算已在廣播分配完畢 (0%)")
            
        remaining_global_share -= share # 扣除新鮮視用掉的
        
        sec_shares = {}
        if len(secs) > 1:
            st.caption("各秒數佔比 (自動連動)")
            ls = 100
            for i, s in enumerate(secs[:-1]):
                v = st.slider(f"{s}秒佔比", 0, ls, int(ls/2), key=f"fv_s_{s}")
                sec_shares[s] = v; ls -= v
            sec_shares[secs[-1]] = ls
            st.info(f"🔹 {secs[-1]}秒: {ls}% (餘額)")
        elif secs: 
            sec_shares[secs[0]] = 100
        
        fv_data = {"is_national": is_nat, "regions": regs, "seconds": secs, "share": share, "sec_shares": sec_shares}

# --- 家樂福 (優先權 3 - 撿剩的) ---
with col_m3:
    cf_act = st.checkbox("開啟家樂福", key="cf_act")
    cf_data = None
    
    if cf_act:
        st.markdown("---")
        st.write("區域：全省")
        
        _secs_input = st.multiselect("秒數", DURATIONS, default=[10], key="cf_sec")
        secs = sorted(_secs_input)
        
        # [修改] 媒體預算佔比 - 瀑布式邏輯
        # 家樂福自動接收所有剩下的
        share = remaining_global_share
        st.info(f"家樂福-總預算佔比: {share}% (自動填滿)")
        # 無需 Slider，直接顯示
        st.progress(share / 100.0 if share <= 100 else 1.0)
        
        sec_shares = {}
        if len(secs) > 1:
            st.caption("各秒數佔比 (自動連動)")
            ls = 100
            for i, s in enumerate(secs[:-1]):
                v = st.slider(f"{s}秒佔比", 0, ls, int(ls/2), key=f"cf_s_{s}")
                sec_shares[s] = v; ls -= v
            sec_shares[secs[-1]] = ls
            st.info(f"🔹 {secs[-1]}秒: {ls}% (餘額)")
        elif secs: 
            sec_shares[secs[0]] = 100
            
        cf_data = {"regions": ["全省"], "seconds": secs, "share": share, "sec_shares": sec_shares}

# 寫入設定檔
if fm_data: config_media["全家廣播"] = fm_data
if fv_data: config_media["新鮮視"] = fv_data
if cf_data: config_media["家樂福"] = cf_data

# ==========================================
# 3. 計算邏輯 (保持不變)
# ==========================================

final_rows = []
all_secs = set()
all_media = set()
total_share_sum = sum(m["share"] for m in config_media.values())

# 注意：因為我們的 UI 邏輯強制總和為 100 (如果有開家樂福的話)，或者小於等於 100
# 所以這裡的 total_share_sum 應該就是 100 (除非只開了廣播選50%)
# 如果只開廣播選 50%，剩下的 50% 就不會被計算。這符合邏輯（部分預算不使用）。

if total_share_sum > 0:
    for m_type, cfg in config_media.items():
        # 分配金額
        if total_share_sum == 100:
            media_budget = total_budget_input * (cfg["share"] / 100.0)
        else:
            # 如果使用者只選了廣播 50% 且沒開其他，這裡我們假設他是想用 50% 的總預算
            # 或者是 正規化? 通常是直接乘比例
            media_budget = total_budget_input * (cfg["share"] / 100.0)
            
        all_media.add(m_type)
        
        for sec, sec_share in cfg["sec_shares"].items():
            all_secs.add(f"{sec}秒")
            sec_budget = media_budget * (sec_share / 100.0)
            if sec_budget <= 0: continue
            
            discount = get_discount(sec)
            
            if m_type in ["全家廣播", "新鮮視"]:
                db = PRICING_DB[m_type]
                calc_regions = ["全省"] if cfg["is_national"] else cfg["regions"]
                display_regions = REGIONS_ORDER if cfg["is_national"] else cfg["regions"]
                
                combined_unit_net = 0
                for reg in calc_regions:
                    net_price = db[reg][1]
                    unit_net = (net_price / db["Std_Spots"]) * discount
                    combined_unit_net += unit_net
                
                if combined_unit_net == 0: continue
                
                target_spots = math.ceil(sec_budget / combined_unit_net)
                if target_spots == 0: target_spots = 1
                
                daily_sch = calculate_schedule(target_spots, days_count)
                
                pkg_cost_total = 0
                if cfg["is_national"]:
                    nat_list = db["全省"][0]
                    mult = 1.1 if target_spots < 720 else 1.0
                    pkg_cost_total = (nat_list / 720.0) * target_spots * discount * mult

                for reg in display_regions:
                    list_price = db.get(reg, [0,0])[0] if cfg["is_national"] else db[reg][0]
                    rate_val = (list_price / 720.0) * target_spots * discount
                    real_c = (combined_unit_net * target_spots) if (not cfg["is_national"] or reg == "北區") else 0
                    pkg_val = pkg_cost_total if (cfg["is_national"] and reg == "北區") else 0
                    
                    final_rows.append({
                        "media": m_type, "region": reg, "program": STORE_COUNTS.get(reg if m_type=="全家廣播" else f"新鮮視_{reg}", reg),
                        "daypart": "07:00-23:00", "seconds": sec, "schedule": daily_sch, "spots": target_spots,
                        "rate_net": rate_val, "pkg_cost": pkg_val, "is_pkg_start": (cfg["is_national"] and reg == "北區"),
                        "is_pkg_member": cfg["is_national"], "real_cost": real_c
                    })

            elif m_type == "家樂福":
                db = PRICING_DB["家樂福"]
                unit_hyp = (db["量販_全省"]["Net"] / db["量販_全省"]["Std_Spots"]) * discount
                unit_sup = (db["超市_全省"]["Net"] / db["超市_全省"]["Std_Spots"]) * discount
                combined = unit_hyp + unit_sup
                
                target_spots = math.ceil(sec_budget / combined)
                if target_spots == 0: target_spots = 1
                
                sch = calculate_schedule(target_spots, days_count)
                
                final_rows.append({
                    "media": "家樂福", "region": "全省量販", "program": "全省", "daypart": "09:00-23:00",
                    "seconds": sec, "schedule": sch, "spots": target_spots,
                    "rate_net": (db["量販_全省"]["List"]/720.0)*target_spots*discount,
                    "pkg_cost": 0, "is_pkg_start": False, "is_pkg_member": False, "real_cost": unit_hyp * target_spots
                })
                final_rows.append({
                    "media": "家樂福", "region": "全省超市", "program": "全省", "daypart": "00:00-24:00",
                    "seconds": sec, "schedule": sch, "spots": target_spots,
                    "rate_net": (db["超市_全省"]["List"]/720.0)*target_spots*discount,
                    "pkg_cost": 0, "is_pkg_start": False, "is_pkg_member": False, "real_cost": unit_sup * target_spots
                })

# ==========================================
# 4. 生成 Excel & HTML (保持不變)
# ==========================================

media_total = sum(r["real_cost"] for r in final_rows)
prod_cost = 10000
vat = (media_total + prod_cost) * 0.05
grand_total = media_total + prod_cost + vat

discount_ratio_str = "N/A"
if grand_total > 0:
    ratio = (total_budget_input / grand_total) * 100
    discount_ratio_str = f"{ratio:.1f}%"

def generate_excel(rows, days_cnt, start_dt, c_name, products, mediums, totals_data):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Cue Sheet")
    fmt_header_info = workbook.add_format({'align': 'left', 'valign': 'top', 'bold': True, 'border': 0, 'bg_color': '#f2f2f2'})
    fmt_col_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#cfe2f3', 'text_wrap': True})
    fmt_date = workbook.add_format({'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#cfe2f3', 'rotation': 90})
    fmt_cell = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 11})
    fmt_cell_left = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'font_size': 11})
    fmt_num = workbook.add_format({'align': 'right', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0'})
    fmt_spots = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bold': True, 'bg_color': '#ffeb3b'})
    fmt_total_label = workbook.add_format({'align': 'right', 'bold': True, 'border': 1, 'bg_color': '#ffffcc'})
    fmt_grand = workbook.add_format({'align': 'right', 'bold': True, 'border': 2, 'bg_color': '#ffffcc', 'num_format': '"NT$ "#,##0'})

    info_text = f"Client: {c_name}\nProduct: {products}\nPeriod: {start_dt.strftime('%Y/%m/%d')} ~ {(start_dt+timedelta(days=days_cnt-1)).strftime('%Y/%m/%d')}\nMedium: {mediums}"
    worksheet.merge_range(0, 0, 3, 5, info_text, fmt_header_info)
    headers = ["Station", "Location", "Program", "Day-part", "Size"]
    for i, h in enumerate(headers): worksheet.write(4, i, h, fmt_col_header)
    curr = start_dt
    for i in range(days_cnt): worksheet.write(4, 5 + i, curr.strftime('%m/%d'), fmt_date); curr += timedelta(days=1)
    last_col = 5 + days_cnt
    worksheet.write(4, last_col, "Total\nSpots", fmt_col_header); worksheet.write(4, last_col + 1, "Rate\n(Net)", fmt_col_header); worksheet.write(4, last_col + 2, "Package\nCost", fmt_col_header)

    current_row = 5
    i = 0
    while i < len(rows):
        row = rows[i]
        j = i + 1
        while j < len(rows) and rows[j]['media'] == row['media'] and rows[j]['seconds'] == row['seconds']: j += 1
        group_size = j - i
        m_name = row['media'].replace("全家廣播", "全家便利商店\n通路廣播廣告").replace("新鮮視", "全家便利商店\n新鮮視")
        if group_size > 1: worksheet.merge_range(current_row, 0, current_row + group_size - 1, 0, m_name, fmt_cell_left)
        else: worksheet.write(current_row, 0, m_name, fmt_cell_left)
        for k in range(group_size):
            r_data = rows[i + k]; r_idx = current_row + k
            worksheet.write(r_idx, 1, r_data['region'], fmt_cell)
            worksheet.write(r_idx, 2, r_data['program'], fmt_cell)
            worksheet.write(r_idx, 3, r_data['daypart'], fmt_cell)
            worksheet.write(r_idx, 4, f"{r_data['seconds']}秒", fmt_cell)
            for d_idx, s_val in enumerate(r_data['schedule']): worksheet.write(r_idx, 5 + d_idx, s_val, fmt_cell)
            worksheet.write(r_idx, last_col, r_data['spots'], fmt_spots)
            worksheet.write(r_idx, last_col + 1, r_data['rate_net'], fmt_num)
        if row['is_pkg_start']:
            if group_size > 1: worksheet.merge_range(current_row, last_col + 2, current_row + group_size - 1, last_col + 2, row['pkg_cost'], fmt_num)
            else: worksheet.write(current_row, last_col + 2, row['pkg_cost'], fmt_num)
        elif not row['is_pkg_member']:
            for k in range(group_size): worksheet.write(current_row + k, last_col + 2, "", fmt_num)
        current_row += group_size; i = j

    worksheet.merge_range(current_row, 0, current_row, 4, "Media Total", fmt_total_label)
    worksheet.merge_range(current_row, 5, current_row, 5 + days_cnt - 1, "", fmt_total_label)
    worksheet.write(current_row, last_col, sum(r['spots'] for r in rows), fmt_total_label)
    worksheet.write(current_row, last_col + 1, totals_data['media_total'], fmt_num); worksheet.write(current_row, last_col + 2, "", fmt_total_label)
    current_row += 1
    worksheet.merge_range(current_row, 0, current_row, 4, "Production Cost", fmt_total_label)
    worksheet.merge_range(current_row, 5, current_row, last_col + 2, totals_data['prod_cost'], fmt_num)
    current_row += 1
    worksheet.merge_range(current_row, 0, current_row, 4, "5% VAT", fmt_total_label)
    worksheet.merge_range(current_row, 5, current_row, last_col + 2, totals_data['vat'], fmt_num)
    current_row += 1
    worksheet.merge_range(current_row, 0, current_row, 4, "Grand Total", fmt_total_label)
    worksheet.merge_range(current_row, 5, current_row, last_col + 2, totals_data['grand_total'], fmt_grand)
    worksheet.set_column(0, 0, 25); worksheet.set_column(1, 1, 15); worksheet.set_column(2, 2, 20); worksheet.set_column(3, 4, 10); worksheet.set_column(5, 5+days_cnt, 4); worksheet.set_column(last_col+1, last_col+2, 12)
    workbook.close()
    return output

st.markdown("### 計算結果摘要")
m1, m2, m3 = st.columns(3)
m1.metric("客戶預算", f"{total_budget_input:,}")
m2.metric("Cue表總金額 (含稅)", f"{int(grand_total):,}", delta=f"差異 +{int(grand_total - total_budget_input):,}")
m3.metric("預算/表價比 (折扣率)", discount_ratio_str, help="計算公式: 客戶預算 / Cue表總額")

date_headers = "".join([f"<th class='date-col'>{(start_date + timedelta(days=i)).strftime('%m/%d')}</th>" for i in range(days_count)])
rows_html = ""
prev_media_sec = None
for idx, row in enumerate(final_rows):
    sch_cells = "".join([f"<td class='sch'>{s}</td>" for s in row["schedule"]])
    pkg_str = f"{int(row['pkg_cost']):,}" if row["is_pkg_start"] else ""
    pkg_style = "border-bottom: none;" if row["is_pkg_start"] else ("border-top: none; border-bottom: none;" if row["is_pkg_member"] else "")
    if row["is_pkg_member"] and (idx == len(final_rows)-1 or not final_rows[idx+1]["is_pkg_member"]): pkg_style = "border-top: none;"
    
    media_display = row['media']
    if media_display == prev_media_sec: media_display = '<span style="color:#ddd">"</span>'
    else:
        prev_media_sec = media_display
        if "全家廣播" in media_display: media_display = "全家便利商店<br>通路廣播廣告"
        if "新鮮視" in media_display: media_display = "全家便利商店<br>新鮮視"

    rows_html += f"""<tr><td style="text-align:left; font-size:10px;">{media_display}</td><td>{row['region']}</td><td>{row['program']}</td><td>{row['daypart']}</td><td>{row['seconds']}</td>{sch_cells}<td style="font-weight:bold; background-color:#ffeb3b4d;">{row['spots']}</td><td style="text-align:right;">{int(row['rate_net']):,}</td><td style="text-align:right; {pkg_style}">{pkg_str}</td></tr>"""

html_template = f"""
<style>table {{ width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; font-size: 11px; }} th, td {{ border: 1px solid #999; padding: 3px; text-align: center; }} .head-info {{ background-color: #f2f2f2; text-align: left; padding: 8px; border:none; }} .col-head {{ background-color: #cfe2f3; font-weight: bold; }} .date-col {{ writing-mode: vertical-rl; transform: rotate(180deg); width: 20px; font-size: 10px; background-color: #cfe2f3; }} .sch {{ font-size: 11px; }} .total-row {{ background-color: #ffffcc; font-weight: bold; }}</style>
<div style="background:white; padding:10px; border: 1px solid #ccc; overflow-x: auto;"><table><tr><td colspan="5" class="head-info"><b>Client:</b> {client_name}<br><b>Product:</b> {", ".join(sorted(list(all_secs)))}<br><b>Period:</b> {start_date.strftime('%Y/%m/%d')} ~ {end_date.strftime('%Y/%m/%d')}<br><b>Medium:</b> {", ".join(list(all_media))}</td><td colspan="{days_count + 3}" style="border:none;"></td></tr><tr class="col-head"><th>Station</th><th>Location</th><th>Program</th><th>Day-part</th><th>Size</th>{date_headers}<th>Total Spots</th><th>Rate (Net)</th><th>Package Cost</th></tr>{rows_html}<tr class="total-row"><td colspan="5" style="text-align:right">Media Total</td><td colspan="{days_count}"></td><td>{sum(r['spots'] for r in final_rows)}</td><td style="text-align:right">{int(media_total):,}</td><td></td></tr><tr><td colspan="5" style="text-align:right">Production Cost</td><td colspan="{days_count + 2}" style="text-align:right">{prod_cost:,}</td><td></td></tr><tr><td colspan="5" style="text-align:right">5% VAT</td><td colspan="{days_count + 2}" style="text-align:right">{int(vat):,}</td><td></td></tr><tr class="total-row" style="border-top: 2px double black; font-size: 13px;"><td colspan="5" style="text-align:right">Grand Total</td><td colspan="{days_count + 2}" style="text-align:right">NT$ {int(grand_total):,}</td><td></td></tr></table></div>
"""
st.components.v1.html(html_template, height=600, scrolling=True)

if final_rows:
    df_xlsx = generate_excel(final_rows, days_count, start_date, client_name, ", ".join(sorted(list(all_secs))), ", ".join(list(all_media)), {"media_total": media_total, "prod_cost": prod_cost, "vat": vat, "grand_total": grand_total})
    st.download_button(label="📥 下載 Excel 報表 (.xlsx)", data=df_xlsx.getvalue(), file_name=f"CueSheet_{client_name}_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
