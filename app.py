import streamlit as st
import pandas as pd
import math
import io
import xlsxwriter
from datetime import timedelta, datetime

# ==========================================
# 1. 基礎資料與設定
# ==========================================

STORE_COUNTS = {
    "全省": "4,437店",
    "北區": "北北基 1,649店",
    "桃竹苗": "桃竹苗 779店",
    "中區": "中彰投 839店",
    "雲嘉南": "雲嘉南 499店",
    "高屏": "高高屏 490店",
    "東區": "宜花東 181店",
    "新鮮視_全省": "3,124面",
    "新鮮視_北區": "北北基 1,127面",
    "新鮮視_桃竹苗": "桃竹苗 616面",
    "新鮮視_中區": "中彰投 528面",
    "新鮮視_雲嘉南": "雲嘉南 365面",
    "新鮮視_高屏": "高高屏 405面",
    "新鮮視_東區": "宜花東 83面",
}

REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]

# 價格資料庫 (修正：明確定義 Std_Spots)
PRICING_DB = {
    "全家廣播": {
        "Std_Spots": 480, # 基準檔次 480
        # 格式: [List Price, Net Price]
        "全省": [400000, 320000], 
        "北區": [250000, 200000], "桃竹苗": [150000, 120000],
        "中區": [150000, 120000], "雲嘉南": [100000, 80000], 
        "高屏": [100000, 80000], "東區": [62500, 50000]
    },
    "新鮮視": {
        "Std_Spots": 504, # 基準檔次 504
        "全省": [150000, 120000], 
        "北區": [150000, 120000], "桃竹苗": [120000, 96000],
        "中區": [90000, 72000], "雲嘉南": [75000, 60000], 
        "高屏": [75000, 60000], "東區": [45000, 36000]
    },
    "家樂福": {
        # 家樂福特殊：直接定義 Net Unit
        "量販_全省": {"List": 310000, "Net_Unit": 595},
        "超市_全省": {"List": 100000, "Net_Unit": 111}
    }
}

# 秒數折扣係數表
DISCOUNT_TABLE = {5: 0.5, 10: 0.6, 15: 0.7, 20: 0.8, 25: 0.9, 30: 1.0, 35: 1.15, 40: 1.3, 45: 1.5, 60: 2.0}

def get_discount(seconds):
    if seconds in DISCOUNT_TABLE: return DISCOUNT_TABLE[seconds]
    for s in sorted(DISCOUNT_TABLE.keys()):
        if s >= seconds: return DISCOUNT_TABLE[s]
    return 1.0

def calculate_schedule(total_spots, days):
    if days == 0: return []
    half_spots = total_spots // 2
    schedule = [0] * days
    base = half_spots // days
    for i in range(days): schedule[i] = base
    remaining = half_spots % days
    for i in range(remaining): schedule[i] += 1
    final_schedule = [x * 2 for x in schedule]
    current_sum = sum(final_schedule)
    diff = total_spots - current_sum
    if diff > 0: final_schedule[0] += diff
    return final_schedule

# ==========================================
# 2. UI 設定
# ==========================================

st.set_page_config(layout="wide", page_title="Cue Sheet Generator Final")
st.title("📺 媒體 Cue 表生成器 (金額計算修正版)")

# --- 1. 基本資料 (移至主畫面) ---
with st.container():
    st.markdown("### 1. 基本資料設定")
    with st.expander("📝 點擊展開/收合基本資料", expanded=True):
        col_b1, col_b2 = st.columns([1, 1])
        with col_b1:
            client_name = st.text_input("客戶名稱", "萬國通路")
            start_date = st.date_input("開始日", datetime(2025, 1, 1))
        with col_b2:
            total_budget_input = st.number_input("總預算 (未稅)", value=1140000, step=10000)
            end_date = st.date_input("結束日", datetime(2025, 1, 31))
        
        days_count = (end_date - start_date).days + 1
        st.info(f"📅 走期共 **{days_count}** 天")

# --- 2. 媒體設定 (瀑布流邏輯) ---
config_media = {}
st.markdown("### 2. 媒體投放設定 (連動總和 100%)")

col_m1, col_m2, col_m3 = st.columns(3)
remaining_global_share = 100 

# 全家廣播
with col_m1:
    st.markdown("#### 📻 全家廣播")
    fm_act = st.checkbox("開啟", value=True, key="fm_act")
    fm_data = None
    if fm_act:
        is_nat = st.checkbox("全省聯播", value=True, key="fm_nat")
        regs = ["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, key="fm_reg")
        _secs_input = st.multiselect("秒數", DURATIONS, default=[20], key="fm_sec")
        secs = sorted(_secs_input)
        share = st.slider("預算佔比%", 0, remaining_global_share, min(70, remaining_global_share), key="fm_share")
        remaining_global_share -= share
        sec_shares = {}
        if len(secs) > 1:
            st.caption("各秒數佔比")
            ls = 100
            for i, s in enumerate(secs[:-1]):
                v = st.slider(f"{s}秒佔比", 0, ls, int(ls/2), key=f"fm_s_{s}")
                sec_shares[s] = v; ls -= v
            sec_shares[secs[-1]] = ls
            st.write(f"🔹 {secs[-1]}秒: {ls}%")
        elif secs: sec_shares[secs[0]] = 100
        fm_data = {"is_national": is_nat, "regions": regs, "seconds": secs, "share": share, "sec_shares": sec_shares}

# 新鮮視
with col_m2:
    st.markdown("#### 📺 新鮮視")
    fv_act = st.checkbox("開啟", value=True, key="fv_act")
    fv_data = None
    if fv_act:
        is_nat = st.checkbox("全省聯播 ", value=False, key="fv_nat")
        regs = ["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, default=["北區", "桃竹苗"], key="fv_reg")
        _secs_input = st.multiselect("秒數", DURATIONS, default=[5], key="fv_sec")
        secs = sorted(_secs_input)
        limit = remaining_global_share
        default_val = min(20, limit)
        share = st.slider("預算佔比% ", 0, limit, default_val, key="fv_share")
        remaining_global_share -= share
        sec_shares = {}
        if len(secs) > 1:
            st.caption("各秒數佔比")
            ls = 100
            for i, s in enumerate(secs[:-1]):
                v = st.slider(f"{s}秒佔比 ", 0, ls, int(ls/2), key=f"fv_s_{s}")
                sec_shares[s] = v; ls -= v
            sec_shares[secs[-1]] = ls
            st.write(f"🔹 {secs[-1]}秒: {ls}%")
        elif secs: sec_shares[secs[0]] = 100
        fv_data = {"is_national": is_nat, "regions": regs, "seconds": secs, "share": share, "sec_shares": sec_shares}

# 家樂福
with col_m3:
    st.markdown("#### 🛒 家樂福")
    cf_act = st.checkbox("開啟", value=True, key="cf_act")
    cf_data = None
    if cf_act:
        st.write("區域：全省")
        _secs_input = st.multiselect("秒數", DURATIONS, default=[20], key="cf_sec")
        secs = sorted(_secs_input)
        share = remaining_global_share
        st.info(f"預算佔比: **{share}%** (自動填滿)")
        st.progress(share / 100.0 if share <= 100 else 1.0)
        sec_shares = {}
        if len(secs) > 1:
            st.caption("各秒數佔比")
            ls = 100
            for i, s in enumerate(secs[:-1]):
                v = st.slider(f"{s}秒佔比  ", 0, ls, int(ls/2), key=f"cf_s_{s}")
                sec_shares[s] = v; ls -= v
            sec_shares[secs[-1]] = ls
            st.write(f"🔹 {secs[-1]}秒: {ls}%")
        elif secs: sec_shares[secs[0]] = 100
        cf_data = {"regions": ["全省"], "seconds": secs, "share": share, "sec_shares": sec_shares}

if fm_data: config_media["全家廣播"] = fm_data
if fv_data: config_media["新鮮視"] = fv_data
if cf_data: config_media["家樂福"] = cf_data

# ==========================================
# 3. 計算邏輯 (核心修正：正確使用 Std_Spots 和 Net Price)
# ==========================================

final_rows = []
all_secs = set()
all_media = set()

if sum(m["share"] for m in config_media.values()) > 0:
    for m_type, cfg in config_media.items():
        media_budget = total_budget_input * (cfg["share"] / 100.0)
        all_media.add(m_type)
        
        for sec, sec_share in cfg["sec_shares"].items():
            all_secs.add(f"{sec}秒")
            sec_budget = media_budget * (sec_share / 100.0)
            if sec_budget <= 0: continue
            
            discount = get_discount(sec)
            
            if m_type in ["全家廣播", "新鮮視"]:
                db = PRICING_DB[m_type]
                std_spots = db["Std_Spots"] # 動態取得 480 或 504
                
                calc_regions = ["全省"] if cfg["is_national"] else cfg["regions"]
                display_regions = REGIONS_ORDER if cfg["is_national"] else cfg["regions"]
                
                combined_unit_net = 0
                for reg in calc_regions:
                    # 使用 Net Price (index 1) 計算 unit_net
                    net_price_total = db[reg][1]
                    unit_net = (net_price_total / std_spots) * discount
                    combined_unit_net += unit_net
                
                if combined_unit_net == 0: continue
                
                target_spots = math.ceil(sec_budget / combined_unit_net)
                if target_spots % 2 != 0: target_spots += 1 
                if target_spots == 0: target_spots = 2
                
                daily_sch = calculate_schedule(target_spots, days_count)
                
                pkg_cost_total = 0
                if cfg["is_national"]:
                    # 使用 Net Price (index 1) 計算打包總價
                    nat_net_total = db["全省"][1] 
                    mult = 1.1 if target_spots < 720 else 1.0 # 這裡的 720 是判斷是否加價的門檻，維持不變或需確認
                    pkg_cost_total = (nat_net_total / std_spots) * target_spots * discount * mult

                for reg in display_regions:
                    # 使用 Net Price (index 1) 計算各區域顯示金額
                    # 如果是全省打包，這裡算出的是該區域分攤的 Net
                    # 如果是區域購買，這裡算出的就是該區域的 Net
                    reg_net_total = db.get(reg, [0,0])[1] if cfg["is_national"] else db[reg][1]
                    rate_val = int(round((reg_net_total / std_spots) * target_spots * discount))
                    
                    real_c = int(round(combined_unit_net * target_spots)) if (not cfg["is_national"] or reg == "北區") else 0
                    pkg_val_if_nat = int(round(pkg_cost_total)) if (cfg["is_national"] and reg == "北區") else 0
                    
                    prog_name = STORE_COUNTS.get(reg, reg)
                    if m_type == "新鮮視": prog_name = STORE_COUNTS.get(f"新鮮視_{reg}", reg)
                    
                    # 顯示邏輯：
                    # 全省打包 -> 顯示全省總價 (pkg_val_if_nat)
                    # 區域購買 -> 顯示該區域價 (rate_val)
                    if cfg["is_national"]:
                         pkg_display_val = pkg_val_if_nat
                    else:
                         pkg_display_val = rate_val

                    final_rows.append({
                        "media": m_type, "region": reg, 
                        "location": f"{reg.replace('區', '')}區-{reg}" if m_type=="全家廣播" else f"{reg.replace('區', '')}區-{reg}",
                        "program": prog_name, "daypart": "00:00-24:00" if m_type=="全家廣播" else "07:00-22:00",
                        "seconds": sec, "schedule": daily_sch, "spots": target_spots,
                        "rate_net": rate_val, 
                        "pkg_cost": pkg_val_if_nat,
                        "pkg_display_val": pkg_display_val,
                        "is_pkg_start": (cfg["is_national"] and reg == "北區"), 
                        "is_pkg_member": cfg["is_national"], 
                        "real_cost": real_c
                    })

            elif m_type == "家樂福":
                db = PRICING_DB["家樂福"]
                unit_hyp = db["量販_全省"]["Net_Unit"] * discount
                unit_sup = db["超市_全省"]["Net_Unit"] * discount
                combined = unit_hyp + unit_sup
                
                target_spots = math.ceil(sec_budget / combined)
                if target_spots % 2 != 0: target_spots += 1
                if target_spots == 0: target_spots = 2

                sch = calculate_schedule(target_spots, days_count)
                
                # 家樂福 Rate Net 顯示邏輯：List Price / 720 (若維持原邏輯)
                # 若要改成 Net，則直接用 Net Unit * spots
                # 這裡假設 rate (Net) 欄位顯示 List 換算，Package-cost 顯示 Net
                rate_hyp = int(round((db["量販_全省"]["List"]/720.0)*target_spots*discount))
                rate_sup = int(round((db["超市_全省"]["List"]/720.0)*target_spots*discount))
                
                final_rows.append({
                    "media": "家樂福", "region": "全省量販", "location": "全省量販", "program": "67店",
                    "daypart": "09:00-23:00", "seconds": sec, "schedule": sch, "spots": target_spots,
                    "rate_net": rate_hyp,
                    "pkg_cost": 0, "pkg_display_val": int(round(unit_hyp * target_spots)),
                    "is_pkg_start": False, "is_pkg_member": False, 
                    "real_cost": int(round(unit_hyp * target_spots))
                })
                final_rows.append({
                    "media": "家樂福", "region": "全省超市", "location": "全省超市", "program": "250店",
                    "daypart": "00:00-24:00", "seconds": sec, "schedule": sch, "spots": target_spots,
                    "rate_net": rate_sup,
                    "pkg_cost": 0, "pkg_display_val": int(round(unit_sup * target_spots)),
                    "is_pkg_start": False, "is_pkg_member": False, 
                    "real_cost": int(round(unit_sup * target_spots))
                })

media_order_map = {"全家廣播": 1, "新鮮視": 2, "家樂福": 3}
final_rows.sort(key=lambda x: media_order_map.get(x['media'], 99))

def parse_sec_int(s):
    return int(s.replace("秒", ""))
sorted_secs_list = sorted(list(all_secs), key=parse_sec_int)
product_str = "、".join(sorted_secs_list)

media_total = sum(r["pkg_display_val"] for r in final_rows if not r['is_pkg_member'] or r['is_pkg_start'])

prod_cost = 10000
vat = int(round((media_total + prod_cost) * 0.05))
grand_total = media_total + prod_cost + vat
discount_ratio_str = f"{(total_budget_input / grand_total * 100):.1f}%" if grand_total > 0 else "N/A"

# ==========================================
# 4. 生成 HTML 預覽
# ==========================================

def generate_html_preview(rows, days_cnt, start_dt, c_name, products, totals_data):
    used_media = sorted(list(set(r['media'] for r in rows)), key=lambda x: media_order_map.get(x, 99))
    mediums_str = "、".join(used_media)

    date_header_row1 = f"<th class='header-blue' colspan='{days_cnt}'>{start_dt.month}月</th>"
    date_header_row2 = ""
    date_header_row3 = ""
    
    curr = start_dt
    weekdays_map = ["一", "二", "三", "四", "五", "六", "日"]
    
    for i in range(days_cnt):
        wd = curr.weekday()
        cls = "header-yellow" if wd >= 5 else "header-blue"
        date_header_row2 += f"<th class='{cls}'>{curr.day}</th>"
        date_header_row3 += f"<th class='{cls}'>{weekdays_map[wd]}</th>"
        curr += timedelta(days=1)
        
    data_rows_html = ""
    i = 0
    while i < len(rows):
        row = rows[i]
        j = i + 1
        while j < len(rows) and rows[j]['media'] == row['media'] and rows[j]['seconds'] == row['seconds']:
            j += 1
        group_size = j - i
        
        m_name = row['media']
        if "全家廣播" in m_name: m_name = "全家便利商店<br>通路廣播廣告"
        if "新鮮視" in m_name: m_name = "全家便利商店<br>新鮮視廣告"
        
        for k in range(group_size):
            r_data = rows[i+k]
            tr = "<tr>"
            if k == 0:
                tr += f"<td rowspan='{group_size}' class='align-left'>{m_name}</td>"
            
            loc_txt = r_data['location']
            if "北北基" in loc_txt and "廣播" in r_data['media']: loc_txt = "北區-北北基+東"
            
            tr += f"<td>{loc_txt}</td>"
            tr += f"<td>{r_data['program']}</td>"
            tr += f"<td>{r_data['daypart']}</td>"
            tr += f"<td>{r_data['seconds']}秒</td>"
            tr += f"<td class='align-right'>{r_data['rate_net']:,}</td>"
            
            if row['is_pkg_start']:
                if k == 0:
                    tr += f"<td rowspan='{group_size}' class='align-right'>{row['pkg_display_val']:,}</td>"
            elif row['is_pkg_member']:
                pass
            else:
                val = r_data['pkg_display_val']
                val_str = f"{val:,}"
                tr += f"<td class='align-right'>{val_str}</td>"
            
            for s_val in r_data['schedule']:
                tr += f"<td>{s_val}</td>"
                
            tr += f"<td class='cell-yellow'>{r_data['spots']}</td>"
            tr += "</tr>"
            data_rows_html += tr
        i = j
    
    total_rate_display = sum(r['rate_net'] for r in rows)

    # CSS 保持高對比 + 不透明
    css_style = """
    <style>
        .preview-table {
            width: 100%;
            border-collapse: collapse;
            font-family: "Microsoft JhengHei", "Arial", sans-serif;
            font-size: 13px;
            color: #000;
            min-width: 1200px;
            background-color: #ffffff;
        }
        .preview-table th, .preview-table td {
            border: 1px solid #555;
            padding: 8px;
            text-align: center;
            vertical-align: middle;
        }
        .header-blue { background-color: #2c3e50; color: white !important; font-weight: bold; }
        .header-yellow { background-color: #f1c40f; color: #000 !important; font-weight: bold; }
        .cell-yellow { background-color: #fff3cd; color: #000 !important; font-weight: bold; }
        .row-total { background-color: #d4edda; color: #000 !important; font-weight: bold; }
        .row-grand-total { background-color: #ffc107; color: #000 !important; font-weight: bold; font-size: 15px; border-top: 2px solid #000; }
        .align-left { text-align: left; }
        .align-right { text-align: right; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        tr:hover { background-color: #e6f7ff; }
    </style>
    """

    html = f"""
    {css_style}
    <div style="overflow-x: auto; width: 100%;">
        <table class="preview-table">
            <tr>
                <td colspan="5" class="align-left" style="background-color:#fff; border:none;">
                    <b>客戶名稱：</b> {c_name}<br>
                    <b>Product：</b> {products}<br>
                    <b>Period：</b> {start_dt.strftime('%Y. %m. %d')} - {end_date.strftime('%Y. %m. %d')}<br>
                    <b>Medium：</b> {mediums_str}
                </td>
                <td colspan="{days_cnt + 3}" style="background-color:#fff; border:none;"></td>
            </tr>
            <tr>
                <th colspan="7" style="border:none;"></th>
                {date_header_row1}
                <th style="border:none;"></th>
            </tr>
            <tr>
                <th rowspan="2" class="header-blue">Station</th>
                <th rowspan="2" class="header-blue">Location</th>
                <th rowspan="2" class="header-blue">Program</th>
                <th rowspan="2" class="header-blue">Day-part</th>
                <th rowspan="2" class="header-blue">Size</th>
                <th rowspan="2" class="header-blue">rate (Net)</th>
                <th rowspan="2" class="header-blue">Package-cost<br>(Net)</th>
                {date_header_row2}
                <th rowspan="2" class="header-blue">檔次</th>
            </tr>
            <tr>
                {date_header_row3}
            </tr>
            {data_rows_html}
            <tr class="row-total">
                <td colspan="5" class="align-right">Total</td>
                <td class="align-right">{total_rate_display:,}</td>
                <td class="align-right">{totals_data['media_total']:,}</td>
                <td colspan="{days_cnt}"></td>
                <td class="cell-yellow">{sum(r['spots'] for r in rows)}</td>
            </tr>
            <tr>
                <td colspan="6" class="align-right">製作</td>
                <td class="align-right">{totals_data['prod_cost']:,}</td>
                <td colspan="{days_cnt + 1}"></td>
            </tr>
            <tr>
                <td colspan="6" class="align-right">5% VAT</td>
                <td class="align-right">{totals_data['vat']:,}</td>
                <td colspan="{days_cnt + 1}"></td>
            </tr>
            <tr class="row-grand-total">
                <td colspan="6" class="align-right">Grand Total</td>
                <td class="align-right">{totals_data['grand_total']:,}</td>
                <td colspan="{days_cnt + 1}"></td>
            </tr>
        </table>
    </div>
    """
    return html

def generate_excel(rows, days_cnt, start_dt, c_name, products, totals_data):
    used_media = sorted(list(set(r['media'] for r in rows)), key=lambda x: media_order_map.get(x, 99))
    mediums = "、".join(used_media)
    
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Media Schedule")

    fmt_title = workbook.add_format({'font_size': 18, 'bold': True, 'align': 'center'})
    fmt_header_left = workbook.add_format({'align': 'left', 'valign': 'top', 'bold': True})
    fmt_col_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#4472C4', 'font_color': 'white', 'text_wrap': True, 'font_size': 10})
    fmt_date_wk = workbook.add_format({'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#4472C4', 'font_color': 'white'})
    fmt_date_we = workbook.add_format({'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFD966'}) 
    
    fmt_cell = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 10})
    fmt_cell_left = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'font_size': 10, 'text_wrap': True})
    fmt_num = workbook.add_format({'align': 'right', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0', 'font_size': 10})
    fmt_spots = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bold': True, 'bg_color': '#FFF2CC', 'font_size': 10})
    
    fmt_total = workbook.add_format({'align': 'right', 'valign': 'vcenter', 'border': 1, 'bold': True, 'bg_color': '#E2EFDA', 'num_format': '#,##0', 'font_size': 10})
    fmt_grand_total = workbook.add_format({'align': 'right', 'valign': 'vcenter', 'border': 1, 'bold': True, 'bg_color': '#FFC107', 'num_format': '#,##0', 'font_size': 10})

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

    worksheet.write(6, 6, f"{start_dt.month}月", fmt_cell)
    weekdays = ["一", "二", "三", "四", "五", "六", "日"]
    curr = start_dt
    for i in range(days_cnt):
        col_idx = 7 + i
        wd = curr.weekday()
        fmt = fmt_date_we if wd >= 5 else fmt_date_wk
        worksheet.write(7, col_idx, curr.day, fmt)
        worksheet.write(8, col_idx, weekdays[wd], fmt)
        curr += timedelta(days=1)

    headers = ["Station", "Location", "Program", "Day-part", "Size", "rate (Net)", "Package-cost\n(Net)"]
    for i, h in enumerate(headers):
        worksheet.write(8, i, h, fmt_col_header)
    
    last_col = 7 + days_cnt
    worksheet.write(8, last_col, "檔次", fmt_col_header)

    current_row = 9
    i = 0
    while i < len(rows):
        row = rows[i]
        j = i + 1
        while j < len(rows) and rows[j]['media'] == row['media'] and rows[j]['seconds'] == row['seconds']:
            j += 1
        group_size = j - i
        
        m_name = row['media']
        if "全家廣播" in m_name: m_name = "全家便利商店\n通路廣播廣告"
        if "新鮮視" in m_name: m_name = "全家便利商店\n新鮮視廣告"
        
        if group_size > 1:
            worksheet.merge_range(current_row, 0, current_row + group_size - 1, 0, m_name, fmt_cell_left)
        else:
            worksheet.write(current_row, 0, m_name, fmt_cell_left)
            
        for k in range(group_size):
            r_data = rows[i + k]
            r_idx = current_row + k
            
            loc_txt = r_data['location']
            if "北北基" in loc_txt and "廣播" in r_data['media']: loc_txt = "北區-北北基+東"
            
            worksheet.write(r_idx, 1, loc_txt, fmt_cell)
            worksheet.write(r_idx, 2, r_data['program'], fmt_cell)
            worksheet.write(r_idx, 3, r_data['daypart'], fmt_cell)
            worksheet.write(r_idx, 4, f"{r_data['seconds']}秒", fmt_cell)
            worksheet.write(r_idx, 5, r_data['rate_net'], fmt_num)
            
            if r_data['is_pkg_start']:
                 if k == 0 and group_size > 1:
                     worksheet.merge_range(current_row, 6, current_row + group_size - 1, 6, r_data['pkg_display_val'], fmt_num)
                 elif k == 0:
                     worksheet.write(r_idx, 6, r_data['pkg_display_val'], fmt_num)
            elif not r_data['is_pkg_member']:
                 worksheet.write(r_idx, 6, r_data['pkg_display_val'], fmt_num)

            for d_idx, s_val in enumerate(r_data['schedule']):
                worksheet.write(r_idx, 7 + d_idx, s_val, fmt_cell)
                
            worksheet.write(r_idx, last_col, r_data['spots'], fmt_spots)

        current_row += group_size
        i = j

    worksheet.write(current_row, 2, "Total", fmt_total)
    worksheet.write(current_row, 5, sum(r['rate_net'] for r in rows), fmt_total)
    worksheet.write(current_row, 6, totals_data['media_total'], fmt_total)
    
    total_spots_daily = [0] * days_cnt
    for r in rows:
        for idx, val in enumerate(r['schedule']):
            total_spots_daily[idx] += val
    for idx, val in enumerate(total_spots_daily):
        worksheet.write(current_row, 7+idx, val, fmt_cell)
    worksheet.write(current_row, last_col, sum(r['spots'] for r in rows), fmt_spots)
    
    current_row += 1
    worksheet.write(current_row, 6, "製作", fmt_cell)
    worksheet.write(current_row, 7, totals_data['prod_cost'], fmt_num)
    current_row += 1
    worksheet.write(current_row, 6, "5% VAT", fmt_cell)
    worksheet.write(current_row, 7, totals_data['vat'], fmt_num)
    current_row += 1
    worksheet.write(current_row, 6, "Grand Total", fmt_grand_total)
    worksheet.write(current_row, 7, totals_data['grand_total'], fmt_grand_total)

    worksheet.set_column('A:A', 20)
    worksheet.set_column('B:B', 15)
    worksheet.set_column('C:E', 12)
    worksheet.set_column('F:G', 12)
    worksheet.set_column(7, last_col, 4)

    workbook.close()
    return output

# ==========================================
# 5. 結果顯示與下載
# ==========================================

st.markdown("### 3. 計算結果摘要")
m1, m2, m3 = st.columns(3)
m1.metric("客戶預算", f"{total_budget_input:,}")
m2.metric("Cue表總金額 (含稅)", f"{grand_total:,}", delta=f"差異 +{grand_total - total_budget_input:,}")
m3.metric("預算/表價比 (折扣率)", discount_ratio_str)

st.markdown("### 4. Cue 表網頁預覽")

if final_rows:
    totals = {"media_total": media_total, "prod_cost": prod_cost, "vat": vat, "grand_total": grand_total}
    
    html_preview = generate_html_preview(final_rows, days_count, start_date, client_name, product_str, totals)
    st.components.v1.html(html_preview, height=600, scrolling=True)

    xlsx_data = generate_excel(final_rows, days_count, start_date, client_name, product_str, totals)
    
    st.download_button(
        label="📥 下載 Excel Cue表 (.xlsx)",
        data=xlsx_data.getvalue(),
        file_name=f"CueSheet_{client_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
