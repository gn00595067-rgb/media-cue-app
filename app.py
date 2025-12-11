import streamlit as st
import pandas as pd
import math
from datetime import timedelta, datetime

# ==========================================
# 1. 基礎資料與設定 (Configuration)
# ==========================================

# 2026 全家店數資料 (根據你的描述)
STORE_COUNTS = {
    "北區": "1649店", # 北北基 + 東 (依描述歸類，若有誤可調整)
    "桃竹苗": "779店",
    "中區": "839店",
    "雲嘉南": "499店",
    "高屏": "490店",
    "東區": "181店",
    "全省": "4437店" # 假設加總，或需填入官方數字
}

# 媒體選項
MEDIA_OPTIONS = ["全家便利商店通路廣播廣告", "全家便利商店新鮮視", "家樂福"]
REGIONS_FM = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區", "全省"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]

# --- [關鍵] 報價與折扣表 (請依照 Excel 2026 企頻報價填入真實數字) ---
# 這裡我先用 0 或 10000 當作範例，你需要修改這裡
PRICING_TABLE = {
    "全家廣播": {
        "北區": 100000, "桃竹苗": 50000, "中區": 60000, "雲嘉南": 40000, 
        "高屏": 40000, "東區": 20000, "全省": 300000
    },
    "新鮮視": {
        "北區": 120000, "桃竹苗": 60000, "中區": 70000, "雲嘉南": 50000, 
        "高屏": 50000, "東區": 25000, "全省": 350000
    },
    "家樂福": {
        "全省量販": 150000, 
        "全省超市": 80000
    }
}

# 秒數折扣表 (秒數: 折扣率) - 範例數據
DISCOUNT_TABLE = {
    5: 0.5, 10: 0.6, 15: 0.7, 20: 0.8, 30: 1.0, 
    40: 1.3, 60: 2.0
}

def get_discount(seconds):
    """取得秒數折扣，若無對應秒數則找大一階的"""
    sorted_secs = sorted(DISCOUNT_TABLE.keys())
    for s in sorted_secs:
        if s >= seconds:
            return DISCOUNT_TABLE[s]
    return 1.0 # fallback

def get_store_count_text(region):
    if region == "北區": return f"北北基 {STORE_COUNTS['北區']}"
    if region == "桃竹苗": return f"桃竹苗 {STORE_COUNTS['桃竹苗']}"
    if region == "中區": return f"中彰投 {STORE_COUNTS['中區']}"
    if region == "雲嘉南": return f"雲嘉南 {STORE_COUNTS['雲嘉南']}"
    if region == "高屏": return f"高高屏 {STORE_COUNTS['高屏']}"
    if region == "東區": return f"宜花東 {STORE_COUNTS['東區']}"
    return "全省門市"

# ==========================================
# 2. 邏輯運算核心 (Logic Core)
# ==========================================

def calculate_spots_distribution(total_spots, days):
    """
    檔次分配邏輯：
    1. 盡量平均
    2. 盡量排偶數或是5的倍數
    3. 前半多後半少
    """
    if days == 0: return []
    
    base_avg = total_spots / days
    schedule = [0] * days
    remaining = total_spots
    
    # 初步分配策略：優先滿足偶數邏輯
    # 這裡使用一個啟發式算法
    for i in range(days):
        # 簡單算法：先求平均，然後嘗試變成偶數
        if i < days // 2: # 前半段
            val = math.ceil(remaining / (days - i))
        else:
            val = math.floor(remaining / (days - i))
        
        # 調整為偶數 (若 val 是奇數，加 1 或 減 1)
        if val % 2 != 0 and val > 1:
            if remaining - (val + 1) >= 0:
                val += 1
            else:
                val -= 1
        
        # 邊界檢查
        if val <= 0 and remaining > 0: val = 1
        if val > remaining: val = remaining
        
        schedule[i] = val
        remaining -= val
        
    # 如果還有剩餘，從前面開始補
    i = 0
    while remaining > 0:
        schedule[i] += 1
        remaining -= 1
        i = (i + 1) % days
        
    return schedule

def calculate_line_item(platform, region, seconds, allocated_budget, days_count):
    """
    計算單一列的數據：檔次、費用、Rate
    """
    # 1. 決定定價 (List Price)
    list_price = 0
    price_key = ""
    if "廣播" in platform:
        price_key = "全家廣播"
    elif "新鮮視" in platform:
        price_key = "新鮮視"
    elif "家樂福" in platform:
        # 家樂福特殊處理，這裡假設分配到的預算還要再拆給量販跟超市
        # 為簡化，此函數只算單一細項，外層需呼叫兩次
        price_key = "家樂福" 

    if platform == "家樂福(量販)":
        list_price = PRICING_TABLE["家樂福"]["全省量販"]
    elif platform == "家樂福(超市)":
        list_price = PRICING_TABLE["家樂福"]["全省超市"]
    else:
        list_price = PRICING_TABLE.get(price_key, {}).get(region, 0)

    # 2. 取得折扣
    disc = get_discount(seconds)
    
    # 3. 逆推檔次 (Target Spots)
    # 公式：Cost = (Price / 720) * Spots * Disc * (1.1 if Package & Spots < 720)
    # 簡化逆推：我們先忽略 1.1 的 Package 規則來算大概檔次，再微調
    
    base_unit_cost = (list_price / 720) * disc
    if base_unit_cost == 0: return None # 避免除以零

    target_spots = math.ceil(allocated_budget / base_unit_cost)
    
    # 4. 根據 Package 規則調整 (全家全省區)
    is_fm_national = ("全家" in platform or "新鮮視" in platform) and region == "全省"
    
    # 迴圈微調以符合預算 (因為有 1.1 的跳變)
    final_spots = target_spots
    calculated_cost = 0
    
    while True:
        multiplier = 1.0
        if is_fm_national and final_spots < 720:
            multiplier = 1.1
            
        calculated_cost = (list_price / 720) * final_spots * disc * multiplier
        
        # 邏輯：費用要大於預算，且最好在 5% 以內 (這裡只確保大於預算)
        if calculated_cost >= allocated_budget:
            break
        final_spots += 1
        
    # 5. 格式化輸出資料
    rate_net = calculated_cost # 這裡依照你的描述，Rate(Net)欄位顯示總金額
    
    # 每日檔次分配
    daily_spots = calculate_spots_distribution(final_spots, days_count)
    
    return {
        "platform": platform,
        "region": region,
        "seconds": seconds,
        "spots": final_spots,
        "total_cost": calculated_cost,
        "daily_schedule": daily_spots,
        "program": get_store_count_text(region) if "家樂福" not in platform else "全省",
        "day_part": "07:00-23:00" if "廣播" in platform or "新鮮視" in platform else ("09:00-23:00" if "量販" in platform else "00:00-24:00"),
        "is_package": is_fm_national and final_spots < 720
    }

# ==========================================
# 3. 前端介面 (UI)
# ==========================================

st.set_page_config(layout="wide", page_title="Cue Sheet Generator")

st.title("媒體 Cue 表生成器 (Beta)")
st.markdown("目標：模擬 Excel 邏輯，自動分配預算並生成報表。")

with st.sidebar:
    st.header("1. 基本資料")
    client_name = st.text_input("客戶名稱", "範例客戶")
    start_date = st.date_input("走期開始日", datetime.today())
    end_date = st.date_input("走期結束日", datetime.today() + timedelta(days=13))
    
    if start_date > end_date:
        st.error("結束日期必須晚於開始日期")
        days_count = 0
    else:
        days_count = (end_date - start_date).days + 1
        st.info(f"走期共 {days_count} 天")

    total_budget = st.number_input("總預算 (未稅)", value=500000, step=10000)

st.header("2. 媒體投放設定")

# 儲存使用者的選擇
selections = []

# --- 全家廣播 ---
with st.expander("全家便利商店通路廣播廣告", expanded=True):
    fm_radio_on = st.checkbox("購買全家廣播")
    if fm_radio_on:
        fm_regions = st.multiselect("廣播-區域 (可多選)", REGIONS_FM, default=["全省"])
        fm_secs = st.multiselect("廣播-秒數", DURATIONS, default=[20])
        fm_budget_alloc = st.slider("廣播-預算佔比 (%)", 0, 100, 50)
        selections.append({
            "type": "全家廣播",
            "regions": fm_regions,
            "seconds": fm_secs,
            "budget_percent": fm_budget_alloc
        })

# --- 全家新鮮視 ---
with st.expander("全家便利商店新鮮視"):
    fv_on = st.checkbox("購買新鮮視")
    if fv_on:
        fv_regions = st.multiselect("新鮮視-區域 (可多選)", REGIONS_FM, default=["全省"])
        fv_secs = st.multiselect("新鮮視-秒數", DURATIONS, default=[10])
        fv_budget_alloc = st.slider("新鮮視-預算佔比 (%)", 0, 100, 30)
        selections.append({
            "type": "新鮮視",
            "regions": fv_regions,
            "seconds": fv_secs,
            "budget_percent": fv_budget_alloc
        })

# --- 家樂福 ---
with st.expander("家樂福"):
    carrefour_on = st.checkbox("購買家樂福")
    if carrefour_on:
        st.write("區域：固定為全省 (包含量販與超市)")
        cf_secs = st.multiselect("家樂福-秒數", DURATIONS, default=[10])
        cf_budget_alloc = st.slider("家樂福-預算佔比 (%)", 0, 100, 20)
        selections.append({
            "type": "家樂福",
            "regions": ["全省"], # 邏輯上雖然分量販超市，但預算分配算一次
            "seconds": cf_secs,
            "budget_percent": cf_budget_alloc
        })

# ==========================================
# 4. 運算與生成 (Execution)
# ==========================================

if st.button("生成 Cue 表"):
    # 1. 檢查輸入
    if not selections:
        st.error("請至少選擇一種媒體")
        st.stop()

    # 2. 預算分配計算
    line_items = []
    
    # 正規化預算比例 (若總和不為100，則按比例縮放，或直接用輸入值)
    total_percent = sum(s["budget_percent"] for s in selections)
    
    grand_total_calculated = 0
    production_cost = 10000
    
    all_seconds_set = set()
    selected_mediums = set()

    for sel in selections:
        # 該媒體總預算
        if total_percent > 0:
            media_budget = total_budget * (sel["budget_percent"] / total_percent)
        else:
            media_budget = 0
            
        # 該媒體下的細項數 (區域數 * 秒數種類)
        # 備註：家樂福雖然選全省，但會生成 量販+超市 兩條
        sub_items_count = 0
        if sel["type"] == "家樂福":
            sub_items_count = len(sel["seconds"]) * 2 # 量販+超市
        else:
            sub_items_count = len(sel["regions"]) * len(sel["seconds"])
            
        if sub_items_count == 0: continue

        # 平均分配預算給每個細項 (也可改為讓業務針對細項微調，這裡先平均)
        item_budget = media_budget / sub_items_count
        
        # 紀錄 Header 用
        selected_mediums.add("全家便利商店通路廣播廣告" if sel["type"]=="全家廣播" else ("全家便利商店新鮮視" if sel["type"]=="新鮮視" else "家樂福"))
        for s in sel["seconds"]:
            all_seconds_set.add(f"{s}秒")

        # 生成細項
        for sec in sel["seconds"]:
            if sel["type"] == "家樂福":
                # 家樂福強制生成兩條
                item1 = calculate_line_item("家樂福(量販)", "全省", sec, item_budget, days_count)
                item2 = calculate_line_item("家樂福(超市)", "全省", sec, item_budget, days_count)
                if item1: line_items.append(item1)
                if item2: line_items.append(item2)
            else:
                platform_name = "全家便利商店通路廣播廣告" if sel["type"] == "全家廣播" else "全家便利商店新鮮視"
                for reg in sel["regions"]:
                    item = calculate_line_item(platform_name, reg, sec, item_budget, days_count)
                    if item: line_items.append(item)

    # 3. 計算總金額
    items_total = sum(item["total_cost"] for item in line_items)
    vat = (items_total + production_cost) * 0.05
    grand_total = items_total + production_cost + vat
    
    # 折扣顯示 (給業務看)
    discount_val = grand_total - (total_budget * 1.05) # 粗略估算
    
    st.success(f"計算完成！ 預算: {total_budget:,} | 媒體費用(未稅): {int(items_total):,} | 總計(含稅): {int(grand_total):,}")
    
    # ==========================================
    # 5. HTML 表格生成 (Output)
    # ==========================================
    
    # 準備日期 Header
    date_headers = ""
    current = start_date
    for i in range(days_count):
        date_str = current.strftime("%m/%d")
        date_headers += f"<th class='date-col'>{date_str}</th>"
        current += timedelta(days=1)

    # 準備內容 Rows
    rows_html = ""
    for idx, item in enumerate(line_items):
        daily_cells = ""
        for spots in item["daily_schedule"]:
            daily_cells += f"<td class='schedule-cell'>{spots}</td>"
            
        # 處理 Packge Cost 顯示邏輯
        package_note = "(Package)" if item["is_package"] else ""
        
        rows_html += f"""
        <tr>
            <td style='text-align:left;'>{item['platform']}</td>
            <td>{item['region']}</td>
            <td>{item['program']}</td>
            <td>{item['day_part']}</td>
            <td>{item['seconds']}</td>
            {daily_cells}
            <td style='font-weight:bold;'>{item['spots']}</td>
            <td style='text-align:right;'>{int(item['total_cost']):,}<br><span style='font-size:10px'>{package_note}</span></td>
        </tr>
        """

    # 準備 Header 資訊
    product_str = "、".join(sorted(list(all_seconds_set)))
    medium_str = "、".join(list(selected_mediums))
    period_str = f"{start_date.strftime('%Y/%m/%d')} ~ {end_date.strftime('%Y/%m/%d')}"

    # CSS 樣式 (模仿 Excel)
    html_template = f"""
    <style>
        .cue-table {{
            width: 100%;
            border-collapse: collapse;
            font-family: 'Arial', 'Microsoft JhengHei', sans-serif;
            font-size: 12px;
        }}
        .cue-table th, .cue-table td {{
            border: 1px solid #000;
            padding: 4px;
            text-align: center;
        }}
        .header-section {{
            background-color: #f0f0f0; /* 淺灰底 */
            text-align: left;
        }}
        .col-header {{
            background-color: #d9e1f2; /* Excel 藍色 header */
            font-weight: bold;
        }}
        .date-col {{
            background-color: #d9e1f2;
            writing-mode: vertical-rl; /* 直式日期 */
            transform: rotate(180deg);
            min-width: 25px;
        }}
        .schedule-cell {{
            background-color: #fff;
        }}
        .total-row {{
            background-color: #ffffcc; /* 淺黃底 */
            font-weight: bold;
        }}
    </style>

    <div style="border: 2px solid #000; padding: 10px; background: white;">
        <table class="cue-table">
            <tr>
                <td colspan="5" class="header-section" style="border:none;">
                    <b>Client：</b>{client_name}<br>
                    <b>Product：</b>{product_str}<br>
                    <b>Period：</b>{period_str}<br>
                    <b>Medium：</b>{medium_str}
                </td>
                <td colspan="{days_count + 2}" style="border:none;"></td>
            </tr>
            
            <tr class="col-header">
                <th>Station</th>
                <th>Location</th>
                <th>Program</th>
                <th>Day-part</th>
                <th>Size</th>
                {date_headers}
                <th>Total<br>Spots</th>
                <th>Rate (Net)</th>
            </tr>
            
            {rows_html}
            
            <tr class="total-row">
                <td colspan="5" style="text-align:right;">Media Total</td>
                <td colspan="{days_count}"></td>
                <td>{sum(i['spots'] for i in line_items)}</td>
                <td style="text-align:right;">{int(items_total):,}</td>
            </tr>
            <tr>
                <td colspan="5" style="text-align:right;">Production Cost</td>
                <td colspan="{days_count + 2}" style="text-align:right;">{production_cost:,}</td>
            </tr>
            <tr>
                <td colspan="5" style="text-align:right;">5% VAT</td>
                <td colspan="{days_count + 2}" style="text-align:right;">{int(vat):,}</td>
            </tr>
            <tr class="total-row" style="font-size:14px; border-top: 2px double #000;">
                <td colspan="5" style="text-align:right;">Grand Total</td>
                <td colspan="{days_count + 2}" style="text-align:right;">NT$ {int(grand_total):,}</td>
            </tr>
        </table>
    </div>
    """
    
    st.markdown("### 預覽 Cue 表")
    st.components.v1.html(html_template, height=600, scrolling=True)
    
    st.info("提示：若要調整金額，請調整上方的「預算佔比」或「總預算」，系統會自動重新計算檔次。")
