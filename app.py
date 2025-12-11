import streamlit as st
import pandas as pd
import math
from datetime import timedelta, datetime

# ==========================================
# 1. 基礎資料與設定 (Configuration)
# ==========================================

# 區域與店數對照 (依據提供資料更新)
STORE_COUNTS = {
    "北區": "北北基 1649店",
    "桃竹苗": "桃竹苗 779店",
    "中區": "中彰投 839店",
    "雲嘉南": "雲嘉南 499店",
    "高屏": "高高屏 490店",
    "東區": "宜花東 181店",
    # 新鮮視的區域店數略有不同，這裡做個對照
    "新鮮視_北區": "北北基 1127店",
    "新鮮視_桃竹苗": "桃竹苗 616店",
    "新鮮視_中區": "中彰投 528店",
    "新鮮視_雲嘉南": "雲嘉南 365店",
    "新鮮視_高屏": "高高屏 405店",
    "新鮮視_東區": "宜花東 83店",
}

# 6大區域標準順序
REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]

# 媒體基礎設定
MEDIA_TYPES = ["全家廣播", "新鮮視", "家樂福"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]

# --- 報價資料庫 (List Price & Net Price) ---
# 結構: [List_Price(定價), Net_Price(實收價)]
# 單位: 元
PRICING_DB = {
    "全家廣播": {
        "Base_Sec": 30, # 基準秒數
        "Std_Spots": 480, # 月標準檔次 (用於計算單檔價格)
        "全省": [400000, 320000],
        "北區": [250000, 200000],
        "桃竹苗": [150000, 120000],
        "中區": [150000, 120000],
        "雲嘉南": [100000, 80000],
        "高屏": [100000, 80000],
        "東區": [62500, 50000]
    },
    "新鮮視": {
        "Base_Sec": 10,
        "Std_Spots": 504, # 月標準檔次
        "全省": [150000, 120000],
        "北區": [150000, 120000],
        "桃竹苗": [120000, 96000],
        "中區": [90000, 72000],
        "雲嘉南": [75000, 60000],
        "高屏": [75000, 60000],
        "東區": [45000, 36000]
    },
    "家樂福": {
        "Base_Sec": 10, # 假設基準
        # 家樂福比較特殊，這裡存全省總價
        "量販_全省": {"List": 300000, "Net": 250000, "Std_Spots": 420},
        # 提示中沒給超市具體價格，這裡依比例暫估或設為0，實際應填入數字
        # 假設超市是打包在專案裡，或是另外計價。這裡先給一個預設值避免報錯
        "超市_全省": {"List": 100000, "Net": 80000, "Std_Spots": 720} 
    }
}

# 秒數折扣表 (秒數: 折扣率) - 範例數據 (需確認真實折扣)
# 這裡假設線性或依慣例，若有特定表格需更新
DISCOUNT_TABLE = {
    5: 0.5, 10: 0.6, 15: 0.7, 20: 0.8, 25: 0.9, 30: 1.0, 
    35: 1.15, 40: 1.3, 45: 1.5, 60: 2.0
}

def get_discount(seconds):
    """取得秒數折扣"""
    if seconds in DISCOUNT_TABLE:
        return DISCOUNT_TABLE[seconds]
    # Fallback: 簡單線性推估或找最近
    sorted_secs = sorted(DISCOUNT_TABLE.keys())
    for s in sorted_secs:
        if s >= seconds:
            return DISCOUNT_TABLE[s]
    return seconds / 30.0 # 極端情況

def calculate_schedule(total_spots, days):
    """
    檔次分配邏輯：
    1. 總數正確
    2. 盡量偶數
    3. 前多後少
    """
    if days == 0: return []
    
    schedule = [0] * days
    remaining = total_spots
    
    # 第一次分配：平均值的整數
    base = remaining // days
    for i in range(days):
        schedule[i] = base
    remaining -= (base * days)
    
    # 餘數分配：優先給前面，並試圖湊偶數
    # 這裡使用一個權重策略：前半段優先加
    idx = 0
    while remaining > 0:
        schedule[idx] += 1
        remaining -= 1
        idx = (idx + 1) % days

    # 微調策略：嘗試將奇數變偶數 (譬如 15, 15 -> 16, 14)
    # 從後面往前面借，或者前面往後面給
    for i in range(days - 1):
        if schedule[i] % 2 != 0:
            if schedule[i+1] > 0:
                # 把下一個的一檔移上來
                schedule[i] += 1
                schedule[i+1] -= 1
            elif schedule[i] > 0:
                 # 把自己的一檔移下去
                schedule[i] -= 1
                schedule[i+1] += 1
                
    return schedule

# ==========================================
# 2. 前端介面 (UI)
# ==========================================

st.set_page_config(layout="wide", page_title="Cue Sheet Generator v2")

# CSS 隱藏預設選單，讓介面更乾淨
st.markdown("""
<style>
    .reportview-container { margin-top: -2em; }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stSlider > div > div > div > div { background-color: #4CAF50; }
</style>
""", unsafe_allow_html=True)

st.title("媒體 Cue 表生成器 (Live Edit)")

# --- Sidebar Inputs ---
with st.sidebar:
    st.header("1. 基本資料")
    client_name = st.text_input("客戶名稱", "範例客戶")
    c1, c2 = st.columns(2)
    start_date = c1.date_input("開始日", datetime.today())
    end_date = c2.date_input("結束日", datetime.today() + timedelta(days=13))
    
    if start_date > end_date:
        st.error("日期錯誤")
        days_count = 0
    else:
        days_count = (end_date - start_date).days + 1
        st.caption(f"共 {days_count} 天")

    total_budget_input = st.number_input("總預算 (未稅)", value=500000, step=10000)

# 儲存設定的容器
config_media = {}

# --- 媒體設定區塊 ---
st.subheader("2. 媒體投放設定與預算分配")

# 建立 3 個 Column 分別放不同媒體
col_m1, col_m2, col_m3 = st.columns(3)

# 1. 全家廣播
with col_m1:
    st.markdown("### 全家廣播")
    fm_active = st.checkbox("開啟", key="fm_radio_active")
    
    if fm_active:
        # 全省 / 區域 互斥邏輯
        fm_is_national = st.checkbox("全省聯播 (廣播)", value=True, key="fm_nat_check")
        
        if fm_is_national:
            fm_regions = ["全省"] # 邏輯上選全省
            st.info("已選擇全省 (將展開為6區)")
        else:
            fm_regions = st.multiselect("選擇區域", REGIONS_ORDER, key="fm_reg_sel")
            
        fm_secs = st.multiselect("秒數", DURATIONS, default=[20], key="fm_sec_sel")
        
        # 預算佔比 (Level 1)
        fm_share = st.slider("廣播-總預算佔比%", 0, 100, 40, key="fm_share")
        
        # 秒數佔比 (Level 2) - 如果選了多個秒數
        fm_sec_shares = {}
        if len(fm_secs) > 1:
            st.markdown("---")
            st.caption("各秒數佔廣播預算比例")
            left_share = 100
            for i, sec in enumerate(fm_secs):
                if i == len(fm_secs) - 1:
                    val = left_share
                    st.write(f"{sec}秒: {val}% (自動填滿)")
                else:
                    val = st.slider(f"{sec}秒 佔比%", 0, left_share, int(left_share/2), key=f"fm_s_{sec}")
                    left_share -= val
                fm_sec_shares[sec] = val
        elif len(fm_secs) == 1:
            fm_sec_shares[fm_secs[0]] = 100

        config_media["全家廣播"] = {
            "is_national": fm_is_national,
            "regions": fm_regions,
            "seconds": fm_secs,
            "share": fm_share,
            "sec_shares": fm_sec_shares
        }

# 2. 新鮮視
with col_m2:
    st.markdown("### 新鮮視")
    fv_active = st.checkbox("開啟", key="fv_active")
    
    if fv_active:
        fv_is_national = st.checkbox("全省聯播 (新鮮視)", value=True, key="fv_nat_check")
        
        if fv_is_national:
            fv_regions = ["全省"]
            st.info("已選擇全省 (將展開為6區)")
        else:
            fv_regions = st.multiselect("選擇區域", REGIONS_ORDER, key="fv_reg_sel")
            
        fv_secs = st.multiselect("秒數", DURATIONS, default=[10], key="fv_sec_sel")
        
        fv_share = st.slider("新鮮視-總預算佔比%", 0, 100, 30, key="fv_share")
        
        fv_sec_shares = {}
        if len(fv_secs) > 1:
            st.markdown("---")
            st.caption("各秒數佔新鮮視預算比例")
            left_share = 100
            for i, sec in enumerate(fv_secs):
                if i == len(fv_secs) - 1:
                    val = left_share
                    st.write(f"{sec}秒: {val}% (自動填滿)")
                else:
                    val = st.slider(f"{sec}秒 佔比%", 0, left_share, int(left_share/2), key=f"fv_s_{sec}")
                    left_share -= val
                fv_sec_shares[sec] = val
        elif len(fv_secs) == 1:
            fv_sec_shares[fv_secs[0]] = 100
            
        config_media["新鮮視"] = {
            "is_national": fv_is_national,
            "regions": fv_regions,
            "seconds": fv_secs,
            "share": fv_share,
            "sec_shares": fv_sec_shares
        }

# 3. 家樂福
with col_m3:
    st.markdown("### 家樂福")
    cf_active = st.checkbox("開啟", key="cf_active")
    
    if cf_active:
        st.write("區域：固定為全省 (量販+超市)")
        cf_secs = st.multiselect("秒數", DURATIONS, default=[10], key="cf_sec_sel")
        
        cf_share = st.slider("家樂福-總預算佔比%", 0, 100, 30, key="cf_share")
        
        cf_sec_shares = {}
        if len(cf_secs) > 1:
            st.markdown("---")
            st.caption("各秒數佔家樂福預算比例")
            left_share = 100
            for i, sec in enumerate(cf_secs):
                if i == len(cf_secs) - 1:
                    val = left_share
                    st.write(f"{sec}秒: {val}% (自動填滿)")
                else:
                    val = st.slider(f"{sec}秒 佔比%", 0, left_share, int(left_share/2), key=f"cf_s_{sec}")
                    left_share -= val
                cf_sec_shares[sec] = val
        elif len(cf_secs) == 1:
            cf_sec_shares[cf_secs[0]] = 100
            
        config_media["家樂福"] = {
            "regions": ["全省"],
            "seconds": cf_secs,
            "share": cf_share,
            "sec_shares": cf_sec_shares
        }

# ==========================================
# 3. 核心計算邏輯 (The Solver)
# ==========================================

final_rows = []
all_selected_secs = set()
all_selected_media = set()

# 總預算正規化 (防止輸入超過100%)
total_share_sum = sum(m["share"] for m in config_media.values())
actual_total_budget = total_budget_input
if total_share_sum == 0 and config_media:
    st.warning("請分配預算比例")
    st.stop()

# 開始計算
for m_type, cfg in config_media.items():
    # 1. 計算該媒體分配到的錢
    # 如果使用者滑桿加總不為100，我們按比例分配輸入的總金額
    if total_share_sum > 0:
        media_budget = actual_total_budget * (cfg["share"] / total_share_sum)
    else:
        media_budget = 0
        
    all_selected_media.add(m_type)
    
    # 2. 針對該媒體內的每個秒數計算
    for sec, sec_share in cfg["sec_shares"].items():
        all_selected_secs.add(f"{sec}秒")
        
        # 該秒數分配到的錢
        sec_budget = media_budget * (sec_share / 100.0)
        if sec_budget <= 0: continue
            
        discount = get_discount(sec)
        
        # --- 處理全家體系 (廣播/新鮮視) ---
        if m_type in ["全家廣播", "新鮮視"]:
            db = PRICING_DB[m_type]
            std_spots = db["Std_Spots"]
            base_sec = db["Base_Sec"]
            
            # 準備區域列表 (如果是全省，邏輯上雖然算一次，但顯示要展開)
            if cfg["is_national"]:
                calc_regions = ["全省"] # 計算用
                display_regions = REGIONS_ORDER # 顯示用
            else:
                calc_regions = cfg["regions"]
                display_regions = cfg["regions"]
            
            # 計算 Unit Price (Net) 總和，用來推算在此預算下，這些區域能同時買幾檔
            # 邏輯：所有選定區域的檔次要一致 -> 總價 = 檔次 * Sum(各區單價)
            combined_unit_net = 0
            
            for reg in calc_regions:
                # 取得 Net Price (實收價)
                # 全省如果選了，就是拿全省的價格
                p_net = db[reg][1]
                # 計算單檔價格: (總價 / 標準月檔次) * (秒數 / 基準秒數) * 秒數折扣 ? 
                # 不，通常秒數折扣是打在總價上。
                # 單檔淨價 = (Net_Price / Std_Spots) * discount
                # 備註：秒數調整通常不是線性的，是看折扣表。若20秒折扣是0.8 (相對於30秒定價)，則價格 = 30秒價 * 0.8
                # 讓我們用定價邏輯修正： Net_Cost = Net_List * Discount
                # 但這裡是算「幾檔」。
                # 假設：實收價表是給基準秒數(如30s)的全月(如480檔)價格。
                # 單檔基準實收 = Net_Price / Std_Spots
                # 該秒數單檔實收 = 單檔基準實收 * discount
                
                # 若 discount 定義為：相對於基準價格的比例 (如20秒是30秒價格的80%)
                unit_net = (p_net / std_spots) * discount
                combined_unit_net += unit_net
            
            if combined_unit_net == 0: continue
                
            # 算出共用檔次
            target_spots = math.floor(sec_budget / combined_unit_net)
            
            # 確保檔次盡量是偶數
            if target_spots % 2 != 0:
                # 判斷往上加會不會爆預算太多，或是往下減
                # 這裡簡單處理：若 (target_spots+1) * cost < sec_budget * 1.05 -> 加
                if (target_spots + 1) * combined_unit_net <= sec_budget * 1.05:
                     target_spots += 1
                else:
                     target_spots -= 1
            if target_spots <= 0: target_spots = 1 # 至少一檔
            
            # 每日排程
            daily_schedule = calculate_schedule(target_spots, days_count)
            
            # 生成資料列 (如果是全省，要展開成6區)
            regions_to_loop = display_regions if cfg["is_national"] else calc_regions
            
            for r_idx, reg in enumerate(regions_to_loop):
                # 取得該區資料
                if cfg["is_national"]:
                    # 如果是全省包，顯示的區域是北中南...但價格來源是?
                    # 按照 Excel 慣例，Package 顯示各區時，Rate 欄位通常顯示 List Price / 720 (定價邏輯)
                    # 或是顯示 Package Cost (僅在全省列顯示?)
                    # 根據需求 1: "全家的Location還是要列出6個區域"
                    # 根據需求 5: "有選到全省的，就會有pakage-cost"
                    
                    # 為了顯示正確，我們需要該區域的 List Price (定價)
                    # 全省包的情況下，各區定價依然存在
                    list_price_reg = db.get(reg, [0,0])[0] # 雖然買全省，但展開列出時，用該區定價算 Rate
                    net_price_reg = db.get(reg, [0,0])[1] # 實收佔比 (隱形)
                else:
                    list_price_reg = db[reg][0]
                    net_price_reg = db[reg][1]
                
                # Rate (Net) 欄位顯示邏輯：
                # 依據提示： = 區域定價 / 720 * 檔次 * 折扣
                # 這裡強制用 720 做為分母 (依據 User 要求 100% 還原 Excel)
                rate_value = (list_price_reg / 720.0) * target_spots * discount
                
                # Package Cost 欄位邏輯：
                # 只有全省包的時候，且只顯示在第一列(或合併)，或是每一列都算?
                # 通常是總合顯示。為了表格整潔，我們算在每一列，但 HTML 生成時可能要處理。
                # 公式：全省區定價 / 720 * 檔次 * 折扣 * 1.1 (if spots < 720)
                package_cost_val = 0
                if cfg["is_national"]:
                    # 注意：這裡是每一行都顯示全省的 Package Cost 還是拆分?
                    # 依照 User: "Pakage-cost...= 全省區定價..."
                    # 這聽起來像是一個總數。我們把它放在一個獨立的屬性，HTML 渲染時處理。
                    # 這裡先算總數
                    nat_list_price = db["全省"][0]
                    pk_multiplier = 1.1 if target_spots < 720 else 1.0
                    package_cost_total = (nat_list_price / 720.0) * target_spots * discount * pk_multiplier
                    # 我們把這個值只附在 "北區" (第一筆) 上，其他區放空值，避免重複加總
                    package_cost_display = package_cost_total if reg == "北區" else 0
                else:
                    package_cost_display = 0

                # 店數文字
                store_key = reg if m_type == "全家廣播" else f"新鮮視_{reg}"
                program_txt = STORE_COUNTS.get(store_key, reg)
                
                final_rows.append({
                    "media": m_type,
                    "region": reg,
                    "seconds": sec,
                    "spots": target_spots,
                    "schedule": daily_schedule,
                    "program": program_txt,
                    "daypart": "07:00-23:00",
                    "rate_net": rate_value, # 這是顯示在 Cue 表上的 Rate (基於定價/720)
                    "package_cost": package_cost_display,
                    "is_national_buy": cfg["is_national"],
                    # 真實成本 (用於計算 Grand Total，不顯示在表內細項)
                    # 若是全省包，成本只算一次(在第一筆)，或者按比例拆?
                    # 簡單作法：全省包的成本算在第一筆(北區)，其他為0，以免加總重複
                    "real_cost": (combined_unit_net * target_spots) if (not cfg["is_national"] or reg == "北區") else 0
                })

        # --- 處理家樂福 ---
        elif m_type == "家樂福":
            # 家樂福固定全省，分量販跟超市兩列
            db = PRICING_DB["家樂福"]
            
            # 取得兩者的 Net Unit Price
            # 量販
            hyp_net = db["量販_全省"]["Net"]
            hyp_std = db["量販_全省"]["Std_Spots"]
            hyp_unit = (hyp_net / hyp_std) * discount
            
            # 超市
            sup_net = db["超市_全省"]["Net"]
            sup_std = db["超市_全省"]["Std_Spots"]
            sup_unit = (sup_net / sup_std) * discount
            
            combined_unit = hyp_unit + sup_unit
            
            # 計算檔次 (量販超市檔次相同? 題目說 "只要有選到就一定會有...兩種")
            # 假設兩者檔次一致 (因預算是一包)
            target_spots = math.floor(sec_budget / combined_unit)
            
            # 偶數調整
            if target_spots % 2 != 0:
                 target_spots -= 1 # 保守一點
            if target_spots <= 0: target_spots = 1
            
            daily_schedule = calculate_schedule(target_spots, days_count)
            
            # 1. 量販 Row
            rate_hyp = (db["量販_全省"]["List"] / 720.0) * target_spots * discount
            final_rows.append({
                "media": "家樂福",
                "region": "全省量販",
                "seconds": sec,
                "spots": target_spots,
                "schedule": daily_schedule,
                "program": "全省",
                "daypart": "09:00-23:00",
                "rate_net": rate_hyp,
                "package_cost": 0,
                "is_national_buy": True,
                "real_cost": hyp_unit * target_spots
            })
            
            # 2. 超市 Row
            rate_sup = (db["超市_全省"]["List"] / 720.0) * target_spots * discount
            final_rows.append({
                "media": "家樂福",
                "region": "全省超市",
                "seconds": sec,
                "spots": target_spots,
                "schedule": daily_schedule,
                "program": "全省",
                "daypart": "00:00-24:00",
                "rate_net": rate_sup,
                "package_cost": 0,
                "is_national_buy": True,
                "real_cost": sup_unit * target_spots
            })

# ==========================================
# 4. 生成 HTML (Output)
# ==========================================

# 計算總金額
total_media_cost = sum(r["real_cost"] for r in final_rows)
production_cost = 10000
vat = (total_media_cost + production_cost) * 0.05
grand_total = total_media_cost + production_cost + vat

# 表格 HTML
date_headers = ""
curr = start_date
for _ in range(days_count):
    date_headers += f"<th class='date-col'>{curr.strftime('%m/%d')}</th>"
    curr += timedelta(days=1)

rows_html = ""
for row in final_rows:
    sch_cells = "".join([f"<td class='sch'>{s}</td>" for s in row["schedule"]])
    
    # Package Cost 顯示
    pkg_cost_str = ""
    if row["package_cost"] > 0:
        pkg_cost_str = f"{int(row['package_cost']):,}"
    
    # 媒體名稱轉換 (全家便利商店通路廣播廣告...)
    media_display = row["media"]
    if media_display == "全家廣播": media_display = "全家便利商店通路廣播廣告"
    if media_display == "新鮮視": media_display = "全家便利商店新鮮視"
    
    # Rate Net 顯示 (整數)
    rate_display = f"{int(row['rate_net']):,}"
    
    rows_html += f"""
    <tr>
        <td style="text-align:left">{media_display}</td>
        <td>{row['region']}</td>
        <td>{row['program']}</td>
        <td>{row['daypart']}</td>
        <td>{row['seconds']}</td>
        {sch_cells}
        <td style="font-weight:bold; background-color:#ffeb3b4d;">{row['spots']}</td>
        <td style="text-align:right;">{rate_display}</td>
        <td style="text-align:right;">{pkg_cost_str}</td>
    </tr>
    """

# 檢查是否需要顯示 Package Cost column
show_pkg_col = any(r["package_cost"] > 0 for r in final_rows)
pkg_header = "<th>Package Cost</th>"
pkg_col_style = "" 
# 為了版面一致，即使沒有值也顯示該欄位，或者動態隱藏。
# 依照需求 5: "有選到...全省的...這個欄位就要出來"。我們動態控制 CSS width 避免太醜
# 這裡簡單處理：總是顯示，若沒值就空著

html = f"""
<style>
    table {{ width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; font-size: 11px; }}
    th, td {{ border: 1px solid #000; padding: 3px; text-align: center; }}
    .head-info {{ background-color: #f2f2f2; text-align: left; border: none; padding: 10px; }}
    .col-head {{ background-color: #cfe2f3; font-weight: bold; }}
    .date-col {{ writing-mode: vertical-rl; transform: rotate(180deg); width: 20px; font-size: 10px; background-color: #cfe2f3; }}
    .sch {{ font-size: 11px; }}
    .total-row {{ background-color: #ffffcc; font-weight: bold; }}
</style>

<div style="background:white; padding:10px; border: 1px solid #ccc;">
    <table>
        <tr>
            <td colspan="5" class="head-info">
                <b>Client:</b> {client_name}<br>
                <b>Product:</b> {", ".join(sorted(list(all_selected_secs)))}<br>
                <b>Period:</b> {start_date.strftime('%Y/%m/%d')} ~ {end_date.strftime('%Y/%m/%d')}<br>
                <b>Medium:</b> {", ".join(list(all_selected_media))}
            </td>
            <td colspan="{days_count + 3}" style="border:none;"></td>
        </tr>
        <tr class="col-head">
            <th>Station</th>
            <th>Location</th>
            <th>Program</th>
            <th>Day-part</th>
            <th>Size</th>
            {date_headers}
            <th>Total Spots</th>
            <th>Rate (Net)</th>
            {pkg_header}
        </tr>
        {rows_html}
        <tr class="total-row">
            <td colspan="5" style="text-align:right">Media Total</td>
            <td colspan="{days_count}"></td>
            <td>{sum(r['spots'] for r in final_rows)}</td>
            <td style="text-align:right">{int(total_media_cost):,}</td>
            <td></td>
        </tr>
        <tr>
            <td colspan="5" style="text-align:right">Production Cost</td>
            <td colspan="{days_count + 2}" style="text-align:right">{production_cost:,}</td>
            <td></td>
        </tr>
        <tr>
            <td colspan="5" style="text-align:right">5% VAT</td>
            <td colspan="{days_count + 2}" style="text-align:right">{int(vat):,}</td>
            <td></td>
        </tr>
        <tr class="total-row" style="border-top: 2px double black;">
            <td colspan="5" style="text-align:right">Grand Total</td>
            <td colspan="{days_count + 2}" style="text-align:right">NT$ {int(grand_total):,}</td>
            <td></td>
        </tr>
    </table>
</div>
"""

st.markdown("---")
st.markdown("### 3. Cue Sheet 預覽")
st.components.v1.html(html, height=600, scrolling=True)

# Debug 資訊 (可選)
with st.expander("查看計算詳情 (Debug)"):
    st.write(f"總預算: {actual_total_budget}")
    st.write(f"實際計算總成本(未稅): {int(total_media_cost)}")
    st.write(f"差異: {int(total_media_cost - actual_total_budget)}")
