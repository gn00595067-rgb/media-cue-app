import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
from jinja2 import Template
from datetime import date

# ==========================================
# 1. 頁面設定
# ==========================================
st.set_page_config(page_title="Cue表排程系統 (業務版)", layout="wide")

# ==========================================
# 2. 內建媒體資料庫 (回復：這是你 G20 出現前就在用的選單邏輯)
# ==========================================
MEDIA_DATABASE = {
    "全家便利商店": [
        {"Location": "北區-北北基", "Program": "北北基 1,649店", "Rate": 416111},
        {"Location": "北區-桃竹苗", "Program": "桃竹苗 779店", "Rate": 249667},
        {"Location": "中區", "Program": "中彰投 839店", "Rate": 249667},
        {"Location": "南區", "Program": "雲嘉南 900店", "Rate": 200000},
        {"Location": "南區-高屏", "Program": "高屏 720店", "Rate": 200000},
    ],
    "家樂福 (量販)": [
        {"Location": "全台", "Program": "全台 67店", "Rate": 350000},
        {"Location": "北區", "Program": "北區 25店", "Rate": 150000},
    ],
    "家樂福 (超市)": [
        {"Location": "全台", "Program": "全台 245店", "Rate": 180000},
    ],
    "新鮮視": [
        {"Location": "全台", "Program": "全台聯播", "Rate": 50000},
    ]
}

# ==========================================
# 3. HTML/CSS 美化模板
# (回復：這裡拿掉了 G20 的 rowspan 合併邏輯，每一列都獨立顯示)
# (保留：你要求的格線與斑馬紋)
# ==========================================
html_template_str = """
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <style>
        body { font-family: "Microsoft JhengHei", sans-serif; margin: 0; padding: 20px; color: #333; }
        
        .header-info { 
            background-color: #f8f9fa; 
            padding: 15px; 
            margin-bottom: 20px; 
            border-left: 5px solid #007bff; 
            border-radius: 4px;
        }
        .header-info p { margin: 5px 0; font-weight: bold; font-size: 14px; }
        .header-info span { font-weight: normal; color: #555; }

        /* 表格美化 */
        table { 
            width: 100%; 
            border-collapse: collapse; 
            font-size: 13px; 
            white-space: nowrap; 
        }

        th, td { 
            border: 1px solid #bbb; /* 清楚格線 */
            padding: 8px; 
            text-align: center; 
            vertical-align: middle; 
        }

        th { 
            background-color: #343a40; /* 深色表頭 */
            color: white; 
            position: sticky; top: 0;
        }

        .text-left { text-align: left; }
        .text-right { text-align: right; }

        /* 斑馬紋 */
        tbody tr:nth-child(even) { background-color: #f2f2f2; } 
        tbody tr:hover { background-color: #e6f7ff; }

        .total-row { 
            background-color: #e9ecef !important; 
            font-weight: bold; 
            border-top: 2px solid #333; 
        }
    </style>
</head>
<body>
    <div class="header-info">
        <p>客戶名稱：<span>{{ client }}</span></p>
        <p>產品：<span>{{ product }}</span></p>
        <p>走期：<span>{{ period }}</span></p>
        <p>預算：<span>{{ budget }}</span></p>
    </div>

    <table>
        <thead>
            <tr>
                <th>Station</th>
                <th>Location</th>
                <th>Program</th>
                <th>Day-part</th>
                <th>Size</th>
                <th>Rate (Net)</th>
                {% for i in range(1, 16) %}
                <th>{{ i }}<br>日</th>
                {% endfor %}
            </tr>
        </thead>
        <tbody>
            {% for row in rows %}
            <tr>
                <td class="text-left">{{ row.Station }}</td>
                <td class="text-left">{{ row.Location }}</td>
                <td class="text-left">{{ row.Program }}</td>
                <td>{{ row.Daypart }}</td>
                <td>{{ row.Size }}</td>
                <td class="text-right">{{ "{:,}".format(row.Rate) }}</td>
                {% for i in range(1, 16) %}
                <td>{{ row.get(i, 0) }}</td>
                {% endfor %}
            </tr>
            {% endfor %}
            <tr class="total-row">
                <td colspan="5" class="text-right">Total:</td>
                <td class="text-right">{{ "{:,}".format(total_rate) }}</td>
                <td colspan="15"></td>
            </tr>
        </tbody>
    </table>
</body>
</html>
"""

# ==========================================
# 4. 主程式
# ==========================================
def main():
    st.title("📺 廣播 Cue 表排程系統")

    # --- Sidebar: 業務輸入區 ---
    with st.sidebar:
        st.header("1. 專案設定")
        client_name = st.text_input("客戶名稱", value="萬國通路")
        product_name = st.text_input("產品名稱", value="20秒形象廣告")
        
        col1, col2 = st.columns(2)
        start_date = col1.date_input("開始", value=date(2025, 1, 1))
        end_date = col2.date_input("結束", value=date(2025, 1, 31))
        
        budget_input = st.number_input("總預算", value=1000000, step=10000)
        st.divider()
        st.info("請在右側加入通路並生成報表。")

    # --- Session State 初始化 ---
    if "schedule_df" not in st.session_state:
        # 定義欄位 (包含 1~15 號)
        cols = ["Station", "Location", "Program", "Daypart", "Size", "Rate"] + [str(i) for i in range(1, 16)]
        st.session_state.schedule_df = pd.DataFrame(columns=cols)

    # ==========================================
    # Step 1: 選單加入通路 (回復重點)
    # ==========================================
    st.subheader("Step 1: 新增通路")
    with st.container():
        c1, c2, c3 = st.columns([1, 1, 2])
        
        # 1. 通路選單
        with c1:
            station_opt = st.selectbox("選擇通路", list(MEDIA_DATABASE.keys()))
        
        # 2. 區域選單 (連動)
        with c2:
            loc_options = MEDIA_DATABASE[station_opt]
            loc_names = [x["Location"] for x in loc_options]
            loc_opt = st.selectbox("選擇區域", loc_names)
            
            # 抓取詳細資料
            selected_data = next(x for x in loc_options if x["Location"] == loc_opt)
            
        # 3. 加入按鈕
        with c3:
            default_spots = st.number_input("預設每日次數", value=50, step=1)
            st.write(f"牌價: **${selected_data['Rate']:,}**")
            
        if st.button("➕ 加入至列表"):
            new_row = {
                "Station": station_opt,
                "Location": selected_data["Location"],
                "Program": selected_data["Program"],
                "Daypart": "00:00-24:00",
                "Size": "20秒",
                "Rate": selected_data["Rate"]
            }
            # 填入次數
            for i in range(1, 16):
                new_row[str(i)] = default_spots
            
            # 加到 DataFrame
            st.session_state.schedule_df = pd.concat(
                [st.session_state.schedule_df, pd.DataFrame([new_row])], 
                ignore_index=True
            )
            st.success("已加入！")

    # ==========================================
    # Step 2: 編輯與確認 (回復：單純列表，無 G20 群組欄位)
    # ==========================================
    st.divider()
    st.subheader("Step 2: 編輯排程")
    
    if not st.session_state.schedule_df.empty:
        # 顯示編輯器
        edited_df = st.data_editor(
            st.session_state.schedule_df,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Rate": st.column_config.NumberColumn("Rate", format="$%d"),
            }
        )
        
        # ==========================================
        # Step 3: 生成報表
        # ==========================================
        st.divider()
        if st.button("🚀 生成報表預覽", type="primary"):
            st.subheader("📊 報表預覽")
            
            # 處理資料
            edited_df['Rate'] = pd.to_numeric(edited_df['Rate'], errors='coerce').fillna(0).astype(int)
            total_rate = edited_df['Rate'].sum()
            
            # 轉成 List[Dict] 供 Jinja2 使用
            rows_data = []
            for _, row in edited_df.iterrows():
                r_dict = row.to_dict()
                for i in range(1, 16):
                    val = r_dict.get(str(i))
                    r_dict[i] = int(val) if val else 0
                rows_data.append(r_dict)
            
            period_str = f"{start_date} - {end_date}"
            
            # 渲染 (無 G20 邏輯)
            template = Template(html_template_str)
            html_output = template.render(
                client=client_name,
                product=product_name,
                period=period_str,
                budget="{:,}".format(budget_input),
                rows=rows_data,
                total_rate=total_rate
            )
            
            # 顯示
            components.html(html_output, height=600, scrolling=True)
            
            # 下載
            st.download_button("📥 下載 HTML", html_output, file_name="cue_schedule.html")
            
    else:
        st.info("目前列表為空，請從上方加入通路。")

if __name__ == "__main__":
    main()
