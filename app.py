import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
from jinja2 import Template
from datetime import date

# 1. 基礎設定
st.set_page_config(page_title="Cue表排程系統 (穩定版)", layout="wide")

# 2. 媒體資料庫 (這是當時你最滿意的功能：不用打字，用選的)
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

# 3. 最原始的 HTML 模板 (沒有格線、沒有 G20 合併，最單純的 HTML)
html_template_str = """
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <style>
        body { font-family: "Microsoft JhengHei", sans-serif; padding: 20px; }
        .header { background-color: #f0f0f0; padding: 15px; margin-bottom: 20px; border-radius: 5px; }
        table { width: 100%; border-collapse: collapse; font-size: 13px; }
        th { background-color: #333; color: white; padding: 8px; text-align: center; }
        td { border: 1px solid #ddd; padding: 8px; text-align: center; } /* 基本邊框 */
        .text-right { text-align: right; }
        .text-left { text-align: left; }
    </style>
</head>
<body>
    <div class="header">
        <p><strong>客戶：</strong>{{ client }}</p>
        <p><strong>產品：</strong>{{ product }}</p>
        <p><strong>走期：</strong>{{ period }}</p>
        <p><strong>預算：</strong>{{ budget }}</p>
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
                <th>{{ i }}</th>
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
            <tr style="background-color: #eee; font-weight: bold;">
                <td colspan="5" class="text-right">Total:</td>
                <td class="text-right">{{ "{:,}".format(total_rate) }}</td>
                <td colspan="15"></td>
            </tr>
        </tbody>
    </table>
</body>
</html>
"""

def main():
    st.title("📺 廣播 Cue 表排程系統 (基礎版)")

    # --- Session State 初始化 (防止資料重置) ---
    if "schedule_data" not in st.session_state:
        # 定義欄位結構
        st.session_state.schedule_data = pd.DataFrame(
            columns=["Station", "Location", "Program", "Daypart", "Size", "Rate"] + [str(i) for i in range(1, 16)]
        )

    # --- 左側選單 ---
    with st.sidebar:
        st.header("1. 專案設定")
        client_name = st.text_input("客戶名稱", "萬國通路")
        product_name = st.text_input("產品", "形象廣告")
        col1, col2 = st.columns(2)
        start_date = col1.date_input("開始", date(2025, 1, 1))
        end_date = col2.date_input("結束", date(2025, 1, 31))
        budget = st.number_input("預算", 1000000, step=10000)

    # --- Step 1: 選單區 (你喜歡的那個功能) ---
    st.subheader("Step 1: 新增通路")
    
    # 這裡用 container 包起來排版
    with st.container():
        c1, c2, c3 = st.columns([1, 1, 1])
        
        with c1:
            station = st.selectbox("選擇通路", list(MEDIA_DATABASE.keys()))
        
        with c2:
            # 連動選單
            loc_opts = [x["Location"] for x in MEDIA_DATABASE[station]]
            location = st.selectbox("選擇區域", loc_opts)
            # 抓取詳細資料
            selected_info = next(x for x in MEDIA_DATABASE[station] if x["Location"] == location)
            
        with c3:
            default_spots = st.number_input("每日次數", value=50)
            
        if st.button("➕ 加入清單"):
            new_row = {
                "Station": station,
                "Location": location,
                "Program": selected_info["Program"],
                "Daypart": "00:00-24:00",
                "Size": "20秒",
                "Rate": selected_info["Rate"]
            }
            # 填入 1~15 次數
            for i in range(1, 16):
                new_row[str(i)] = default_spots
            
            # 加入資料表
            st.session_state.schedule_data = pd.concat(
                [st.session_state.schedule_data, pd.DataFrame([new_row])], 
                ignore_index=True
            )
            st.success(f"已加入 {location}")

    # --- Step 2: 表格顯示與編輯 ---
    st.divider()
    st.subheader("Step 2: 確認排程")
    
    if not st.session_state.schedule_data.empty:
        # 顯示可編輯表格
        edited_df = st.data_editor(
            st.session_state.schedule_data,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Rate": st.column_config.NumberColumn("Rate", format="$%d")
            }
        )
        
        # --- Step 3: 生成報表 ---
        st.divider()
        if st.button("🚀 生成報表"):
            # 準備資料
            edited_df['Rate'] = pd.to_numeric(edited_df['Rate'], errors='coerce').fillna(0).astype(int)
            total = edited_df['Rate'].sum()
            
            rows_list = []
            for _, row in edited_df.iterrows():
                r = row.to_dict()
                for i in range(1, 16):
                    # 確保是整數
                    val = r.get(str(i))
                    r[i] = int(val) if val else 0
                rows_list.append(r)
            
            period_str = f"{start_date} - {end_date}"
            
            # 渲染 HTML
            template = Template(html_template_str)
            html_out = template.render(
                client=client_name,
                product=product_name,
                period=period_str,
                budget="{:,}".format(budget),
                rows=rows_list,
                total_rate=total
            )
            
            st.subheader("📊 報表預覽")
            components.html(html_out, height=500, scrolling=True)
            
            st.download_button("📥 下載 HTML", html_out, "cue_schedule.html")
    else:
        st.info("目前清單是空的，請上方加入通路。")

if __name__ == "__main__":
    main()
