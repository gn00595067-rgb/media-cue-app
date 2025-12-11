import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
from jinja2 import Template
from datetime import datetime, date

# 1. 頁面設定
st.set_page_config(page_title="Cue表排程系統", layout="wide")

# 2. HTML/CSS 模板 (保持原樣，這是你喜歡的樣式)
html_template_str = """
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <style>
        body { font-family: "Microsoft JhengHei", "Heiti TC", sans-serif; margin: 0; padding: 10px; color: #333; }
        .header-info { background-color: #f1f3f4; padding: 15px; margin-bottom: 20px; border-left: 6px solid #1a73e8; border-radius: 4px; }
        .header-info p { margin: 5px 0; font-weight: bold; font-size: 14px; }
        .header-info span { font-weight: normal; color: #555; }
        table { width: 100%; border-collapse: collapse; font-size: 13px; white-space: nowrap; }
        th, td { border: 1px solid #c0c0c0; padding: 8px; text-align: center; vertical-align: middle; }
        th { background-color: #3c4043; color: #ffffff; position: sticky; top: 0; }
        .text-left { text-align: left; }
        .text-right { text-align: right; }
        tbody tr:nth-child(even) { background-color: #f8f9fa; }
        tbody tr:hover { background-color: #e8f0fe; }
        .package-cell { background-color: #fff !important; font-weight: bold; color: #d93025; border-bottom: 1px solid #bbb; }
        .total-row { background-color: #e8eaed !important; font-weight: bold; border-top: 2px solid #333; }
    </style>
</head>
<body>
    <div class="header-info">
        <p>客戶名稱：<span>{{ client }}</span></p>
        <p>Product：<span>{{ product }}</span></p>
        <p>Period：<span>{{ period }}</span></p>
        <p>Budget：<span>{{ budget }}</span></p>
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
                <th>Package-cost</th>
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
                {% if row.is_first %}
                    <td class="text-right package-cell" rowspan="{{ row.rowspan }}">
                        {{ "{:,}".format(row.package_cost) }}
                    </td>
                {% endif %}
                {% for i in range(1, 16) %}
                <td>{{ row.get(i, 0) }}</td>
                {% endfor %}
            </tr>
            {% endfor %}
            <tr class="total-row">
                <td colspan="5" class="text-right">Total:</td>
                <td class="text-right">{{ "{:,}".format(total_rate) }}</td>
                <td></td>
                <td colspan="15"></td>
            </tr>
        </tbody>
    </table>
</body>
</html>
"""

# 3. 初始化資料函數 (關鍵修復：確保欄位型態一致)
def get_initial_df():
    # 預設資料
    data = [
        {"PackageGroup": "A", "Station": "全家廣播", "Location": "北區", "Program": "北北基", "Daypart": "全天", "Size": "20秒", "Rate": 416111},
        {"PackageGroup": "A", "Station": "全家廣播", "Location": "桃竹苗", "Program": "桃竹苗", "Daypart": "全天", "Size": "20秒", "Rate": 249667},
        {"PackageGroup": "B", "Station": "全家廣播", "Location": "中區", "Program": "中彰投", "Daypart": "全天", "Size": "20秒", "Rate": 249667}
    ]
    df = pd.DataFrame(data)
    
    # 預先建立 1~15 號的欄位，全部填入預設值 50 (整數)
    # 這一步很重要，避免 data_editor 讀不到欄位而卡住
    for i in range(1, 16):
        df[str(i)] = 50 
        
    return df

# 4. 資料處理邏輯
def process_data(df):
    # 確保數值型態正確
    df['Rate'] = pd.to_numeric(df['Rate'], errors='coerce').fillna(0).astype(int)
    
    # 計算 G20 Package Cost
    if 'PackageGroup' in df.columns:
        group_sums = df.groupby('PackageGroup')['Rate'].sum().to_dict()
    else:
        group_sums = {}

    processed_rows = []
    seen_groups = set()
    
    for index, row in df.iterrows():
        row_dict = row.to_dict()
        group = row_dict.get('PackageGroup', 'Unknown')
        
        if group not in seen_groups:
            count = len(df[df['PackageGroup'] == group])
            row_dict['rowspan'] = count
            row_dict['package_cost'] = group_sums.get(group, 0)
            row_dict['is_first'] = True
            seen_groups.add(group)
        else:
            row_dict['is_first'] = False
            
        # 處理日期欄位
        for i in range(1, 16):
            # 確保抓取的是字串 key
            val = row_dict.get(str(i), 0)
            row_dict[i] = int(val)
            
        processed_rows.append(row_dict)
        
    total_rate = df['Rate'].sum()
    return processed_rows, total_rate

# 5. 主程式
def main():
    st.title("📺 廣播 Cue 表排程系統 (穩定版)")
    
    # --- Sidebar ---
    with st.sidebar:
        st.header("1. 基本設定")
        client_name = st.text_input("客戶名稱", "萬國通路")
        product_name = st.text_input("產品", "20秒形象廣告")
        col1, col2 = st.columns(2)
        start_date = col1.date_input("開始", date(2025, 1, 1))
        end_date = col2.date_input("結束", date(2025, 1, 31))
        budget = st.number_input("預算", 1000000, step=10000)
        
        st.markdown("---")
        st.info("請在右側表格編輯資料，確認無誤後點擊下方按鈕生成報表。")
        
        # 【關鍵修改】加入按鈕，避免即時運算造成卡頓
        generate_btn = st.button("🚀 生成 / 更新報表", type="primary")

    # --- Main Area ---
    st.subheader("📝 編輯排程資料")
    
    # 使用 session_state 防止資料重置
    if 'df_data' not in st.session_state:
        st.session_state.df_data = get_initial_df()

    # 顯示編輯器
    edited_df = st.data_editor(
        st.session_state.df_data,
        num_rows="dynamic",
        use_container_width=True,
        key="editor", # 給予 key 讓 streamlit 追蹤狀態
        column_config={
            "Rate": st.column_config.NumberColumn("Rate (Net)", format="$%d"),
            "PackageGroup": st.column_config.TextColumn("群組 (G20)", help="相同代號會合併計算費用"),
        }
    )

    # --- 只有按下按鈕時才執行耗時的渲染 ---
    if generate_btn:
        with st.spinner("報表生成中..."):
            # 1. 處理資料
            rows, total_rate = process_data(edited_df)
            
            # 2. 渲染 HTML
            period_str = f"{start_date} - {end_date}"
            budget_str = "{:,}".format(budget)
            
            template = Template(html_template_str)
            html_output = template.render(
                client=client_name,
                product=product_name,
                period=period_str,
                budget=budget_str,
                rows=rows,
                total_rate=total_rate
            )
            
            # 3. 顯示結果
            st.success("✅ 報表已更新")
            st.divider()
            components.html(html_output, height=600, scrolling=True)
            
            # 4. 下載按鈕
            st.download_button(
                label="📥 下載 HTML",
                data=html_output,
                file_name="cue_report.html",
                mime="text/html"
            )

if __name__ == "__main__":
    main()
