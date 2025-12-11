import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
from jinja2 import Template
from datetime import date

# ==========================================
# 1. 頁面設定
# ==========================================
st.set_page_config(page_title="Cue表排程系統", layout="wide")

# ==========================================
# 2. HTML/CSS 樣式 (保留剛才做好的格線美化)
# ==========================================
html_template_str = """
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <style>
        body { font-family: "Microsoft JhengHei", sans-serif; margin: 0; padding: 20px; color: #333; }
        
        /* 表頭資訊 */
        .header-info { 
            background-color: #f8f9fa; 
            padding: 15px; 
            margin-bottom: 20px; 
            border-left: 5px solid #007bff; 
            border-radius: 4px;
        }
        .header-info p { margin: 5px 0; font-weight: bold; font-size: 14px; }
        .header-info span { font-weight: normal; color: #555; }

        /* 表格美化核心 (格線+斑馬紋) */
        table { 
            width: 100%; 
            border-collapse: collapse; /* 格線合併 */
            font-size: 13px; 
            white-space: nowrap; 
        }

        th, td { 
            border: 1px solid #bbb; /* 加上格線 */
            padding: 8px; 
            text-align: center; 
            vertical-align: middle; 
        }

        th { 
            background-color: #343a40; /* 深色表頭 */
            color: white; 
        }

        /* 文字對齊 */
        .text-left { text-align: left; }
        .text-right { text-align: right; }

        /* 斑馬紋 */
        tbody tr:nth-child(even) { background-color: #f2f2f2; }
        tbody tr:hover { background-color: #e6f7ff; }

        /* 總計列 */
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
        <p>總預算：<span>{{ budget }}</span></p>
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
# 3. 初始化預設資料 (讓業務不用從零開始打)
# ==========================================
def get_initial_data():
    df = pd.DataFrame([
        {
            "Station": "全家廣播", "Location": "北區", "Program": "北北基 1649店", 
            "Daypart": "00:00-24:00", "Size": "20秒", "Rate": 416111
        },
        {
            "Station": "全家廣播", "Location": "桃竹苗", "Program": "桃竹苗 779店", 
            "Daypart": "00:00-24:00", "Size": "20秒", "Rate": 249667
        }
    ])
    # 預設每天播 50 次
    for i in range(1, 16):
        df[str(i)] = 50
    return df

# ==========================================
# 4. 主程式邏輯
# ==========================================
def main():
    st.title("📺 廣播 Cue 表排程系統")

    # --- 左側 Sidebar：業務輸入區 ---
    with st.sidebar:
        st.header("1. 專案基本資料")
        client_name = st.text_input("客戶名稱", value="萬國通路")
        product_name = st.text_input("產品名稱", value="20秒形象廣告")
        
        st.header("2. 走期與預算")
        col1, col2 = st.columns(2)
        start_date = col1.date_input("開始日期", value=date(2025, 1, 1))
        end_date = col2.date_input("結束日期", value=date(2025, 1, 31))
        
        budget_input = st.number_input("總預算 (Budget)", value=1000000, step=10000)
        
        st.markdown("---")
        # 這裡放生成按鈕，避免業務還沒打完字網頁就一直閃
        run_btn = st.button("🚀 生成報表", type="primary")

    # --- 中間：資料編輯區 (業務操作的核心) ---
    st.subheader("📝 編輯排程明細")
    st.info("請在下方表格直接新增、修改電台與播放次數：")

    # 使用 session state 記住業務輸入的資料，才不會不見
    if "editor_data" not in st.session_state:
        st.session_state.editor_data = get_initial_data()

    # 顯示可編輯表格 (Data Editor)
    edited_df = st.data_editor(
        st.session_state.editor_data,
        num_rows="dynamic", # 允許業務按 + 新增列，按垃圾桶刪除列
        use_container_width=True,
        column_config={
            "Rate": st.column_config.NumberColumn("Rate (Net)", format="$%d"),
            "Station": st.column_config.TextColumn("Station", width="medium"),
            "Program": st.column_config.TextColumn("Program", width="medium"),
        }
    )

    # --- 下方：報表預覽區 ---
    if run_btn:
        st.divider()
        st.subheader("📊 報表預覽")
        
        # 1. 資料整理
        # 確保 Rate 是數字
        edited_df['Rate'] = pd.to_numeric(edited_df['Rate'], errors='coerce').fillna(0).astype(int)
        
        # 轉換成列表供 HTML 使用
        rows_data = []
        for _, row in edited_df.iterrows():
            r_dict = row.to_dict()
            # 處理 1~15 日期的數字 (確保是整數)
            for i in range(1, 16):
                val = r_dict.get(str(i))
                r_dict[i] = int(val) if val else 0
            rows_data.append(r_dict)
            
        total_rate = edited_df['Rate'].sum()
        period_str = f"{start_date.strftime('%Y.%m.%d')} - {end_date.strftime('%Y.%m.%d')}"

        # 2. 渲染 HTML
        template = Template(html_template_str)
        html_output = template.render(
            client=client_name,
            product=product_name,
            period=period_str,
            budget="{:,}".format(budget_input),
            rows=rows_data,
            total_rate=total_rate
        )

        # 3. 顯示結果 (加上 scrolling 確保寬度足夠)
        components.html(html_output, height=600, scrolling=True)

        # 4. 下載按鈕
        st.download_button(
            label="📥 下載 HTML 報表",
            data=html_output,
            file_name=f"Cue表_{client_name}.html",
            mime="text/html"
        )

if __name__ == "__main__":
    main()
