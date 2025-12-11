import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
from jinja2 import Template
from datetime import datetime, date

# ==========================================
# 1. 頁面基本設定
# ==========================================
st.set_page_config(page_title="Cue表排程系統", layout="wide")

# CSS 樣式模板 (包含你指定的格線、斑馬紋、靠右對齊)
html_template_str = """
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <style>
        body { font-family: "Microsoft JhengHei", "Heiti TC", sans-serif; margin: 0; padding: 10px; color: #333; }
        
        /* 表頭資訊區塊 */
        .header-info { 
            background-color: #f1f3f4; 
            padding: 15px; 
            margin-bottom: 20px; 
            border-left: 6px solid #1a73e8; 
            border-radius: 4px;
        }
        .header-info p { margin: 5px 0; font-weight: bold; font-size: 14px; }
        .header-info span { font-weight: normal; color: #555; }

        /* 表格核心設定 */
        table { 
            width: 100%; 
            border-collapse: collapse; /* 格線合併 */
            font-size: 13px; 
            white-space: nowrap; 
        }

        th, td { 
            border: 1px solid #c0c0c0; /* 清楚的格線 */
            padding: 8px; 
            text-align: center; 
            vertical-align: middle; 
        }

        th { 
            background-color: #3c4043; 
            color: #ffffff; 
            position: sticky; 
            top: 0; 
        }

        .text-left { text-align: left; }
        .text-right { text-align: right; }

        /* 斑馬紋 */
        tbody tr:nth-child(even) { background-color: #f8f9fa; }
        tbody tr:hover { background-color: #e8f0fe; }

        /* Package Cost 重點欄位 */
        .package-cell { 
            background-color: #fff !important; 
            font-weight: bold; 
            color: #d93025; 
            border-bottom: 1px solid #bbb; 
        }
        
        /* 總計列 */
        .total-row { 
            background-color: #e8eaed !important; 
            font-weight: bold; 
            border-top: 2px solid #333; 
        }
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
                
                {# 處理 G20 合併儲存格邏輯 #}
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

# ==========================================
# 2. 核心邏輯函數
# ==========================================

def get_default_data():
    """提供預設的編輯資料"""
    return pd.DataFrame([
        {
            "PackageGroup": "A", "Station": "全家廣播", "Location": "北區", 
            "Program": "北北基 1649店", "Daypart": "00:00-24:00", "Size": "20秒", "Rate": 416111
        },
        {
            "PackageGroup": "A", "Station": "全家廣播", "Location": "桃竹苗", 
            "Program": "桃竹苗 779店", "Daypart": "00:00-24:00", "Size": "20秒", "Rate": 249667
        },
        {
            "PackageGroup": "B", "Station": "全家廣播", "Location": "中區", 
            "Program": "中彰投 839店", "Daypart": "00:00-24:00", "Size": "20秒", "Rate": 249667
        }
    ])

def process_data_for_report(df):
    """將 DataFrame 轉換為報表需要的格式 (含 Package Cost 計算)"""
    
    # 確保 Rate 是數字
    df['Rate'] = pd.to_numeric(df['Rate'], errors='coerce').fillna(0).astype(int)
    
    # 計算 Group Sum (G20 邏輯)
    if 'PackageGroup' in df.columns:
        group_sums = df.groupby('PackageGroup')['Rate'].sum().to_dict()
    else:
        group_sums = {}

    processed_rows = []
    seen_groups = set()
    
    # 建立 1~15 的日期欄位 (如果資料表沒有，就補0)
    for i in range(1, 16):
        col_name = str(i)
        if col_name not in df.columns:
            df[col_name] = 50 # 預設次數，方便演示
            
    for index, row in df.iterrows():
        row_dict = row.to_dict()
        group = row_dict.get('PackageGroup', 'Unknown')
        
        # Rowspan 邏輯
        if group not in seen_groups:
            count = len(df[df['PackageGroup'] == group])
            row_dict['rowspan'] = count
            row_dict['package_cost'] = group_sums.get(group, 0)
            row_dict['is_first'] = True
            seen_groups.add(group)
        else:
            row_dict['is_first'] = False
            
        # 處理日期欄位 key (轉成 int 1-15 方便模板讀取)
        for i in range(1, 16):
            # 嘗試讀取 string key '1' 或 int key 1
            val = row_dict.get(str(i)) or row_dict.get(i)
            row_dict[i] = int(val) if val else 0
            
        processed_rows.append(row_dict)
        
    total_rate = df['Rate'].sum()
    return processed_rows, total_rate

# ==========================================
# 3. 主程式 UI
# ==========================================

def main():
    st.title("📺 廣播 Cue 表排程系統")
    
    # --- 左側 Sidebar：業務輸入區 ---
    with st.sidebar:
        st.header("1. 專案設定")
        client_name = st.text_input("客戶名稱", value="萬國通路")
        product_name = st.text_input("產品名稱", value="20秒、5秒形象廣告")
        
        st.header("2. 走期選擇")
        # 日期區間選擇器
        col1, col2 = st.columns(2)
        start_date = col1.date_input("開始", value=date(2025, 1, 1))
        end_date = col2.date_input("結束", value=date(2025, 1, 31))
        period_str = f"{start_date.strftime('%Y.%m.%d')} - {end_date.strftime('%Y.%m.%d')}"
        
        st.header("3. 預算設定")
        budget_input = st.number_input("總預算 (Budget)", value=1000000, step=10000)
        budget_str = "{:,}".format(budget_input)
        
        st.markdown("---")
        st.info("💡 提示：在右側表格直接修改數據，PackageGroup 相同的項目，金額會自動加總。")

    # --- 右側主畫面：資料編輯與預覽 ---
    
    st.subheader("📝 排程資料編輯")
    
    # 初始化 Session State 以保存編輯後的資料
    if 'editor_data' not in st.session_state:
        st.session_state.editor_data = get_default_data()

    # 顯示可編輯的 DataFrame (Data Editor)
    # 這裡讓業務可以直接打字，不用上傳 Excel
    edited_df = st.data_editor(
        st.session_state.editor_data,
        num_rows="dynamic", # 允許新增刪除列
        column_config={
            "Rate": st.column_config.NumberColumn("Rate (Net)", format="$%d"),
            "PackageGroup": st.column_config.TextColumn("群組代碼 (G20邏輯)", help="代碼相同的列，費用會合併計算"),
        },
        hide_index=True,
        use_container_width=True
    )
    
    # --- 生成報表邏輯 ---
    st.divider()
    st.subheader("📊 Cue 表預覽")
    
    if not edited_df.empty:
        # 呼叫處理函數
        rows, total_rate = process_data_for_report(edited_df)
        
        # 渲染 HTML
        template = Template(html_template_str)
        html_output = template.render(
            client=client_name,
            product=product_name,
            period=period_str,
            budget=budget_str,
            rows=rows,
            total_rate=total_rate
        )
        
        # 顯示 HTML
        components.html(html_output, height=600, scrolling=True)
        
        # 下載按鈕
        st.download_button(
            label="📥 下載 HTML 報表",
            data=html_output,
            file_name=f"cue_schedule_{client_name}.html",
            mime="text/html"
        )
    else:
        st.warning("請在上方表格輸入資料。")

if __name__ == "__main__":
    main()
