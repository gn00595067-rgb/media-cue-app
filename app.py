import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
from jinja2 import Template

# 設定頁面寬度為寬版，讓報表更好看
st.set_page_config(layout="wide", page_title="Cue 表預覽")

# ---------------------------------------------------------
# 1. 準備模擬資料 (模擬你截圖中的數據結構)
# ---------------------------------------------------------
raw_data = [
    {
        "station": "全家便利商店通路廣播",
        "location": "北區-北區",
        "program": "北北基 1,649店",
        "daypart": "00:00-24:00",
        "size": "20秒",
        "rate": 416111,
        "package_group": "A", # 群組 A
        "spots": [50, 50, 50, 50, 50, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48]
    },
    {
        "station": "全家便利商店通路廣播",
        "location": "桃竹苗區-桃竹苗",
        "program": "桃竹苗 779店",
        "daypart": "00:00-24:00",
        "size": "20秒",
        "rate": 249667,
        "package_group": "A", # 群組 A (與上一筆同組，費用會加總)
        "spots": [50, 50, 50, 50, 50, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48]
    },
    {
        "station": "全家便利商店通路廣播",
        "location": "中區-中區",
        "program": "中彰投 839店",
        "daypart": "00:00-24:00",
        "size": "20秒",
        "rate": 249667,
        "package_group": "B", # 群組 B (新的群組，費用分開算)
        "spots": [50, 50, 50, 50, 50, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48]
    },
    {
        "station": "全家便利商店通路廣播",
        "location": "雲嘉南區",
        "program": "雲嘉南 900店",
        "daypart": "00:00-24:00",
        "size": "20秒",
        "rate": 200000, 
        "package_group": "B", # 群組 B
        "spots": [50, 50, 50, 50, 50, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48]
    }
]

# ---------------------------------------------------------
# 2. 資料處理邏輯：計算 Package Cost (G20 的秘密)
# ---------------------------------------------------------
df = pd.DataFrame(raw_data)

# 計算每個群組的總金額
group_sums = df.groupby('package_group')['rate'].sum().to_dict()

processed_rows = []
seen_groups = set()

for index, row in df.iterrows():
    group = row['package_group']
    rate = row['rate']
    row_data = row.to_dict()
    
    # 判斷是否為該群組的第一筆資料 (為了做 HTML rowspan)
    if group not in seen_groups:
        group_count = len(df[df['package_group'] == group])
        row_data['rowspan'] = group_count
        row_data['package_cost'] = group_sums[group] # 填入加總後的金額
        row_data['is_first'] = True
        seen_groups.add(group)
    else:
        row_data['is_first'] = False
    
    processed_rows.append(row_data)

# 計算總 Total (這只是為了讓畫面更完整)
total_amount = df['rate'].sum()

# ---------------------------------------------------------
# 3. 定義 HTML 模板 (包含 CSS 格線與樣式)
# ---------------------------------------------------------
html_template = """
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <style>
        /* 全域字體設定 */
        body { 
            font-family: "Microsoft JhengHei", "Heiti TC", sans-serif; 
            margin: 0; 
            padding: 10px;
            color: #333;
        }
        
        /* 表頭資訊區塊 */
        .header-info { 
            background-color: #f8f9fa; 
            padding: 15px; 
            margin-bottom: 20px; 
            border-left: 5px solid #2c3e50; 
            border-radius: 4px;
        }
        .header-info p { margin: 5px 0; font-weight: bold; font-size: 14px; }

        /* 表格核心設定 */
        table { 
            width: 100%; 
            border-collapse: collapse; /* 重要：讓邊框合併 */
            font-size: 13px; 
            background-color: #fff;
        }

        /* 儲存格設定 */
        th, td { 
            border: 1px solid #bbb; /* 設定格線顏色 */
            padding: 10px 8px; 
            text-align: center; 
            vertical-align: middle; 
        }

        /* 表頭特別設定 */
        th { 
            background-color: #34495e; /* 深藍灰色背景 */
            color: white; 
            font-weight: normal; 
            white-space: nowrap; 
        }

        /* 對齊輔助類別 */
        .text-left { text-align: left; }
        .text-right { text-align: right; }

        /* 斑馬紋 (隔行變色) */
        tbody tr:nth-child(even) { background-color: #f2f2f2; }
        
        /* 滑鼠經過變色 */
        tbody tr:hover { background-color: #e6f7ff; }

        /* Package Cost 欄位特別樣式 */
        .package-cell { 
            background-color: #fff !important; 
            font-weight: bold; 
            color: #c0392b; /* 紅色數字 */
            border-bottom: 1px solid #bbb;
        }
        
        /* 總計列樣式 */
        .total-row {
            background-color: #e2e6ea !important;
            font-weight: bold;
            border-top: 2px solid #333;
        }
    </style>
</head>
<body>

    <div class="header-info">
        <p>客戶名稱：萬國通路</p>
        <p>Product：20秒、5秒</p>
        <p>Period：2025. 01. 01 - 2025. 01. 31</p>
        <p>Medium：家樂福、全家廣播、新鮮視</p>
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
                <th>Package-cost (Net)</th>
                {% for i in range(1, 16) %}
                <th>{{ i }}<br>三</th>
                {% endfor %}
            </tr>
        </thead>
        <tbody>
            {% for row in rows %}
            <tr>
                <td class="text-left">{{ row.station }}</td>
                <td class="text-left">{{ row.location }}</td>
                <td class="text-left">{{ row.program }}</td>
                <td>{{ row.daypart }}</td>
                <td>{{ row.size }}</td>
                <td class="text-right">{{ "{:,}".format(row.rate) }}</td>
                
                {# 這裡處理合併儲存格邏輯 #}
                {% if row.is_first %}
                    <td class="text-right package-cell" rowspan="{{ row.rowspan }}">
                        {{ "{:,}".format(row.package_cost) }}
                    </td>
                {% endif %}

                {% for spot in row.spots %}
                <td>{{ spot }}</td>
                {% endfor %}
            </tr>
            {% endfor %}

            <tr class="total-row">
                <td colspan="5" class="text-right">Total:</td>
                <td class="text-right">{{ "{:,}".format(total_amt) }}</td>
                <td></td> <td colspan="15"></td>
            </tr>
        </tbody>
    </table>

</body>
</html>
"""

# ---------------------------------------------------------
# 4. 渲染 HTML 並顯示
# ---------------------------------------------------------

st.title("Cue 表排程預覽系統")
st.info("已套用樣式：格線、斑馬紋背景、自動計算 Package Cost")

# 使用 Jinja2 渲染 HTML
template = Template(html_template)
html_output = template.render(rows=processed_rows, total_amt=total_amount)

# 【關鍵】使用 Streamlit components 顯示 HTML
# height 設定為 600 或更高，scrolling=True 讓表格寬度超出時可以捲動
components.html(html_output, height=600, scrolling=True)

# 下載按鈕
st.download_button(
    label="📥 下載完整 HTML 報表",
    data=html_output,
    file_name="cue_schedule_report.html",
    mime="text/html"
)
