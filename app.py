import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
from jinja2 import Template
from datetime import datetime, timedelta

# 設定網頁配置為寬版，方便看報表
st.set_page_config(page_title="Cue表自動生成系統", layout="wide")

def main():
    st.title("📺 廣播 Cue 表排程產生器")
    st.markdown("---")

    # ==========================================
    # 1. 模擬資料輸入 (實際應用時這裡可以是 pd.read_excel)
    # ==========================================
    # 這裡我們模擬截圖中的資料結構
    # 'package_group': 用來控制哪些列要算在一起 (例如北區+桃竹苗是同一組)
    raw_data = [
        {
            "station": "全家便利商店通路廣播",
            "location": "北區-北區",
            "program": "北北基 1,649店",
            "daypart": "00:00-24:00",
            "size": "20秒",
            "rate": 416111,
            "package_group": "A", # 群組 A
            "spots": [50] * 15 # 模擬 1~15 號每天播 50 次
        },
        {
            "station": "全家便利商店通路廣播",
            "location": "桃竹苗區-桃竹苗",
            "program": "桃竹苗 779店",
            "daypart": "00:00-24:00",
            "size": "20秒",
            "rate": 249667,
            "package_group": "A", # 群組 A (費用會跟上面加在一起)
            "spots": [50] * 15
        },
        {
            "station": "全家便利商店通路廣播",
            "location": "中區-中區",
            "program": "中彰投 839店",
            "daypart": "00:00-24:00",
            "size": "20秒",
            "rate": 249667,
            "package_group": "B", # 群組 B (新的一組)
            "spots": [50] * 5 + [48] * 10 # 模擬有些天數次數不同
        },
        {
            "station": "全家便利商店通路廣播",
            "location": "雲嘉南區",
            "program": "雲嘉南 900店",
            "daypart": "00:00-24:00",
            "size": "20秒",
            "rate": 200000,
            "package_group": "B", # 群組 B
            "spots": [48] * 15
        }
    ]

    # ==========================================
    # 2. Python 資料處理核心邏輯
    # ==========================================
    df = pd.DataFrame(raw_data)

    # [關鍵步驟] 計算 Package Cost
    # 這是算出 G20 (665,778) 數字的地方
    group_sums = df.groupby('package_group')['rate'].sum().to_dict()

    # 準備渲染用的資料列表
    processed_rows = []
    seen_groups = set()

    for index, row in df.iterrows():
        group = row['package_group']
        row_dict = row.to_dict()
        
        # 處理合併儲存格邏輯 (Rowspan)
        if group not in seen_groups:
            # 如果是該群組的第一筆，設定 rowspan 和總金額
            count = len(df[df['package_group'] == group])
            row_dict['rowspan'] = count
            row_dict['package_cost'] = group_sums[group]
            row_dict['is_first'] = True
            seen_groups.add(group)
        else:
            # 如果不是第一筆，就不顯示 Package Cost
            row_dict['is_first'] = False
        
        processed_rows.append(row_dict)

    # 計算整張表的總 Total
    total_rate = df['rate'].sum()

    # ==========================================
    # 3. HTML/CSS 模板設計 (包含格線與樣式)
    # ==========================================
    html_template = """
    <!DOCTYPE html>
    <html lang="zh-Hant">
    <head>
        <meta charset="UTF-8">
        <style>
            /* 基礎字體設定 */
            body { 
                font-family: "Microsoft JhengHei", "Heiti TC", sans-serif; 
                margin: 0; padding: 10px; color: #333; 
            }
            
            /* 表頭資訊區塊 */
            .header-info {
                background-color: #f1f3f4;
                padding: 15px;
                margin-bottom: 20px;
                border-left: 6px solid #1a73e8;
                border-radius: 4px;
            }
            .header-info p { margin: 5px 0; font-weight: bold; font-size: 14px; }

            /* 表格主體設定 */
            table {
                width: 100%;
                border-collapse: collapse; /* 重要：讓格線合併，不會有雙線 */
                font-size: 13px;
                white-space: nowrap; /* 避免文字自動換行導致版面亂掉 */
            }

            /* 欄位 (Cell) 設定 */
            th, td {
                border: 1px solid #c0c0c0; /* 設定格線顏色 (灰色) */
                padding: 10px 8px;
                text-align: center;
                vertical-align: middle;
            }

            /* 表頭 (Header) 設定 */
            th {
                background-color: #3c4043; /* 深灰底 */
                color: #ffffff;            /* 白字 */
                position: sticky;          /* 固定表頭 */
                top: 0;
                z-index: 2;
            }

            /* 斑馬紋 (Zebra Striping) - 偶數行變色 */
            tbody tr:nth-child(even) {
                background-color: #f8f9fa; 
            }
            
            /* 滑鼠滑過變色 */
            tbody tr:hover {
                background-color: #e8f0fe;
            }

            /* 輔助樣式 */
            .text-left { text-align: left; }
            .text-right { text-align: right; }
            
            /* Package Cost 欄位特別樣式 */
            .package-cell {
                background-color: #fff !important; /* 蓋過斑馬紋，保持白色 */
                font-weight: bold;
                color: #d93025; /* 紅字突顯 */
                border-bottom: 1px solid #c0c0c0;
            }

            /* 總計列樣式 */
            .total-row {
                background-color: #e8eaed !important;
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
                    <th style="min-width: 80px;">Station</th>
                    <th style="min-width: 100px;">Location</th>
                    <th style="min-width: 120px;">Program</th>
                    <th>Day-part</th>
                    <th>Size</th>
                    <th>Rate (Net)</th>
                    <th>Package-cost (Net)</th>
                    {% for i in range(1, 16) %}
                    <th>{{ i }}<br>日</th>
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
                    
                    {# 這裡處理 Package Cost 合併欄位 #}
                    {% if row.is_first %}
                        <td class="text-right package-cell" rowspan="{{ row.rowspan }}">
                            {{ "{:,}".format(row.package_cost) }}
                        </td>
                    {% endif %}

                    {# 填入每日次數 #}
                    {% for spot in row.spots %}
                    <td>{{ spot }}</td>
                    {% endfor %}
                </tr>
                {% endfor %}

                <tr class="total-row">
                    <td colspan="5" class="text-right">Total:</td>
                    <td class="text-right">{{ "{:,}".format(total_rate) }}</td>
                    <td></td> <td colspan="15"></td>
                </tr>
            </tbody>
        </table>

    </body>
    </html>
    """

    # ==========================================
    # 4. 渲染與顯示
    # ==========================================
    
    # 使用 Jinja2 填入資料
    template = Template(html_template)
    html_output = template.render(
        rows=processed_rows, 
        total_rate=total_rate
    )

    # 在 Streamlit 中顯示 HTML
    # height 設定為 600px, scrolling=True 允許表格過長時捲動
    st.subheader("📊 預覽結果")
    components.html(html_output, height=600, scrolling=True)

    # 下載按鈕
    st.download_button(
        label="📥 下載 HTML 報表",
        data=html_output,
        file_name="cue_schedule_report.html",
        mime="text/html"
    )

if __name__ == "__main__":
    main()
