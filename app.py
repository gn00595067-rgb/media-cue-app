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
        product_name = st.
