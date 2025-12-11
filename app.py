import pandas as pd
from jinja2 import Template

# 1. 準備模擬資料
raw_data = [
    {
        "station": "全家便利商店通路廣播",
        "location": "北區-北區",
        "program": "北北基 1,649店",
        "daypart": "00:00-24:00",
        "size": "20秒",
        "rate": 416111,
        "package_group": "A",
        "spots": [50, 50, 50, 50, 50, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48]
    },
    {
        "station": "全家便利商店通路廣播",
        "location": "桃竹苗區-桃竹苗",
        "program": "桃竹苗 779店",
        "daypart": "00:00-24:00",
        "size": "20秒",
        "rate": 249667,
        "package_group": "A",
        "spots": [50, 50, 50, 50, 50, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48]
    },
    {
        "station": "全家便利商店通路廣播",
        "location": "中區-中區",
        "program": "中彰投 839店",
        "daypart": "00:00-24:00",
        "size": "20秒",
        "rate": 249667,
        "package_group": "B",
        "spots": [50, 50, 50, 50, 50, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48]
    }
]

# 2. 資料處理邏輯
df = pd.DataFrame(raw_data)
group_sums = df.groupby('package_group')['rate'].sum().to_dict()

processed_rows = []
seen_groups = set()

for index, row in df.iterrows():
    group = row['package_group']
    rate = row['rate']
    row_data = row.to_dict()
    
    if group not in seen_groups:
        group_count = len(df[df['package_group'] == group])
        row_data['rowspan'] = group_count
        row_data['package_cost'] = group_sums[group]
        row_data['is_first'] = True
        seen_groups.add(group)
    else:
        row_data['is_first'] = False
    
    processed_rows.append(row_data)

# 3. 定義 HTML 模板 (請注意這裡的開始與結束引號)
html_template = """
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <title>Cue 表網頁預覽</title>
    <style>
        body { font-family: "Microsoft JhengHei", Arial, sans-serif; margin: 20px; color: #333; }
        .header-info { background-color: #f9f9f9; padding: 15px; margin-bottom: 20px; border-left: 5px solid #333; }
        .header-info p { margin: 5px 0; font-weight: bold; }
        table { width: 100%; border-collapse: collapse; font-size: 12px; }
        th, td { border: 1px solid #ccc; padding: 8px 5px; text-align: center; vertical-align: middle; }
        th { background-color: #444; color: white; font-weight: normal; white-space: nowrap; }
        .text-left { text-align: left; }
        .text-right { text-align: right; }
        tbody tr:nth-child(even) { background-color: #f2f2f2; }
        tbody tr:hover { background-color: #e6f7ff; }
        .package-cell { background-color: #fff !important; font-weight: bold; color: #d9534f; }
    </style>
</head>
<body>
    <h2>4. Cue 表網頁預覽</h2>
    <div class="header-info">
        <p>客戶名稱：萬國通路</p>
        <p>Product：20秒、5秒</p>
        <p>Period：2025. 01. 01 - 2025. 01. 31</p>
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
        </tbody>
    </table>
</body>
</html>
""" 
# ↑↑↑ 請確保上面這個 """ 有被複製到 ↑↑↑

# 4. 渲染 HTML
template = Template(html_template)
html_output = template.render(rows=processed_rows)

# 5. 存檔或輸出
# 如果你在 Streamlit 環境，可以用 components.html(html_output)
# 這裡示範存成檔案
with open("cue_schedule_updated.html", "w", encoding="utf-8") as f:
    f.write(html_output)

print("HTML 生成完畢，請檢查 cue_schedule_updated.html")
