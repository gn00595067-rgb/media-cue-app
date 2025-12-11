import pandas as pd
from jinja2 import Template

# 1. 準備模擬資料 (模擬你截圖中的數據)
# 我們增加了一個 'package_group' 欄位來模擬為什麼前兩筆會被算在一起
raw_data = [
    {
        "station": "全家便利商店通路廣播",
        "location": "北區-北區",
        "program": "北北基 1,649店",
        "daypart": "00:00-24:00",
        "size": "20秒",
        "rate": 416111,
        "package_group": "A",  # 群組 A
        "spots": [50, 50, 50, 50, 50, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48] # 模擬每天的次數
    },
    {
        "station": "全家便利商店通路廣播",
        "location": "桃竹苗區-桃竹苗",
        "program": "桃竹苗 779店",
        "daypart": "00:00-24:00",
        "size": "20秒",
        "rate": 249667,
        "package_group": "A",  # 群組 A (跟上面同一組，所以錢會加在一起)
        "spots": [50, 50, 50, 50, 50, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48]
    },
    {
        "station": "全家便利商店通路廣播",
        "location": "中區-中區",
        "program": "中彰投 839店",
        "daypart": "00:00-24:00",
        "size": "20秒",
        "rate": 249667,
        "package_group": "B",  # 群組 B (新的群組)
        "spots": [50, 50, 50, 50, 50, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48]
    }
]

# 2. 資料處理邏輯：計算 Package Cost (G20 的秘密)
# 我們需要將資料依照 package_group 分組，並計算總金額
df = pd.DataFrame(raw_data)

# 計算每個群組的總金額 (Total Package Cost)
group_sums = df.groupby('package_group')['rate'].sum().to_dict()

# 將資料轉換為適合 HTML rowspan 的結構
processed_rows = []
seen_groups = set()

for index, row in df.iterrows():
    group = row['package_group']
    rate = row['rate']
    
    row_data = row.to_dict()
    
    # 判斷是否為該群組的第一筆資料
    if group not in seen_groups:
        # 如果是第一筆，計算該群組有幾列 (rowspan用)
        group_count = len(df[df['package_group'] == group])
        row_data['rowspan'] = group_count
        row_data['package_cost'] = group_sums[group] # 這裡就是填入 665,778 的地方
        row_data['is_first'] = True
        seen_groups.add(group)
    else:
        # 如果不是第一筆，就不顯示 Package Cost
        row_data['is_first'] = False
    
    processed_rows.append(row_data)

# 3. 定義 HTML 模板 (包含你要求的 CSS 格線與背景色)
html_template = """
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <title>Cue 表網頁預覽</title>
    <style>
        body {
            font-family: "Microsoft JhengHei", Arial, sans-serif;
            margin: 20px;
            color: #333;
        }
        
        /* 標題區塊 */
        .header-info {
            background-color: #f9f9f9;
            padding: 15px;
            margin-bottom: 20px;
            border-left: 5px solid #333;
        }
        .header-info p { margin: 5px 0; font-weight: bold; }

        /*
