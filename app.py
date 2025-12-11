import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
from jinja2 import Template
import io

# 1. 基礎設定
st.set_page_config(page_title="Cue表生成系統", layout="wide")

# 2. 定義 HTML/CSS 樣式 (這是你要的美化部分：格線 + 斑馬紋)
html_template_str = """
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <style>
        /* 全域設定 */
        body { font-family: "Microsoft JhengHei", sans-serif; margin: 0; padding: 20px; color: #333; }
        
        /* 表頭資訊區 */
        .header-info { 
            background-color: #f8f9fa; 
            padding: 15px; 
            margin-bottom: 20px; 
            border-left: 5px solid #007bff; 
            border-radius: 4px;
        }
        .header-info p { margin: 5px 0; font-weight: bold; }
        .header-info span { font-weight: normal; color: #555; }

        /* 表格樣式 (核心美化) */
        table { 
            width: 100%; 
            border-collapse: collapse; /* 讓格線合併 */
            font-size: 13px; 
            white-space: nowrap; 
        }

        th, td { 
            border: 1px solid #bbb; /* 清楚的灰色格線 */
            padding: 10px 8px; 
            text-align: center; 
            vertical-align: middle; 
        }

        th { 
            background-color: #343a40; /* 深色表頭 */
            color: white; 
            position: sticky; 
            top: 0; 
        }

        /* 靠左與靠右對齊設定 */
        .text-left { text-align: left; }
        .text-right { text-align: right; }

        /* 斑馬紋 (隔行變色) */
        tbody tr:nth-child(even) { background-color: #f2f2f2; }
        tbody tr:hover { background-color: #e6f7ff; }

        /* Package Cost 特別樣式 (白色背景、紅色字) */
        .package-cell { 
            background-color: #fff !important; 
            font-weight: bold; 
            color: #d9534f; 
            border-bottom: 1px solid #bbb;
        }
        
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
                
                {# G20 合併儲存格邏輯 #}
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

# 3. 產生 Excel 範本功能
def get_template_excel():
    output = io.BytesIO()
    # 建立範本資料
    df = pd.DataFrame({
        'PackageGroup': ['A', 'A', 'B'], 
        'Station': ['全家廣播', '全家廣播', '全家廣播'],
        'Location': ['北區', '桃竹苗', '中區'],
        'Program': ['北北基', '桃竹苗', '中彰投'],
        'Daypart': ['全天', '全天', '全天'],
        'Size': ['20秒', '20秒', '20秒'],
        'Rate': [416111, 249667, 200000]
    })
    # 加入 1~15 號的欄位
    for i in range(1, 16):
        df[i] = 50
        
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# 4. 主程式
def main():
    st.title("📺 廣播 Cue 表排程系統")
    
    # --- 左側選單：業務輸入 ---
    with st.sidebar:
        st.header("1. 專案資訊")
        client_name = st.text_input("客戶名稱", "萬國通路")
        product_name = st.text_input("產品名稱", "形象廣告")
        period_input = st.text_input("走期", "2025.01.01 - 2025.01.31")
        budget_input = st.number_input("預算", value=1000000, step=10000)
        
        st.markdown("---")
        st.header("2. 資料準備")
        
        # 下載範本按鈕
        st.download_button(
            label="📥 下載 Excel 範本",
            data=get_template_excel(),
            file_name="cue_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # 上傳檔案
        uploaded_file = st.file_uploader("上傳填好的 Excel", type=['xlsx'])

    # --- 右側主畫面 ---
    if uploaded_file:
        try:
            # 讀取 Excel
            df = pd.read_excel(uploaded_file)
            
            # --- 資料處理邏輯 (G20) ---
            # 確保 Rate 是數字
            df['Rate'] = pd.to_numeric(df['Rate'], errors='coerce').fillna(0).astype(int)
            
            # 檢查是否有 PackageGroup 欄位，有的話就計算合併
            if 'PackageGroup' in df.columns:
                group_sums = df.groupby('PackageGroup')['Rate'].sum().to_dict()
            else:
                # 如果沒有這個欄位，每個都是獨立的
                df['PackageGroup'] = df.index
                group_sums = df.set_index('PackageGroup')['Rate'].to_dict()

            processed_rows = []
            seen_groups = set()
            
            for index, row in df.iterrows():
                row_dict = row.to_dict()
                group = row_dict.get('PackageGroup')
                
                # 判斷是否為群組第一筆 (為了 rowspan)
                if group not in seen_groups:
                    count = len(df[df['PackageGroup'] == group])
                    row_dict['rowspan'] = count
                    row_dict['package_cost'] = group_sums.get(group, 0)
                    row_dict['is_first'] = True
                    seen_groups.add(group)
                else:
                    row_dict['is_first'] = False
                
                # 處理日期欄位 1~15
                for i in range(1, 16):
                    # 處理欄位名稱可能是整數 1 或字串 "1"
                    val = row_dict.get(i) or row_dict.get(str(i))
                    row_dict[i] = int(val) if val else 0
                    
                processed_rows.append(row_dict)
            
            total_rate = df['Rate'].sum()

            # --- 渲染 HTML ---
            template = Template(html_template_str)
            html_output = template.render(
                client=client_name,
                product=product_name,
                period=period_input,
                budget="{:,}".format(budget_input),
                rows=processed_rows,
                total_rate=total_rate
            )
            
            st.success("✅ 報表生成成功！")
            
            # 顯示報表 (scrolling=True 讓寬表格可以左右滑)
            components.html(html_output, height=600, scrolling=True)
            
            # 下載按鈕
            st.download_button(
                label="📥 下載 HTML 檔案",
                data=html_output,
                file_name="cue_report.html",
                mime="text/html"
            )
            
        except Exception as e:
            st.error(f"檔案讀取錯誤：{e}")
            st.info("請確認上傳的 Excel 格式是否正確。")
            
    else:
        st.info("👈 請先從左側下載範本，填寫後上傳 Excel 檔案。")

if __name__ == "__main__":
    main()
