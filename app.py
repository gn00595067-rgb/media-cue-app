import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
from jinja2 import Template
import io

# ==========================================
# 1. 頁面設定與 CSS
# ==========================================
st.set_page_config(page_title="Cue表自動生成系統", layout="wide")

# 定義 HTML/CSS 模板 (樣式與之前相同)
html_template_str = """
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <style>
        body { font-family: "Microsoft JhengHei", sans-serif; margin: 0; padding: 10px; color: #333; }
        .header-info { background-color: #f1f3f4; padding: 15px; margin-bottom: 20px; border-left: 6px solid #1a73e8; }
        .header-info p { margin: 5px 0; font-weight: bold; font-size: 14px; }
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
        <p>客戶名稱：{{ client_name }}</p>
        <p>走期：{{ period }}</p>
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
                
                {% if row.is_first %}
                    <td class="text-right package-cell" rowspan="{{ row.rowspan }}">
                        {{ "{:,}".format(row.package_cost) }}
                    </td>
                {% endif %}

                {% for i in range(1, 16) %}
                <td>{{ row.get(i, 0) }}</td> {% endfor %}
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
# 2. 輔助功能：產生範本與處理資料
# ==========================================

def get_excel_template():
    """產生一個標準的 Excel 範本供使用者下載"""
    # 定義標準欄位
    columns = ['PackageGroup', 'Station', 'Location', 'Program', 'Daypart', 'Size', 'Rate']
    # 增加 1~15 號的欄位
    day_columns = [i for i in range(1, 16)]
    
    # 建立範例資料
    data = {
        'PackageGroup': ['A', 'A', 'B'], # 關鍵欄位：用來群組計算 Package Cost
        'Station': ['全家廣播', '全家廣播', '全家廣播'],
        'Location': ['北區', '桃竹苗', '中區'],
        'Program': ['北北基', '桃竹苗店', '中彰投'],
        'Daypart': ['00:00-24:00', '00:00-24:00', '00:00-24:00'],
        'Size': ['20秒', '20秒', '20秒'],
        'Rate': [416111, 249667, 200000]
    }
    
    df = pd.DataFrame(data)
    # 補上天數欄位 (預設填 50)
    for d in day_columns:
        df[d] = 50
        
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='CueData')
    return output.getvalue()

def process_uploaded_file(df):
    """處理上傳的 DataFrame，計算 Package Cost"""
    
    # 1. 確保欄位名稱正確 (轉成字串避免數字欄位出錯)
    df.columns = [str(c) for c in df.columns]
    
    # 2. 核心邏輯：計算 Package Cost (G20)
    # 依照 'PackageGroup' 欄位分組並加總 Rate
    if 'PackageGroup' not in df.columns:
        st.error("錯誤：Excel 中找不到 'PackageGroup' 欄位，無法計算組合價格。")
        return None, 0

    group_sums = df.groupby('PackageGroup')['Rate'].sum().to_dict()
    
    # 3. 整理資料結構給 Jinja2
    processed_rows = []
    seen_groups = set()
    
    for index, row in df.iterrows():
        group = row['PackageGroup']
        row_dict = row.to_dict()
        
        if group not in seen_groups:
            # 計算 rowspan (該群組有幾列)
            count = len(df[df['PackageGroup'] == group])
            row_dict['rowspan'] = count
            row_dict['package_cost'] = group_sums[group]
            row_dict['is_first'] = True
            seen_groups.add(group)
        else:
            row_dict['is_first'] = False
            
        # 處理日期欄位 (1~15)，將 NaN 轉為空字串或 0
        for i in range(1, 16):
            key = str(i)
            if key in row_dict:
                 # 如果是 NaN 轉成空字串，否則轉成整數
                val = row_dict[key]
                row_dict[i] = int(val) if pd.notna(val) else 0
            else:
                row_dict[i] = 0
                
        processed_rows.append(row_dict)
        
    total_rate = df['Rate'].sum()
    return processed_rows, total_rate

# ==========================================
# 3. 主程式介面 (Sidebar 與 Main)
# ==========================================

def main():
    st.sidebar.title("🎛️ 設定控制台")
    
    # Step 1: 下載範本
    st.sidebar.header("1. 下載資料範本")
    st.sidebar.markdown("請先下載 Excel 範本，填寫後上傳。")
    template_file = get_excel_template()
    st.sidebar.download_button(
        label="📥 下載 Excel 範本 (.xlsx)",
        data=template_file,
        file_name="cue_schedule_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.sidebar.markdown("---")
    
    # Step 2: 上傳檔案
    st.sidebar.header("2. 上傳 Cue 表資料")
    uploaded_file = st.sidebar.file_uploader("選擇 Excel 檔案", type=['xlsx'])
    
    # Step 3: 輸入基本資訊
    st.sidebar.header("3. 報表資訊")
    client_name = st.sidebar.text_input("客戶名稱", "萬國通路")
    period = st.sidebar.text_input("走期", "2025. 01. 01 - 2025. 01. 31")

    # 主畫面邏輯
    st.title("📺 廣播 Cue 表排程產生器")

    if uploaded_file is not None:
        try:
            # 讀取 Excel
            df = pd.read_excel(uploaded_file)
            
            # 顯示原始資料預覽 (Debug用)
            with st.expander("查看上傳的原始資料"):
                st.dataframe(df)

            # 處理資料
            rows, total_rate = process_uploaded_file(df)

            if rows:
                # 渲染 HTML
                template = Template(html_template_str)
                html_output = template.render(
                    rows=rows,
                    total_rate=total_rate,
                    client_name=client_name,
                    period=period
                )

                st.success("✅ 報表生成成功！")
                
                # 顯示 HTML
                st.subheader("報表預覽")
                components.html(html_output, height=600, scrolling=True)

                # 下載按鈕
                st.download_button(
                    label="📥 下載完整 HTML 報表",
                    data=html_output,
                    file_name="cue_report_final.html",
                    mime="text/html"
                )
        except Exception as e:
            st.error(f"檔案處理發生錯誤: {e}")
            st.warning("請確保您上傳的是從左側下載的標準範本格式。")
    else:
        st.info("👈 請從左側側邊欄下載範本，並上傳資料以開始使用。")

if __name__ == "__main__":
    main()
