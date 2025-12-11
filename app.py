import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
from jinja2 import Template
from datetime import date

# 1. 頁面設定
st.set_page_config(page_title="Cue表排程系統 (三通路版)", layout="wide")

# 2. 媒體資料庫 (含詳細資料)
MEDIA_DB = {
    "無": {"Rate": 0, "Program": ""}, # 預設空選項
    "全家-北北基": {"Rate": 416111, "Program": "北北基 1,649店"},
    "全家-桃竹苗": {"Rate": 249667, "Program": "桃竹苗 779店"},
    "全家-中彰投": {"Rate": 249667, "Program": "中彰投 839店"},
    "全家-雲嘉南": {"Rate": 200000, "Program": "雲嘉南 900店"},
    "全家-高屏":   {"Rate": 200000, "Program": "高屏 720店"},
    "家樂福-全台": {"Rate": 350000, "Program": "量販全台 67店"},
    "家樂福-超市": {"Rate": 180000, "Program": "超市全台 245店"},
}

# 3. 核心邏輯：自動分配檔次 (這就是產生 44 44 42... 的地方)
def distribute_spots(total_spots, days=15):
    """
    將總檔數 (例如 640) 平均分配到 15 天。
    如果除不盡，前面的天數會多 1 檔。
    例子：640 / 15 -> 10天43檔, 5天42檔 (接近 44/42 的邏輯)
    """
    if total_spots <= 0:
        return [0] * days
    
    base = total_spots // days
    remainder = total_spots % days
    
    spots = []
    for i in range(days):
        if i < remainder:
            spots.append(base + 1)
        else:
            spots.append(base)
    return spots

# 4. HTML 樣式 (保留你喜歡的格線)
html_template = """
<!DOCTYPE html>
<html>
<head>
<style>
    body { font-family: "Microsoft JhengHei", sans-serif; padding: 10px; }
    .header { background: #f4f4f4; padding: 10px; border-left: 5px solid #2b5797; margin-bottom: 20px;}
    table { width: 100%; border-collapse: collapse; font-size: 13px; white-space: nowrap; }
    th { background: #333; color: #fff; padding: 8px; border: 1px solid #999; }
    td { padding: 8px; border: 1px solid #999; text-align: center; }
    .text-left { text-align: left; }
    .text-right { text-align: right; }
    tr:nth-child(even) { background: #f9f9f9; } /* 斑馬紋 */
</style>
</head>
<body>
    <div class="header">
        <p>客戶：{{ client }} | 產品：{{ product }} | 走期：{{ period }}</p>
    </div>
    <table>
        <thead>
            <tr>
                <th>Station</th>
                <th>Location</th>
                <th>Program</th>
                <th>Day-part</th>
                <th>Size</th>
                <th>Rate</th>
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
                <td>00:00-24:00</td>
                <td>20秒</td>
                <td class="text-right">{{ "{:,}".format(row.Rate) }}</td>
                {% for s in row.spots %}
                <td>{{ s }}</td>
                {% endfor %}
            </tr>
            {% endfor %}
            <tr style="font-weight:bold; background:#e0e0e0;">
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
    st.title("📺 Cue 表排程 (三通路比較版)")
    
    # --- Sidebar 設定 ---
    with st.sidebar:
        st.header("基本設定")
        client = st.text_input("客戶", "萬國通路")
        product = st.text_input("產品", "形象廣告")
        col_d1, col_d2 = st.columns(2)
        start_d = col_d1.date_input("開始", date(2025,1,1))
        end_d = col_d2.date_input("結束", date(2025,1,31))
        
    # --- 主畫面：三個通路的選擇區 (這是你要的選3個通路) ---
    st.subheader("🛠️ 設定排程組合")
    st.info("請在下方設定 3 個主要的媒體組合，系統會自動分配播放檔次 (例如 44/42)。")

    # 建立 3 個 Column 讓用戶選
    col1, col2, col3 = st.columns(3)
    
    selections = [] # 存使用者的選擇
    
    # 定義一個函數來產生每一欄的 UI
    def render_media_column(col, idx):
        with col:
            st.markdown(f"### 媒體 {idx}")
            # 選通路
            key_select = st.selectbox(f"選擇通路 {idx}", list(MEDIA_DB.keys()), key=f"sel_{idx}")
            
            # 只有選了非「無」的選項才顯示設定
            if key_select != "無":
                data = MEDIA_DB[key_select]
                st.caption(f"牌價: ${data['Rate']:,}")
                
                # 自動分配邏輯輸入框
                total_spots = st.number_input(f"總檔數 (15天) {idx}", value=640, step=10, key=f"spot_{idx}")
                
                # 計算分配結果 (這就是 44 44 42 的來源)
                spots_list = distribute_spots(total_spots)
                st.write(f"分配預覽: `{spots_list[:5]}...`") # 讓你看一下是不是 43, 43, 42...
                
                return {
                    "Station": key_select.split("-")[0], # 取前面當 Station
                    "Location": key_select.split("-")[-1], # 取後面當 Location
                    "Program": data["Program"],
                    "Rate": data["Rate"],
                    "spots": spots_list
                }
            return None

    # 執行渲染三欄
    sel1 = render_media_column(col1, "A")
    sel2 = render_media_column(col2, "B")
    sel3 = render_media_column(col3, "C")

    # 收集有選的資料
    if sel1: selections.append(sel1)
    if sel2: selections.append(sel2)
    if sel3: selections.append(sel3)

    st.markdown("---")

    # --- 生成報表區 ---
    if st.button("🚀 生成 / 更新報表", type="primary"):
        if not selections:
            st.warning("請至少選擇一個媒體！")
        else:
            # 計算總金額
            total_rate = sum([x["Rate"] for x in selections])
            period_str = f"{start_d} - {end_d}"
            
            # 渲染 HTML
            t = Template(html_template)
            html_out = t.render(
                client=client,
                product=product,
                period=period_str,
                rows=selections,
                total_rate=total_rate
            )
            
            st.subheader("📊 排程表預覽")
            components.html(html_out, height=500, scrolling=True)
            
            st.download_button("📥 下載報表", html_out, "cue_schedule.html")

if __name__ == "__main__":
    main()
