import streamlit as st
import pandas as pd
import io
import xlsxwriter
from datetime import timedelta, date, datetime

# ==========================================
# 0. 系統設定
# ==========================================
st.set_page_config(page_title="東吳媒體 Cue 表生成系統", layout="wide")

# 預設單價 (可替換為真實邏輯)
UNIT_PRICES = {
    "全家便利商店": {"10s": 150, "15s": 200, "20s": 260},
    "全家新鮮視": {"10s": 400, "15s": 500, "20s": 600},
    "家樂福": {"10s": 130, "15s": 180, "20s": 230},
}

# ==========================================
# 1. 核心邏輯：計算每日檔次
# ==========================================
def calculate_schedule_data(start_d, end_d, budget_allocations):
    """
    依據預算分配，計算每一天的檔次
    回傳: (DataFrame 用於顯示, List 用於 Excel 生成)
    """
    days = (end_d - start_d).days + 1
    
    # 建立日期標題
    date_cols = []
    curr = start_d
    for _ in range(days):
        date_cols.append(curr)
        curr += timedelta(days=1)

    display_rows = []
    excel_rows = []
    
    total_cost_final = 0

    for item in budget_allocations:
        media = item['media']
        sec = item['seconds']
        budget = item['budget']
        
        if budget <= 0: continue
        
        # 取得單價
        price = UNIT_PRICES.get(media, {}).get(sec, 0)
        if price == 0: continue
            
        total_spots = int(budget / price)
        actual_cost = total_spots * price
        total_cost_final += actual_cost
        
        # 平均分配檔次 (模擬東吳 CSV 的每日數字)
        base = total_spots // days
        remainder = total_spots % days
        
        daily_spots = []
        for i in range(days):
            val = base + (1 if i < remainder else 0)
            daily_spots.append(val)
            
        # 準備資料
        row_data = {
            "媒體": media,
            "秒數": sec,
            "總檔次": total_spots,
            "費用": actual_cost
        }
        # 填入每日數據 (用於網頁顯示)
        for i, d in enumerate(date_cols):
            row_data[d.strftime('%m/%d')] = daily_spots[i]
            
        display_rows.append(row_data)
        
        # 準備 Excel 用資料 (保持原始型態)
        excel_rows.append({
            "media": media,
            "sec": sec,
            "total_spots": total_spots,
            "cost": actual_cost,
            "daily_spots": daily_spots
        })

    return pd.DataFrame(display_rows), excel_rows, date_cols, total_cost_final

# ==========================================
# 2. 東吳專屬 Excel 繪圖引擎 (重點！)
# ==========================================
def generate_dongwu_excel(client, start_d, end_d, data_rows, date_list):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book
        ws = wb.add_worksheet('東吳Cue表')
        
        # --- A. 定義樣式 (Styles) ---
        # 標題樣式
        fmt_title = wb.add_format({
            'bold': True, 'font_size': 16, 'font_name': '微軟正黑體',
            'align': 'center', 'valign': 'vcenter'
        })
        # 表頭樣式 (東吳風格：假設為深色底白字，或素雅風格)
        fmt_header = wb.add_format({
            'bold': True, 'font_size': 11, 'font_name': '微軟正黑體',
            'bg_color': '#44546A', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        # 日期表頭 (直式或橫式)
        fmt_date_header = wb.add_format({
            'bold': True, 'font_size': 10, 'font_name': 'Arial',
            'bg_color': '#D9E1F2', 'border': 1, 
            'align': 'center', 'valign': 'vcenter', 'rotation': 90 # 日期轉直的比較省空間
        })
        # 一般文字格
        fmt_text = wb.add_format({
            'font_size': 11, 'font_name': '微軟正黑體',
            'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        # 數字/金額格
        fmt_num = wb.add_format({
            'font_size': 11, 'font_name': 'Arial',
            'border': 1, 'align': 'center', 'valign': 'vcenter',
            'num_format': '#,##0'
        })
        # 金額格 (帶$)
        fmt_currency = wb.add_format({
            'font_size': 11, 'font_name': 'Arial',
            'border': 1, 'align': 'center', 'valign': 'vcenter',
            'num_format': '$#,##0'
        })
        # 資訊欄位 (左上角)
        fmt_info = wb.add_format({
            'font_size': 12, 'font_name': '微軟正黑體', 'bold': True
        })

        # --- B. 繪製表頭資訊 (Header Info) ---
        ws.merge_range('A1:H1', '媒體排程表 (Media Schedule)', fmt_title)
        
        ws.write('A3', f"客戶名稱：{client}", fmt_info)
        ws.write('A4', f"走期：{start_d.strftime('%Y/%m/%d')} - {end_d.strftime('%Y/%m/%d')}", fmt_info)
        
        # --- C. 繪製表格欄位 (Table Headers) ---
        # 固定欄位：媒體(A), 秒數(B), 總檔次(C), 費用(D)
        start_row = 5
        ws.write(start_row, 0, "媒體平台", fmt_header)
        ws.write(start_row, 1, "秒數", fmt_header)
        ws.write(start_row, 2, "總檔次", fmt_header)
        ws.write(start_row, 3, "費用 (未稅)", fmt_header)
        
        # 動態日期欄位 (從 E 欄開始)
        col_idx = 4
        for d in date_list:
            # 顯示格式：12/03 (三)
            w_str = ["(一)","(二)","(三)","(四)","(五)","(六)","(日)"][d.weekday()]
            d_str = f"{d.strftime('%m/%d')}\n{w_str}"
            ws.write(start_row, col_idx, d_str, fmt_date_header)
            col_idx += 1
            
        # --- D. 填入資料 (Data Rows) ---
        curr_row = start_row + 1
        for row in data_rows:
            ws.write(curr_row, 0, row['media'], fmt_text)
            ws.write(curr_row, 1, row['sec'], fmt_text)
            ws.write(curr_row, 2, row['total_spots'], fmt_num)
            ws.write(curr_row, 3, row['cost'], fmt_currency)
            
            # 填入每日檔次
            daily_col = 4
            for spots in row['daily_spots']:
                # 0 顯示為 "-" 看起來比較乾淨，或顯示空白
                val = spots if spots > 0 else "-"
                ws.write(curr_row, daily_col, val, fmt_num)
                daily_col += 1
            
            curr_row += 1
            
        # --- E. 調整欄寬 (Column Width) ---
        ws.set_column('A:A', 20) # 媒體
        ws.set_column('B:B', 10) # 秒數
        ws.set_column('C:D', 15) # 檔次與費用
        # 日期欄位設窄一點
        ws.set_column(4, 4 + len(date_list), 5) 
        
        # --- F. 加上合計列 (Footer) ---
        ws.write(curr_row, 0, "總計", fmt_header)
        ws.write(curr_row, 1, "", fmt_header)
        # Excel 公式 SUM
        ws.write_formula(curr_row, 2, f"=SUM(C{start_row+2}:C{curr_row})", fmt_header)
        ws.write_formula(curr_row, 3, f"=SUM(D{start_row+2}:D{curr_row})", fmt_header)
        
        # 每日合計公式
        for i in range(len(date_list)):
            col_letter = xlsxwriter.utility.xl_col_to_name(4 + i)
            ws.write_formula(curr_row, 4+i, f"=SUM({col_letter}{start_row+2}:{col_letter}{curr_row})", fmt_header)

    output.seek(0)
    return output

# ==========================================
# 3. UI 介面
# ==========================================
st.title("📄 東吳媒體 - 智慧 Cue 表生成器")

# 左側：輸入條件
with st.sidebar:
    st.header("1. 基礎設定")
    client_name = st.text_input("客戶名稱", "東吳測試專案")
    
    c1, c2 = st.columns(2)
    start_date = c1.date_input("開始日期", date.today())
    end_date = c2.date_input("結束日期", date.today() + timedelta(days=29))
    
    st.header("2. 預算分配")
    # 模擬輸入介面
    budget_fm = st.number_input("全家便利商店 (預算)", 0, 1000000, 165000, step=1000)
    sec_fm = st.selectbox("全家秒數", ["10s", "15s", "20s"], index=1)
    
    budget_fv = st.number_input("全家新鮮視 (預算)", 0, 1000000, 165000, step=1000)
    sec_fv = st.selectbox("新鮮視秒數", ["10s", "15s", "20s"], index=1)
    
    budget_cf = st.number_input("家樂福 (預算)", 0, 1000000, 57800, step=1000)
    sec_cf = st.selectbox("家樂福秒數", ["10s", "15s", "20s"], index=2)

    # 選擇下載格式 (未來擴充用)
    st.divider()
    format_type = st.selectbox("選擇匯出格式", ["東吳-格式", "聲活-格式(開發中)"])

# 整合輸入資料
allocations = [
    {"media": "全家便利商店", "budget": budget_fm, "seconds": sec_fm},
    {"media": "全家新鮮視", "budget": budget_fv, "seconds": sec_fv},
    {"media": "家樂福", "budget": budget_cf, "seconds": sec_cf},
]

# 執行計算
df_display, excel_rows, date_list, total_cost = calculate_schedule_data(start_date, end_date, allocations)

# 右側：預覽與下載
st.subheader(f"📊 {client_name} - 排程預覽")
st.metric("專案總金額", f"${total_cost:,}")

if not df_display.empty:
    st.dataframe(df_display, use_container_width=True)
    
    # 生成 Excel
    if format_type == "東吳-格式":
        excel_file = generate_dongwu_excel(client_name, start_date, end_date, excel_rows, date_list)
        file_name = f"Cue_{client_name}_東吳版.xlsx"
        
        st.download_button(
            label="📥 下載 Excel (東吳專用格式)",
            data=excel_file,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
else:
    st.info("請在左側輸入預算以產生報表")
