import os
import glob
import pandas as pd

# 1. 設定來源根目錄與輸出報表路徑
base_path = "/Users/chu-chun/Mirror/Eva/input/sony_TME/"
output_path = "/Users/chu-chun/Mirror/Eva/output/bill_revenue_report.xlsx"

# 2. 搜尋所有 .xlsx 檔案，並過濾出檔名以 "bill_" 開頭的檔案 (不區分大小寫)
all_files = glob.glob(os.path.join(base_path, "**/*.xlsx"), recursive=True)
bill_files = [
    f for f in all_files 
    if os.path.basename(f).lower().startswith("bill_") 
    and not os.path.basename(f).startswith("~$")
]

data = []

# 3. 讀取每個檔案的「结算汇总」Sheet 內特定儲存格
for file in bill_files:
    # 讀取 Sheet (不設 header 以便精確對應 Row/Col 索引)
    df = pd.read_excel(file, sheet_name="结算汇总", header=None)
    
    # B8 儲存格 (Row 8, Column B) 對應 Pandas 索引 [7, 1]
    b8_val = df.iloc[7, 1]
    
    # C17 儲存格 (Row 17, Column C) 對應 Pandas 索引 [16, 2]
    c17_val = df.iloc[17, 2]
    
    # 處理日期：如果是日期物件則格式化為 YYYYMM，否則取字串前 6 位
    if hasattr(b8_val, "strftime"):
        date_str = b8_val.strftime("%Y%m")
    else:
        date_str = str(b8_val).strip()[:6]
        
    data.append({
        "日期": date_str,
        "金額": c17_val,
        "完整路徑": os.path.dirname(file)
    })

# 4. 匯出結果為 Excel 報表
df_result = pd.DataFrame(data)
os.makedirs(os.path.dirname(output_path), exist_ok=True)
df_result.to_excel(output_path, index=False)

print(f"🎉 掃描完成！已將結果匯出至：{output_path}")
