import os
import glob
import pandas as pd

# 1. 設定要搜尋的根目錄
base_path = "/Users/chu-chun/Mirror/Eva/input/sony_TME/"
output_path = "/Users/chu-chun/Mirror/Eva/output/xlsx_file_list.xlsx"

# 2. 遞迴搜尋該目錄下所有的 .xlsx 檔案
all_files = glob.glob(os.path.join(base_path, "**/*.xlsx"), recursive=True)

file_list = []
for file_path in all_files:
    # 取得檔名 (例如: "Song_clicks.xlsx")
    filename = os.path.basename(file_path)
    
    # 取得「直接上一層」的資料夾名稱 (例如: "2025.01")
    parent_folder = os.path.basename(os.path.dirname(file_path))
    
    # 取得「相對於根目錄」的完整路徑 (例如: "2025/2025.01/Song_clicks.xlsx")
    relative_path = os.path.relpath(file_path, base_path)
    
    # 取得該檔案的完整目錄路徑（不含檔名）
    relative_dir = os.path.dirname(relative_path)
    
    file_list.append({
        "檔名": filename,
        "直接上一層資料夾": parent_folder,
        "相對資料夾路徑": relative_dir,
        "相對檔案路徑": relative_path,
        "絕對路徑": file_path
    })

# 3. 轉成 DataFrame
df_files = pd.DataFrame(file_list)

# 4. 排序並匯出為 Excel
if not df_files.empty:
    df_files = df_files.sort_values(by="相對檔案路徑").reset_index(drop=True)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    df_files.to_excel(output_path, index=False)
    
    print(f"成功抓取！共有 {len(df_files)} 個 .xlsx 檔案。")
    print(f"清單已成功匯出至：{output_path}")
    
    # 在 Notebook 中預覽前 10 筆
    print("\n【清單前 10 筆預覽】")
    print(df_files[["直接上一層資料夾", "相對資料夾路徑", "檔名"]].head(10).to_string())
else:
    print("在此目錄下找不到任何 .xlsx 檔案。")
