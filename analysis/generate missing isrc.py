import os
import glob
import pandas as pd
import numpy as np

# 1. 設定輸入路徑與輸出路徑
base_path = "/Users/chu-chun/Mirror/Eva/input/sony_TME/"
output_path = "/Users/chu-chun/Mirror/Eva/output/TME_missing_ISCR_Report.xlsx"

# 2. 定義欄位的別名對照表（不分大小寫）
isrc_aliases = ["isrc", "歌曲isrc"]
# song_aliases = ["song", "mv", "music video", "track", "title", "歌曲名", "MV名"]
song_aliases = ["song", "歌曲名", "mv名"]
artist_aliases = ["artist", "歌手名"]

# 尋找所有 .xlsx 檔案（自動排除以 'bill' 開頭的檔案）
all_files = glob.glob(os.path.join(base_path, "**/*.xlsx"), recursive=True)
xlsx_files = [
    f for f in all_files 
    if not os.path.basename(f).lower().startswith('bill') 
    and not os.path.basename(f).startswith('~$')
]

print(f"找到 {len(xlsx_files)} 個 Excel 檔案準備讀取...")

all_dfs = []

# 開始逐一讀取檔案
for file in xlsx_files:
    try:
        # 先唯讀表頭 (nrows=0)，取得欄位名稱進行對照，避免一次讀入整張表
        header = pd.read_excel(file, nrows=0).columns
        rename_map = {}
        
        for col in header:
            col_str = str(col).strip().lower()
            if col_str in isrc_aliases:
                rename_map[col] = 'ISRC'
            elif col_str in song_aliases:
                rename_map[col] = 'Song'
            elif col_str in artist_aliases:
                rename_map[col] = 'Artist'
        
        # 如果對照表不為空，代表有匹配到我們想要的欄位
        if rename_map:
            # 只讀取有匹配到的這三列欄位 (usecols)
            df = pd.read_excel(file, usecols=list(rename_map.keys()))
            df = df.rename(columns=rename_map)
            
            # 防呆：如果有些檔案剛好缺了某一列，自動用 NaN 補齊，防止合併時出錯
            for target_col in ['ISRC', 'Song', 'Artist']:
                if target_col not in df.columns:
                    df[target_col] = np.nan
            
            # 額外記錄來源檔案名稱，方便你後續追蹤是哪張報表出問題
            df['source_file'] = os.path.basename(file)
            all_dfs.append(df)
            
    except Exception as e:
        print(f"讀取檔案失敗: {file}, 錯誤: {e}")

# 3. 合併所有資料並篩選缺失 ISRC
if not all_dfs:
    print("沒有讀取到任何含有對應欄位的 Excel 資料。")
else:
    # 合併多個 DataFrame
    df_all = pd.concat(all_dfs, ignore_index=True)
    print(f"合併完成！共有 {len(df_all)} 筆原始資料。")

    # 篩選 ISRC 為空值（包含 NaN、空字串、空格或字串 'nan'）的列
    isrc_missing_mask = (
        df_all['ISRC'].isna() |
        (df_all['ISRC'].astype(str).str.strip() == '') |
        (df_all['ISRC'].astype(str).str.strip().str.upper() == 'NAN')
    )
    df_missing = df_all[isrc_missing_mask].copy()
    print(f"篩選完成：共有 {len(df_missing)} 列資料缺失 ISRC。")

    # 4. 資料整理：將歌名與歌手統一轉為字串並去除首尾空格，避免分組時因為型態不同而重複
    df_missing['Song'] = df_missing['Song'].astype(str).str.strip()
    df_missing['Artist'] = df_missing['Artist'].astype(str).str.strip()

    # 5. 根據「歌名 + 歌手」分組，統計筆數（一列算一筆）並記錄來源檔案
    missing_report = (
        df_missing
        .groupby(['Song', 'Artist'])
        .agg(
            records=('Song', 'size'),  # 統計總筆數
            source_files=('source_file', lambda x: ", ".join(sorted(list(set(x.dropna())))))  # 顯示這些缺失出現在哪些檔案中
        )
        .reset_index()
    )

    # 依照筆數由高到低排序
    missing_report = missing_report.sort_values(by='records', ascending=False).reset_index(drop=True)

    # 6. 匯出 Excel
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    missing_report.to_excel(output_path, index=False)
    
    print("\n" + "="*50)
    print(f"🎉 缺失 ISRC 統計報告已產出！")
    print(f"💾 檔案路徑：{output_path}")
    print(f"📊 共有 {len(missing_report)} 首不重複的歌曲缺失 ISRC。")
    print("="*50)
