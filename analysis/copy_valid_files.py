import os
import glob
import shutil

# 1. 設定來源資料夾與目標資料夾路徑
src_dir = "/Users/chu-chun/Mirror/Eva/input/sony_網易/"
dest_dir = "/Users/chu-chun/Mirror/Eva/input/sony_網易_valid/"

# 確保目標資料夾存在，若不存在則自動建立
os.makedirs(dest_dir, exist_ok=True)

# 2. 遞迴搜尋所有的 .xlsx 和 .xls 檔案
all_files = glob.glob(os.path.join(src_dir, "**/*.xlsx"), recursive=True) + \
            glob.glob(os.path.join(src_dir, "**/*.xls"), recursive=True)

copied_count = 0
conflict_count = 0

print(f"🔍 開始掃描資料夾：{src_dir} ...")

for file_path in all_files:
    filename = os.path.basename(file_path)
    
    # 排除 Excel 產生的暫存鎖定檔案 (以 ~$ 開頭的檔案)
    if filename.startswith("~$"):
        continue
        
    # 3. 檢查檔名是否包含「結算」或「结算」字眼 (支援繁簡體)
    if "結算" in filename or "结算" in filename:
        # 設定目標路徑
        dest_path = os.path.join(dest_dir, filename)
            
        # 5. 執行複製 (shutil.copy2 會連同檔案的建立/修改時間等元資料一併複製)
        try:
            shutil.copy2(file_path, dest_path)
            copied_count += 1
            print(f"✅ 成功複製: {filename} -> {filename}")
        except Exception as e:
            print(f"❌ 複製失敗: {filename}, 錯誤: {e}")

print("\n" + "="*50)
print(f"🎉 複製工作已結束！")
print(f"📂 目標資料夾：{dest_dir}")
print(f"📊 共成功複製 {copied_count} 個檔案。")
if conflict_count > 0:
    print(f"   (其中有 {conflict_count} 個檔案因檔名重複已自動加序號重新命名，防止覆蓋)")
print("="*50)
