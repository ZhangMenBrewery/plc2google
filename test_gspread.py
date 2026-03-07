#!/usr/bin/env python3
"""使用 gspread 測試 Service Account 金鑰"""

import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 金鑰檔案路徑
KEY_FILE = r"D:\掌門事業股份有限公司\汐止廠 - 文件\釀造生產部\SourceCode\Other\plc2google\zmb54685508-c88132768091.json"

# 設定權限範圍
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

def test_gspread():
    """測試 gspread 連接"""
    try:
        print("=" * 50)
        print("使用 gspread 測試 Service Account 金鑰")
        print("=" * 50 + "\n")
        
        # 1. 從 JSON 金鑰檔案建立 credentials
        print("1. 載入金鑰檔案...")
        creds = ServiceAccountCredentials.from_json_keyfile_name(KEY_FILE, scope)
        print(f"   ✓ 金鑰載入成功")
        print(f"   - Client Email: {creds.service_account_email}\n")
        
        # 2. 驗證並登入
        print("2. 使用 gspread 授權...")
        client = gspread.authorize(creds)
        print("   ✓ 授權成功\n")
        
        # 3. 列出所有試算表
        print("3. 列出所有試算表...")
        try:
            spreadsheets = client.list_spreadsheet_files()
            print(f"   找到 {len(spreadsheets)} 個試算表:")
            for i, ss in enumerate(spreadsheets, 1):
                print(f"   {i}. {ss.get('name', '無標題')} (ID: {ss.get('id', 'N/A')})")
            print()
        except Exception as e:
            print(f"   ⚠ 無法列出試算表：{e}\n")
        
        # 4. 測試建立新的試算表
        print("4. 測試建立新的試算表...")
        from datetime import datetime
        test_title = f"Test-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        sh = client.create(test_title)
        print(f"   ✓ 建立成功：{sh.title} (ID: {sh.id})")
        
        # 5. 寫入測試數據
        print("5. 測試寫入數據...")
        worksheet = sh.sheet1
        test_data = [
            ['測試時間', gspread.utils.datetime_now().strftime('%Y-%m-%d %H:%M:%S')],
            ['測試項目', 'Service Account 金鑰測試'],
            ['測試結果', '成功']
        ]
        worksheet.update('A1', test_data)
        print("   ✓ 寫入成功")
        
        # 6. 讀取驗證
        print("6. 讀取驗證寫入的數據...")
        data = worksheet.get('A1:B3')
        print(f"   讀取的數據：{data}")
        print("   ✓ 驗證成功\n")
        
        # 7. 清理：刪除測試試算表
        print("7. 清理測試試算表...")
        client.delspreadsheet(sh.id)
        print("   ✓ 清理完成\n")
        
        print("=" * 50)
        print("測試結果：金鑰有效 ✓")
        print("=" * 50)
        return True
        
    except Exception as e:
        print(f"\n✗ 錯誤：{type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        print("\n" + "=" * 50)
        print("測試結果：金鑰無效 ✗")
        print("=" * 50)
        return False

if __name__ == "__main__":
    success = test_gspread()
    exit(0 if success else 1)
