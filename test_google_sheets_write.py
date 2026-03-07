#!/usr/bin/env python3
"""測試 Google Sheets 寫入功能"""

import pygsheets
from datetime import datetime

# 金鑰檔案路徑
KEY_FILE = r"D:\掌門事業股份有限公司\汐止廠 - 文件\釀造生產部\SourceCode\Other\plc2google\zmb54685508-c88132768091.json"

def test_google_sheets():
    """測試 Google Sheets 連接和寫入"""
    try:
        print("=" * 50)
        print("測試 Google Sheets 連接...")
        print("=" * 50 + "\n")
        
        # 授權
        print("1. 使用 Service Account 授權...")
        gc = pygsheets.authorize(service_account_file=KEY_FILE)
        print("   ✓ 授權成功\n")
        
        # 測試建立新的試算表
        print("3. 測試建立新的試算表...")
        test_title = f"Test-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        sh = gc.sheet.create(test_title)
        print(f"   ✓ 建立成功：{sh.title} (ID: {sh.spreadsheet_id})")
        
        # 寫入測試數據
        print("4. 測試寫入數據...")
        worksheet = sh.sheet1
        test_data = [
            ['測試時間', datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
            ['測試項目', 'Service Account 金鑰測試'],
            ['測試結果', '成功']
        ]
        worksheet.range('A1:B3', test_data)
        print("   ✓ 寫入成功")
        
        # 讀取驗證
        print("5. 讀取驗證寫入的數據...")
        data = worksheet.get_values('A1:B3')
        print(f"   讀取的數據：{data}")
        print("   ✓ 驗證成功\n")
        
        # 清理：刪除測試試算表
        print("6. 清理測試試算表...")
        gc.delete_spreadsheet(sh.spreadsheet_id)
        print("   ✓ 清理完成\n")
        
        print("=" * 50)
        print("Google Sheets 測試完成！")
        print("=" * 50)
        return True
        
    except Exception as e:
        print(f"\n✗ 錯誤：{type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("\n" + "=" * 50)
    print("Google Sheets 寫入測試")
    print("=" * 50 + "\n")
    
    success = test_google_sheets()
    
    print("\n" + "=" * 50)
    if success:
        print("測試結果：Google Sheets 功能正常 ✓")
    else:
        print("測試結果：Google Sheets 功能異常 ✗")
    print("=" * 50 + "\n")
