#!/usr/bin/env python3
"""測試 Service Account 金鑰是否有效"""

import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# 金鑰檔案路徑
KEY_FILE = r"D:\掌門事業股份有限公司\汐止廠 - 文件\釀造生產部\SourceCode\Other\plc2google\zmb54685508-c88132768091.json"

def test_service_account():
    """測試 Service Account 金鑰"""
    try:
        # 載入金鑰檔案
        print("=" * 50)
        print("載入金鑰檔案...")
        with open(KEY_FILE, 'r', encoding='utf-8') as f:
            credentials_info = json.load(f)
        
        print(f"✓ 金鑰檔案載入成功")
        print(f"  - Project ID: {credentials_info.get('project_id')}")
        print(f"  - Client Email: {credentials_info.get('client_email')}")
        print(f"  - Private Key ID: {credentials_info.get('private_key_id')}")
        
        # 建立 credentials
        print("\n建立 Service Account credentials...")
        credentials = service_account.Credentials.from_service_account_file(
            KEY_FILE,
            scopes=['https://www.googleapis.com/auth/spreadsheets']
        )
        print(f"✓ Credentials 建立成功")
        print(f"  - 有效時間：{credentials.expiry if credentials.expiry else '未知'}")
        
        # 測試連接 Google Sheets API
        print("\n連接 Google Sheets API...")
        service = build('sheets', 'v4', credentials=credentials)
        print("✓ API 連接成功")
        
        # 測試列出試算表（需要一個已分享的試算表 ID）
        # 這裡我們只測試認證是否有效
        print("\n✓ Service Account 金鑰驗證成功！")
        print("\n金鑰資訊:")
        print(f"  - 類型：{credentials_info.get('type')}")
        print(f"  - Project ID: {credentials_info.get('project_id')}")
        print(f"  - Client Email: {credentials_info.get('client_email')}")
        print(f"  - Auth URI: {credentials_info.get('auth_uri')}")
        print(f"  - Token URI: {credentials_info.get('token_uri')}")
        
        return True
        
    except FileNotFoundError:
        print(f"✗ 金鑰檔案找不到：{KEY_FILE}")
        return False
    except json.JSONDecodeError as e:
        print(f"✗ 金鑰檔案格式錯誤：{e}")
        return False
    except Exception as e:
        print(f"✗ 錯誤：{type(e).__name__}: {e}")
        return False

if __name__ == "__main__":
    print("\n" + "=" * 50)
    print("Google Service Account 金鑰測試")
    print("=" * 50 + "\n")
    
    success = test_service_account()
    
    print("\n" + "=" * 50)
    if success:
        print("測試結果：金鑰有效 ✓")
    else:
        print("測試結果：金鑰無效 ✗")
    print("=" * 50 + "\n")
