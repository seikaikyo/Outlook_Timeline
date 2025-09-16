#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
M365 IMAP 連接測試工具
"""

import imaplib
import sys
import getpass
from dotenv import load_dotenv
import os

# 載入環境變數
load_dotenv()

def test_imap_connection():
    """測試 IMAP 連接"""
    
    # 取得帳號資訊
    username = os.getenv('M365_USERNAME')
    password = os.getenv('M365_PASSWORD')
    server = os.getenv('IMAP_SERVER', 'outlook.office365.com')
    port = int(os.getenv('IMAP_PORT', 993))
    
    print("=== M365 IMAP 連接診斷工具 ===")
    print(f"帳號: {username}")
    print(f"伺服器: {server}:{port}")
    print()
    
    # 如果沒有密碼，詢問使用者
    if not password:
        print("請輸入密碼（建議使用應用程式密碼）:")
        password = getpass.getpass()
    
    # 測試連接
    try:
        print("正在連接到伺服器...")
        connection = imaplib.IMAP4_SSL(server, port)
        print("✓ SSL 連接成功")
        
        print("正在嘗試登入...")
        connection.login(username, password)
        print("✓ 登入成功!")
        
        # 取得資料夾清單
        print("\n正在取得資料夾清單...")
        status, folders = connection.list()
        if status == 'OK':
            print(f"✓ 找到 {len(folders)} 個資料夾:")
            for folder in folders[:5]:  # 只顯示前5個
                folder_name = folder.decode().split('"')[-2]
                print(f"  - {folder_name}")
            if len(folders) > 5:
                print(f"  ... 還有 {len(folders)-5} 個資料夾")
        
        # 測試選擇收件夾
        print("\n正在測試收件夾...")
        connection.select('INBOX')
        status, messages = connection.search(None, 'ALL')
        if status == 'OK':
            msg_count = len(messages[0].split()) if messages[0] else 0
            print(f"✓ 收件夾有 {msg_count} 封郵件")
        
        connection.logout()
        print("\n✓ 所有測試通過！IMAP 設定正確。")
        return True
        
    except imaplib.IMAP4.error as e:
        print(f"\n✗ IMAP 錯誤: {e}")
        print("\n可能的解決方案:")
        print("1. 確認 IMAP 已在 Outlook 設定中啟用")
        print("2. 如果啟用了多重驗證 (MFA)，請使用應用程式密碼")
        print("3. 檢查帳號密碼是否正確")
        return False
        
    except Exception as e:
        print(f"\n✗ 連接錯誤: {e}")
        print("\n可能的解決方案:")
        print("1. 檢查網路連接")
        print("2. 確認伺服器位址和端口正確")
        print("3. 檢查防火牆設定")
        return False

def check_outlook_settings():
    """檢查 Outlook 設定指引"""
    print("\n=== Outlook IMAP 設定指引 ===")
    print("1. 登入 Outlook 網頁版 (https://outlook.office365.com)")
    print("2. 點擊右上角設定圖示 → 檢視所有 Outlook 設定")
    print("3. 選擇「郵件」→「同步處理電子郵件」")
    print("4. 確認「POP 和 IMAP」已啟用")
    print()
    print("=== 應用程式密碼設定 ===")
    print("1. 登入 Microsoft 帳戶安全性 (https://account.microsoft.com/security)")
    print("2. 選擇「進階安全性選項」")
    print("3. 在「應用程式密碼」區塊選擇「建立新的應用程式密碼」")
    print("4. 輸入應用程式名稱（例如：Outlook Timeline）")
    print("5. 記下產生的密碼並替換 .env 檔案中的 M365_PASSWORD")

if __name__ == "__main__":
    success = test_imap_connection()
    
    if not success:
        print()
        response = input("是否顯示設定指引？(y/n): ")
        if response.lower() in ['y', 'yes', '是']:
            check_outlook_settings()