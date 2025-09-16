#!/usr/bin/env python3
"""
更新 .env 密碼工具
"""

import os
import getpass

def update_password():
    """更新 .env 檔案中的密碼"""
    print("=== 更新 M365 密碼 ===")
    print("請輸入新的應用程式密碼（或原始密碼）：")
    
    new_password = getpass.getpass("密碼: ")
    
    if not new_password:
        print("未輸入密碼，取消更新")
        return
    
    # 讀取現有的 .env 檔案
    env_path = ".env"
    lines = []
    
    if os.path.exists(env_path):
        with open(env_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
    
    # 更新密碼行
    updated = False
    for i, line in enumerate(lines):
        if line.startswith('M365_PASSWORD=') or line.startswith('# M365_PASSWORD='):
            lines[i] = f'M365_PASSWORD={new_password}\n'
            updated = True
            break
    
    # 如果沒找到密碼行，添加一行
    if not updated:
        lines.append(f'M365_PASSWORD={new_password}\n')
    
    # 寫回檔案
    with open(env_path, 'w', encoding='utf-8') as f:
        f.writelines(lines)
    
    print("✓ 密碼已更新")
    print("\n現在可以執行以下命令測試連接：")
    print("python3 test_connection.py")

if __name__ == "__main__":
    update_password()