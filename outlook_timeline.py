#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook Timeline - M365 郵件關鍵字搜尋與時間軸分析工具
"""

import imaplib
import email
import email.header
import re
import json
import csv
import os
from datetime import datetime, timedelta
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass
from email.mime.text import MIMEText
import argparse
import getpass
import sys
from dotenv import load_dotenv

# 載入環境變數
load_dotenv()

@dataclass
class EmailInfo:
    """郵件資訊結構"""
    uid: str
    subject: str
    sender: str
    receiver: str
    date: datetime
    body: str
    attachments: List[str]
    folder: str
    keywords_found: List[str]

class OutlookTimeline:
    """M365 Outlook 郵件時間軸分析器"""
    
    def __init__(self, username: str = None, password: str = None):
        self.username = username or os.getenv('M365_USERNAME')
        self.password = password or os.getenv('M365_PASSWORD')
        self.imap_server = os.getenv('IMAP_SERVER', "outlook.office365.com")
        self.imap_port = int(os.getenv('IMAP_PORT', 993))
        self.connection = None
        self.emails: List[EmailInfo] = []
        
    def connect(self) -> bool:
        """連接到M365 IMAP伺服器"""
        try:
            self.connection = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
            self.connection.login(self.username, self.password)
            print(f"✓ 成功連接到 {self.username}")
            return True
        except Exception as e:
            print(f"✗ 連接失敗: {e}")
            return False
    
    def disconnect(self):
        """中斷IMAP連接"""
        if self.connection:
            self.connection.logout()
            print("✓ 已中斷連接")
    
    def get_folders(self) -> List[str]:
        """取得所有郵件資料夾"""
        if not self.connection:
            return []
        
        try:
            status, folders = self.connection.list()
            folder_list = []
            for folder in folders:
                folder_name = folder.decode().split('"')[-2]
                folder_list.append(folder_name)
            return folder_list
        except Exception as e:
            print(f"✗ 取得資料夾失敗: {e}")
            return []
    
    def decode_header(self, header: str) -> str:
        """解碼郵件標頭"""
        try:
            decoded_header = email.header.decode_header(header)
            result = ""
            for text, encoding in decoded_header:
                if isinstance(text, bytes):
                    if encoding:
                        result += text.decode(encoding)
                    else:
                        result += text.decode('utf-8', errors='ignore')
                else:
                    result += text
            return result.strip()
        except:
            return header
    
    def extract_email_body(self, msg) -> str:
        """提取郵件內容"""
        body = ""
        
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == "text/plain":
                    charset = part.get_content_charset() or 'utf-8'
                    try:
                        body += part.get_payload(decode=True).decode(charset, errors='ignore')
                    except:
                        body += str(part.get_payload())
        else:
            charset = msg.get_content_charset() or 'utf-8'
            try:
                body = msg.get_payload(decode=True).decode(charset, errors='ignore')
            except:
                body = str(msg.get_payload())
        
        return body.strip()
    
    def search_emails(self, keywords: List[str], folders: List[str] = None, 
                     days_back: int = 30, include_sent: bool = True) -> List[EmailInfo]:
        """搜尋包含關鍵字的郵件"""
        if not self.connection:
            print("✗ 請先連接到伺服器")
            return []
        
        if folders is None:
            folders = ["INBOX"]
            if include_sent:
                folders.append("Sent Items")
        
        all_emails = []
        
        # 計算日期範圍
        since_date = (datetime.now() - timedelta(days=days_back)).strftime("%d-%b-%Y")
        
        for folder in folders:
            try:
                print(f"🔍 搜尋資料夾: {folder}")
                self.connection.select(folder)
                
                # 先按日期篩選，再在本地搜尋關鍵字
                status, messages = self.connection.search(None, f'SINCE {since_date}')
                
                if status != 'OK' or not messages[0]:
                    continue
                
                email_ids = messages[0].split()
                print(f"  找到 {len(email_ids)} 封郵件 (最近 {days_back} 天)")
                
                for email_id in email_ids:
                    try:
                        # 取得郵件
                        status, msg_data = self.connection.fetch(email_id, '(RFC822)')
                        if status != 'OK':
                            continue
                        
                        msg = email.message_from_bytes(msg_data[0][1])
                        
                        # 解析郵件資訊
                        subject = self.decode_header(msg.get('Subject', ''))
                        sender = self.decode_header(msg.get('From', ''))
                        receiver = self.decode_header(msg.get('To', ''))
                        date_str = msg.get('Date', '')
                        
                        # 解析日期
                        try:
                            date = email.utils.parsedate_to_datetime(date_str)
                        except:
                            date = datetime.now()
                        
                        # 提取郵件內容
                        body = self.extract_email_body(msg)
                        
                        # 搜尋關鍵字
                        found_keywords = []
                        search_text = f"{subject} {body}".lower()
                        
                        for keyword in keywords:
                            if keyword.lower() in search_text:
                                found_keywords.append(keyword)
                        
                        # 如果找到關鍵字，加入結果
                        if found_keywords:
                            email_info = EmailInfo(
                                uid=email_id.decode(),
                                subject=subject,
                                sender=sender,
                                receiver=receiver,
                                date=date,
                                body=body[:1000],  # 限制內容長度
                                attachments=[],  # 可以後續擴展
                                folder=folder,
                                keywords_found=found_keywords
                            )
                            all_emails.append(email_info)
                            
                    except Exception as e:
                        print(f"  ⚠️ 處理郵件時發生錯誤: {e}")
                        continue
                        
            except Exception as e:
                print(f"✗ 搜尋資料夾 {folder} 時發生錯誤: {e}")
                continue
        
        # 按日期排序
        all_emails.sort(key=lambda x: x.date)
        self.emails = all_emails
        
        print(f"✓ 總共找到 {len(all_emails)} 封相關郵件")
        return all_emails
    
    def generate_timeline_report(self, output_format: str = "json") -> str:
        """產生時間軸報告"""
        if not self.emails:
            return "沒有找到相關郵件"
        
        report_data = {
            "generated_at": datetime.now().isoformat(),
            "total_emails": len(self.emails),
            "date_range": {
                "start": self.emails[0].date.isoformat(),
                "end": self.emails[-1].date.isoformat()
            },
            "timeline": []
        }
        
        for email_info in self.emails:
            timeline_entry = {
                "date": email_info.date.isoformat(),
                "subject": email_info.subject,
                "sender": email_info.sender,
                "receiver": email_info.receiver,
                "folder": email_info.folder,
                "keywords_found": email_info.keywords_found,
                "preview": email_info.body[:200] + "..." if len(email_info.body) > 200 else email_info.body
            }
            report_data["timeline"].append(timeline_entry)
        
        if output_format.lower() == "json":
            return json.dumps(report_data, ensure_ascii=False, indent=2)
        elif output_format.lower() == "csv":
            return self._generate_csv_report()
        elif output_format.lower() == "html":
            return self._generate_html_report()
        else:
            return self._generate_text_report()
    
    def _generate_csv_report(self) -> str:
        """產生CSV格式報告"""
        import io
        output = io.StringIO()
        writer = csv.writer(output)
        
        # 寫入標頭
        writer.writerow(['日期', '主旨', '寄件者', '收件者', '資料夾', '找到的關鍵字', '內容預覽'])
        
        # 寫入資料
        for email_info in self.emails:
            writer.writerow([
                email_info.date.strftime('%Y-%m-%d %H:%M:%S'),
                email_info.subject,
                email_info.sender,
                email_info.receiver,
                email_info.folder,
                ', '.join(email_info.keywords_found),
                email_info.body[:200] + "..." if len(email_info.body) > 200 else email_info.body
            ])
        
        return output.getvalue()
    
    def _generate_text_report(self) -> str:
        """產生文字格式報告"""
        report = f"=== Outlook 郵件時間軸報告 ===\n"
        report += f"產生時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        report += f"總計郵件數: {len(self.emails)}\n"
        report += f"時間範圍: {self.emails[0].date.strftime('%Y-%m-%d')} ~ {self.emails[-1].date.strftime('%Y-%m-%d')}\n\n"
        
        for i, email_info in enumerate(self.emails, 1):
            report += f"[{i}] {email_info.date.strftime('%Y-%m-%d %H:%M:%S')}\n"
            report += f"    主旨: {email_info.subject}\n"
            report += f"    寄件者: {email_info.sender}\n"
            report += f"    收件者: {email_info.receiver}\n"
            report += f"    資料夾: {email_info.folder}\n"
            report += f"    關鍵字: {', '.join(email_info.keywords_found)}\n"
            report += f"    內容預覽: {email_info.body[:150]}...\n"
            report += "-" * 80 + "\n"
        
        return report
    
    def _generate_html_report(self) -> str:
        """產生HTML格式報告"""
        html_template = """<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Outlook 郵件時間軸報告</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Microsoft JhengHei', 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            background-color: #f5f5f5;
            padding: 20px;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            font-weight: 300;
        }
        
        .stats {
            display: flex;
            justify-content: space-around;
            margin-top: 20px;
            flex-wrap: wrap;
        }
        
        .stat-item {
            text-align: center;
            min-width: 150px;
        }
        
        .stat-number {
            font-size: 2em;
            font-weight: bold;
            display: block;
        }
        
        .stat-label {
            font-size: 0.9em;
            opacity: 0.9;
        }
        
        .content {
            padding: 30px;
        }
        
        .search-info {
            background: #e8f4fd;
            border-left: 4px solid #2196F3;
            padding: 15px;
            margin-bottom: 30px;
            border-radius: 0 8px 8px 0;
        }
        
        .timeline {
            position: relative;
            margin: 20px 0;
        }
        
        .timeline::before {
            content: '';
            position: absolute;
            left: 30px;
            top: 0;
            bottom: 0;
            width: 2px;
            background: #e0e0e0;
        }
        
        .email-item {
            position: relative;
            margin-bottom: 30px;
            padding-left: 70px;
            border-radius: 8px;
            background: #ffffff;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        
        .email-item:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
        }
        
        .email-item::before {
            content: '';
            position: absolute;
            left: -39px;
            top: 20px;
            width: 16px;
            height: 16px;
            border: 3px solid #2196F3;
            border-radius: 50%;
            background: white;
        }
        
        .email-header {
            background: #f8f9fa;
            padding: 20px;
            border-bottom: 1px solid #e9ecef;
        }
        
        .email-subject {
            font-size: 1.3em;
            font-weight: bold;
            color: #2c3e50;
            margin-bottom: 10px;
        }
        
        .email-meta {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 15px;
            font-size: 0.9em;
            color: #666;
        }
        
        .email-body {
            padding: 20px;
        }
        
        .email-preview {
            background: #f8f9fa;
            border-radius: 6px;
            padding: 15px;
            margin: 10px 0;
            font-style: italic;
            color: #555;
        }
        
        .keywords {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin: 15px 0;
        }
        
        .keyword-tag {
            background: #ff6b6b;
            color: white;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 0.8em;
            font-weight: bold;
        }
        
        .meta-item {
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .meta-icon {
            width: 16px;
            height: 16px;
            opacity: 0.7;
        }
        
        .folder-tag {
            background: #28a745;
            color: white;
            padding: 2px 8px;
            border-radius: 12px;
            font-size: 0.75em;
            font-weight: bold;
        }
        
        .date-badge {
            background: #6c757d;
            color: white;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 0.8em;
            font-weight: bold;
        }
        
        @media (max-width: 768px) {
            body {
                padding: 10px;
            }
            
            .header {
                padding: 20px;
            }
            
            .header h1 {
                font-size: 2em;
            }
            
            .content {
                padding: 20px;
            }
            
            .email-meta {
                grid-template-columns: 1fr;
                gap: 10px;
            }
            
            .stats {
                flex-direction: column;
                gap: 15px;
            }
        }
        
        .footer {
            background: #343a40;
            color: white;
            text-align: center;
            padding: 20px;
            font-size: 0.9em;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📧 Outlook 郵件時間軸報告</h1>
            <div class="stats">
                <div class="stat-item">
                    <span class="stat-number">{total_emails}</span>
                    <span class="stat-label">總計郵件</span>
                </div>
                <div class="stat-item">
                    <span class="stat-number">{date_range_days}</span>
                    <span class="stat-label">天數範圍</span>
                </div>
                <div class="stat-item">
                    <span class="stat-number">{keywords_count}</span>
                    <span class="stat-label">搜尋關鍵字</span>
                </div>
            </div>
        </div>
        
        <div class="content">
            <div class="search-info">
                <h3>🔍 搜尋資訊</h3>
                <p><strong>時間範圍：</strong>{start_date} 至 {end_date}</p>
                <p><strong>搜尋關鍵字：</strong>{all_keywords}</p>
                <p><strong>報告產生時間：</strong>{generated_time}</p>
            </div>
            
            <div class="timeline">
                {email_items}
            </div>
        </div>
        
        <div class="footer">
            <p>© 2025 Outlook Timeline - M365 郵件關鍵字搜尋與時間軸分析工具</p>
        </div>
    </div>
</body>
</html>"""
        
        if not self.emails:
            return html_template.format(
                total_emails=0,
                date_range_days=0,
                keywords_count=0,
                start_date="無",
                end_date="無",
                all_keywords="無",
                generated_time=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                email_items="<p style='text-align: center; color: #666; font-size: 1.2em; margin: 50px 0;'>沒有找到相關郵件</p>"
            )
        
        # 計算統計資訊
        start_date = self.emails[0].date
        end_date = self.emails[-1].date
        date_range_days = (end_date - start_date).days + 1
        
        # 收集所有關鍵字
        all_keywords_set = set()
        for email_info in self.emails:
            all_keywords_set.update(email_info.keywords_found)
        
        # 產生郵件項目HTML
        email_items_html = ""
        for i, email_info in enumerate(self.emails, 1):
            keywords_html = ""
            for keyword in email_info.keywords_found:
                keywords_html += f'<span class="keyword-tag">{keyword}</span>'
            
            email_item_html = f"""
            <div class="email-item">
                <div class="email-header">
                    <div class="email-subject">{self._escape_html(email_info.subject)}</div>
                    <div class="email-meta">
                        <div class="meta-item">
                            <span>📅</span>
                            <span class="date-badge">{email_info.date.strftime('%Y-%m-%d %H:%M:%S')}</span>
                        </div>
                        <div class="meta-item">
                            <span>📁</span>
                            <span class="folder-tag">{email_info.folder}</span>
                        </div>
                        <div class="meta-item">
                            <span>📤</span>
                            <span>{self._escape_html(email_info.sender)}</span>
                        </div>
                        <div class="meta-item">
                            <span>📥</span>
                            <span>{self._escape_html(email_info.receiver)}</span>
                        </div>
                    </div>
                </div>
                <div class="email-body">
                    <div class="keywords">
                        {keywords_html}
                    </div>
                    <div class="email-preview">
                        {self._escape_html(email_info.body[:300])}{'...' if len(email_info.body) > 300 else ''}
                    </div>
                </div>
            </div>
            """
            email_items_html += email_item_html
        
        # 填入模板
        return html_template.format(
            total_emails=len(self.emails),
            date_range_days=date_range_days,
            keywords_count=len(all_keywords_set),
            start_date=start_date.strftime('%Y-%m-%d'),
            end_date=end_date.strftime('%Y-%m-%d'),
            all_keywords=', '.join(sorted(all_keywords_set)),
            generated_time=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            email_items=email_items_html
        )
    
    def _escape_html(self, text: str) -> str:
        """轉義HTML特殊字符"""
        if not text:
            return ""
        return (text.replace('&', '&amp;')
                   .replace('<', '&lt;')
                   .replace('>', '&gt;')
                   .replace('"', '&quot;')
                   .replace("'", '&#x27;'))

def main():
    """主程式"""
    parser = argparse.ArgumentParser(description='Outlook Timeline - M365 郵件關鍵字搜尋與時間軸分析工具')
    parser.add_argument('keywords', nargs='+', help='搜尋關鍵字')
    parser.add_argument('-u', '--username', help='M365 帳號')
    parser.add_argument('-p', '--password', help='密碼 (建議使用應用程式密碼)')
    parser.add_argument('-d', '--days', type=int, default=30, help='搜尋天數 (預設: 30)')
    parser.add_argument('-f', '--folders', nargs='*', help='指定搜尋的資料夾')
    parser.add_argument('-o', '--output', choices=['json', 'csv', 'text', 'html'], default='text', help='輸出格式')
    parser.add_argument('--no-sent', action='store_true', help='不搜尋寄件備份')
    parser.add_argument('--save', help='儲存報告到檔案')
    
    args = parser.parse_args()
    
    # 取得帳號密碼 (優先使用環境變數)
    username = args.username or os.getenv('M365_USERNAME')
    password = args.password or os.getenv('M365_PASSWORD')
    
    # 如果環境變數都沒有設定，才詢問使用者輸入
    if not username:
        username = input("M365 帳號: ")
    if not password:
        password = getpass.getpass("密碼 (建議使用應用程式密碼): ")
    
    # 從環境變數取得預設值
    default_days = int(os.getenv('DEFAULT_DAYS_BACK', args.days))
    default_output = os.getenv('DEFAULT_OUTPUT_FORMAT', args.output)
    
    # 建立分析器
    analyzer = OutlookTimeline(username, password)
    
    try:
        # 連接
        if not analyzer.connect():
            sys.exit(1)
        
        # 搜尋郵件
        emails = analyzer.search_emails(
            keywords=args.keywords,
            folders=args.folders,
            days_back=args.days if args.days != 30 else default_days,
            include_sent=not args.no_sent
        )
        
        if not emails:
            print("沒有找到相關郵件")
            sys.exit(0)
        
        # 產生報告
        output_format = args.output if args.output != 'text' else default_output
        report = analyzer.generate_timeline_report(output_format)
        
        # 輸出或儲存報告
        if args.save:
            with open(args.save, 'w', encoding='utf-8') as f:
                f.write(report)
            print(f"✓ 報告已儲存至 {args.save}")
        else:
            print(report)
            
    except KeyboardInterrupt:
        print("\n中斷執行")
    except Exception as e:
        print(f"✗ 發生錯誤: {e}")
    finally:
        analyzer.disconnect()

if __name__ == "__main__":
    main()