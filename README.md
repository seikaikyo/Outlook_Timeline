# Outlook Timeline - M365 郵件關鍵字搜尋與時間軸分析工具

## 功能特色

- 透過 IMAP 連接 M365 Outlook 郵件
- 多關鍵字搜尋功能
- 時間軸排序與事件分析
- 支援多種輸出格式 (JSON、CSV、文字)
- 可搜尋多個郵件資料夾
- 內建危機事件關鍵字預設

## 系統需求

- Python 3.7 或更新版本
- M365 帳號 (需啟用 IMAP)
- 應用程式密碼 (建議使用)

## 安裝步驟

### 本地部署

1. 下載專案檔案
2. 安裝相依套件：
   ```bash
   pip install -r requirements.txt
   ```
3. 設定環境變數：
   ```bash
   cp .env .env.local  # 複製環境變數檔案 (可選)
   ```
4. 啟動網頁應用程式：
   ```bash
   streamlit run app.py
   ```

### Streamlit Cloud 部署

1. 將專案上傳到 GitHub
2. 連接到 [Streamlit Cloud](https://streamlit.io/cloud)
3. 在 Secrets 設定中添加：
   ```toml
   [outlook]
   M365_USERNAME = "your_email@company.com"
   M365_PASSWORD = "your_app_password"
   ```

## M365 IMAP 設定

### 1. 啟用 IMAP
1. 登入 [Outlook 網頁版](https://outlook.office365.com)
2. 點選右上角設定圖示 → 檢視所有 Outlook 設定
3. 選擇「郵件」→「同步處理電子郵件」
4. 啟用「IMAP」選項

### 2. 建立應用程式密碼 (建議)
1. 登入 [Microsoft 帳戶安全性](https://account.microsoft.com/security)
2. 選擇「進階安全性選項」
3. 在「應用程式密碼」區塊選擇「建立新的應用程式密碼」
4. 輸入應用程式名稱 (例如：Outlook Timeline)
5. 記下產生的密碼

## 環境變數設定

在 `.env` 檔案中設定您的 M365 帳號資訊：

```env
# M365 Outlook 帳號設定
M365_USERNAME=your_email@company.com
M365_PASSWORD=your_app_password

# IMAP 伺服器設定
IMAP_SERVER=outlook.office365.com
IMAP_PORT=993

# 預設搜尋設定
DEFAULT_DAYS_BACK=30
DEFAULT_OUTPUT_FORMAT=text
```

## 使用方法

### 🌐 網頁版本 (推薦)

啟動 Streamlit 網頁應用程式：
```bash
streamlit run app.py
```

特色功能：
- 視覺化介面，無需命令列操作
- 即時圖表分析 (時間分布、關鍵字統計、資料夾分布)
- 互動式郵件瀏覽
- 多種匯出格式 (CSV、JSON、HTML)
- 響應式設計，支援手機瀏覽

### 📱 命令列版本

#### 基本用法 (使用環境變數)
```bash
python outlook_timeline.py 關鍵字1 關鍵字2
```

#### 基本用法 (手動輸入帳號)
```bash
python outlook_timeline.py 關鍵字1 關鍵字2 -u your_email@company.com
```

### 進階參數
```bash
python outlook_timeline.py 緊急 危機 問題 \
  --days 60 \
  --folders "INBOX" "Sent Items" "重要郵件" \
  --output json \
  --save report.json
```

**注意**：如果已設定環境變數，就不需要在命令中指定 `--username` 和 `--password`

### 參數說明
- `keywords`: 搜尋關鍵字 (必要)
- `-u, --username`: M365 帳號 (可用環境變數 M365_USERNAME)
- `-p, --password`: 密碼或應用程式密碼 (可用環境變數 M365_PASSWORD)
- `-d, --days`: 搜尋天數 (預設使用環境變數 DEFAULT_DAYS_BACK 或 30)
- `-f, --folders`: 指定搜尋資料夾
- `-o, --output`: 輸出格式 (text/json/csv/html，預設使用環境變數 DEFAULT_OUTPUT_FORMAT 或 text)
- `--no-sent`: 不搜尋寄件備份
- `--save`: 儲存報告到檔案

## 使用範例

### 1. 危機事件追蹤
```bash
python outlook_timeline.py 緊急 危機 事故 異常 --days 90 --output csv --save 危機事件報告.csv
```

### 2. 專案進度追蹤
```bash
python outlook_timeline.py 專案A 里程碑 進度 截止 --folders "INBOX" "專案郵件" --days 60
```

### 3. 安全事件分析
```bash
python outlook_timeline.py 資安 安全 入侵 威脅 --output json --save 安全事件.json
```

### 4. 產生 HTML 視覺化報告
```bash
python outlook_timeline.py 緊急 危機 問題 --output html --save 事件報告.html
```

## 輸出格式說明

### 文字格式 (預設)
- 易於閱讀的時間軸格式
- 包含完整郵件資訊
- 適合快速瀏覽

### JSON 格式
- 結構化資料
- 適合程式處理
- 包含完整中繼資料

### CSV 格式
- 表格化資料
- 適合 Excel 分析
- 便於統計處理

### HTML 格式
- 視覺化時間軸介面
- 響應式網頁設計
- 包含統計資訊圖表
- 關鍵字標籤顯示
- 支援行動裝置瀏覽

## 安全性建議

1. 使用應用程式密碼而非主要密碼
2. 定期更換應用程式密碼
3. 不要在腳本中硬編碼密碼
4. 限制搜尋時間範圍以提升效能

## 常見問題

### Q: 無法連線到伺服器
A: 請確認：
- M365 帳號已啟用 IMAP
- 使用正確的應用程式密碼
- 網路連線正常

### Q: 找不到郵件
A: 請檢查：
- 關鍵字是否正確
- 搜尋時間範圍是否足夠
- 郵件是否在指定資料夾內

### Q: 效能較慢
A: 建議：
- 縮小搜尋時間範圍
- 限制搜尋資料夾數量
- 使用更精確的關鍵字

## 技術支援

如有技術問題或建議，請提出 Issue。

## 授權條款

本專案為教育用途，請遵守 M365 服務條款。