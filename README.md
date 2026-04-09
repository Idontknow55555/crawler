# 司法網站爬蟲 (公開資訊)

## 功能介紹

此Python程式是一個自動化爬蟲工具，用於從台灣司法院司法資料庫網站 (https://judgment.judicial.gov.tw/) 爬取公開的民事判決書資料。程式主要針對"返還墊款"相關的判決書進行搜索和提取，並將結果儲存至Excel檔案和MySQL資料庫。

### 主要功能

1. **自動化瀏覽器操作**：
   - 使用Selenium WebDriver自動控制Chrome瀏覽器
   - 模擬用戶行為進行搜索和頁面導航

2. **條件化搜索**：
   - 案件類別：民事 (V)
   - 關鍵字搜索："返還墊款"
   - 裁判案由："判決"
   - 支援按年度和月份分批搜索

3. **資料提取**：
   - 裁判字號 (標題)
   - 裁判日期
   - 裁判案由
   - 判決書連結
   - 完整內文
   - 主文 (判決主文)
   - 法官姓名

4. **資料清理**：
   - 移除非法字符
   - 清理重複內容
   - 日期格式轉換 (民國年轉西元年)

5. **資料儲存**：
   - 輸出Excel檔案 (`all.xlsx`)
   - 插入MySQL資料庫

### 技術實現

- **核心技術**：Selenium, BeautifulSoup, ChromeDriver
- **資料處理**：pandas, openpyxl
- **資料庫**：MySQLdb
- **自動化**：webdriver-manager

### 使用方法

1. 安裝依賴項：
   ```bash
   pip install selenium beautifulsoup4 pandas openpyxl mysqlclient webdriver-manager
   ```

2. 配置資料庫連線：
   - 在 `crawler.py` 中修改資料庫連線參數：
     - host: 資料庫主機
     - port: 連接埠
     - user: 用戶名
     - passwd: 密碼
     - db: 資料庫名稱

3. 運行程式：
   ```bash
   python crawler.py
   ```

### 注意事項

- 此程式僅用於爬取公開資訊，請遵守網站使用條款
- 請勿過度頻繁訪問網站，以免造成伺服器負擔
- 確保Chrome瀏覽器已安裝
- 資料庫表結構需預先建立

### 輸出檔案

- `all.xlsx`：包含所有提取的判決書資料的Excel檔案
