import openpyxl
import re
import pandas as pd
import MySQLdb
from bs4 import BeautifulSoup
from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager



# 示例數據
all_titles = []
all_dates = []
all_reasons = []
all_contents = []
all_main_texts = []
all_judge_names = []
all_links = []


# 移除非法字符的函数
def remove_illegal_characters(value):
    ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')
    return ILLEGAL_CHARACTERS_RE.sub("", value)

# 清理內容的函數，移除特定的模式
def clean_content(content):
    patterns_to_remove = [
        r'去格式引用分享網址名詞查詢名詞收集裁判易讀小幫手友善列印轉存PDF分享P分享網址：若您有連結此資料內容之需求，請直接複製下述網址請選取上方網址後，按 Ctrl\+C 或按滑鼠右鍵選取複製，即可複製網址。'
    ]
    for pattern in patterns_to_remove:
        content = re.sub(pattern, '', content, flags=re.DOTALL)
    return content

# 將民國日期轉換為西元日期的函數
def convert_date(roc_date):
    try:
        year, month, day = map(int, roc_date.split('.'))
        year += 1911  # 轉換為西元年
        return f"{year:04d}-{month:02d}-{day:02d}"
    except:
        return None  # 如果轉換失敗，返回 None

# 從詳細頁面提取數據的函數
# 更新提取內文的部分
# 更新提取內文的部分，保留判決書形式的換行
def extract_detail_data(driver, link):
    content = ''
    judge_names = set()
    main_text = ''

    try:
        # 在新標籤頁中打開鏈接
        driver.execute_script("window.open('{}');".format(link))
        driver.switch_to.window(driver.window_handles[1])
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, 'htmlcontent')))

        # 使用BeautifulSoup解析詳細頁面的內容
        soup = BeautifulSoup(driver.page_source, 'html.parser')

        # 專注於 class 為 'htmlcontent' 的區塊
        htmlcontent_div = soup.find('div', {'class': 'htmlcontent'})
        if htmlcontent_div:
            # 找到所有的 <div> 標籤
            content_tags = htmlcontent_div.find_all(['p', 'div', 'span'])
            for tag in content_tags:
                tag_text = tag.get_text(strip=True)
                if tag_text and tag_text not in content:
                    content += tag_text + '\n'

        # 清理内容（删除多余的換行符）
        content = re.sub(r'\n+', '\n', content)  # 删除多餘的換行符
        content = content.strip()  # 移除開頭和結尾的空白字符

        # 提取主文
        jud_content_div = soup.find('div', {'class': 'jud_content'})
        if jud_content_div:
            main_text_element = jud_content_div.find('abbr', {'id': '%e4%b8%bb%e6%96%87'})
            if main_text_element:
                main_text_lines = []
                current_element = main_text_element.parent
                while current_element:
                    if current_element.name == 'abbr' and current_element.get('id') == '%e5%81%87%e5%9f%b7%e8%a1%8c':
                        break
                    text = current_element.get_text(strip=True)
                    if text:
                        main_text_lines.append(text)
                    current_element = current_element.find_next_sibling()
                main_text = '\n'.join(main_text_lines)
            else:
                print(f"找不到主文在链接: {link}")

        # 清除主文下方的事實及理由資料
        if main_text:
            index = main_text.find('事實及理由' or '事  實' or '事  實  及  理  由' or '理　　　由' or'理  由' or '理      由'or'事  實  及  理  由' or'事 實 及 理 由' or '理 由')
            if index != -1:
                main_text = main_text[:index]

        # 提取法官名稱和後面的內容
        judge_elements = soup.find_all(text=re.compile(r'法\s*官'))
        if judge_elements:
            for elem in judge_elements:
                parent = elem.find_parent()
                if parent:
                    judge_info = parent.get_text(strip=True)
                    match = re.findall(r'法\s*官\s*([\u4e00-\u9fa5])\s*([\u4e00-\u9fa5])\s*([\u4e00-\u9fa5])',
                                       judge_info)
                    match = ["".join(judge_name) for judge_name in match]
                    if match:
                        judge_names.update(match)
        else:
            print(f"找不到法官姓名在鏈接: {link}")

        judge_name = ', '.join(judge_names)

        # 關閉當前標籤頁並切回主頁面
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
    except Exception as e:
        print(f"提取詳細資料時發生錯誤: {e}")
        driver.close()
        driver.switch_to.window(driver.window_handles[0])

    return content, judge_name, main_text


# 初始化一個集合來跟踪唯一鏈接
seen_links = set()
seen_contents = set()

# 從搜索結果頁面提取數據的函數
def extract_data(driver):
    try:
        while True:
            soup = BeautifulSoup(driver.page_source, 'html.parser')

            #檢查是否超過500條結果
            over_500_message = soup.find('h3', string='查詢結果超出 500 筆')
            if over_500_message:
                print("查詢結果超出 500 筆，請縮小查詢條件。")
                return

            # 查找判決書資料表
            table = soup.find('table', id='jud')
            if not table:
                print("找不到判決書資料表。")
                return

            # 遍歷表格行並提取數據
            rows = table.find_all('tr')
            for row in rows[1:]:  # 從第二行開始遍歷，跳過表頭
                cells = row.find_all('td')
                if len(cells) >= 4:  # 確保每一行都有足夠的欄位
                    title_cell = cells[1]
                    title_link = title_cell.find('a')
                    title = title_link.text
                    link = title_link.get('href')

                    date = cells[2].text
                    reason = cells[3].text

                    all_titles.append(title)
                    all_dates.append(date)
                    all_reasons.append(reason)
                    all_links.append(link)

                    # 提取詳細資料
                    content, judge_name, main_text = extract_detail_data(driver, link)
                    all_contents.append(content)
                    all_judge_names.append(judge_name)
                    all_main_texts.append(main_text)

            #檢查是否有下一頁
            try:
                iframe = driver.find_element(By.CSS_SELECTOR, 'iframe#iframe-data')
                driver.switch_to.frame(iframe)

                next_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, 'hlNext'))
                )
                if 'disabled' in next_button.get_attribute('class'):
                    break  # No more pages
                else:
                    next_button.click()
                    sleep(3)  # 等待頁面加載
            except Exception as e:
                print(f"找不到下一頁按鈕或其他錯誤: {e}")
                break

    except Exception as e:
        print(f"提取資料時發生錯誤: {e}")


# 將數據寫入Excel文件的函數
def write_to_excel():
    wb = openpyxl.Workbook()
    ws = wb.active

    # 寫入表頭
    ws.append(['裁判字號', '裁判日期', '裁判案由', '連結', '內文', '主文', '法官姓名'])  # 添加主文標題

    # 寫入資料
    for i in range(len(all_titles)):
        ws.append([
            remove_illegal_characters(all_titles[i]),
            remove_illegal_characters(all_dates[i]),
            remove_illegal_characters(all_reasons[i]),
            remove_illegal_characters(all_links[i]),
            remove_illegal_characters(all_contents[i]),
            remove_illegal_characters(all_judge_names[i]),
            remove_illegal_characters(all_main_texts[i]),  # 添加主文資料
        ])

    # 儲存Excel檔案
    wb.save('all.xlsx')

# 將數據插入SQL數據庫的函數
def insert_data_to_sql(all_titles, all_dates, all_reasons, all_contents, all_main_texts, all_judge_names):
    db = None
    cursor = None
    try:
        # 連接到MySQL數據庫
        db = MySQLdb.connect(
            host="163.13.201.96",
            port=3307,
            user="root",
            passwd="JudgementAnalysis_0000",
            db="判決書分析比對系統_db"
        )
        print("成功連接到資料庫。")

        cursor = db.cursor()

        # 獲取當前最大ID
        cursor.execute("SELECT MAX(id) FROM `Table`")
        max_id = cursor.fetchone()[0] or 0
        print(f"目前最大 id: {max_id}")

        max_content_length = 8000

        # 檢查輸入數據的有效性
        if not (all_titles and all_dates and all_reasons and all_contents and all_main_texts and all_judge_names):
            print("有一個或多個輸入列表為空。")
            return

        if not (len(all_titles) == len(all_dates) == len(all_reasons) == len(all_contents) == len(all_main_texts) == len(all_judge_names)):
            print("輸入列表長度不一致。")
            return

        # 準備數據
        df = pd.DataFrame({
            'id': [max_id + i + 1 for i in range(len(all_titles))],
            '裁判字號': [remove_illegal_characters(title) if title else None for title in all_titles],
            '裁判日期': [convert_date(remove_illegal_characters(date)) if date else None for date in all_dates],
            '裁判案由': [remove_illegal_characters(reason) if reason else None for reason in all_reasons],
            '內文': [remove_illegal_characters(content)[:max_content_length] if content else None for content in all_contents],
            '主文': [remove_illegal_characters(main_text) if main_text else None for main_text in all_main_texts],
            '法官姓名': [remove_illegal_characters(judge_name) if judge_name else None for judge_name in all_judge_names]
        })

        # 準備SQL插入語句
        sql = """
        INSERT INTO `Table` (
            id,
            裁判字號, 
            裁判日期, 
            裁判案由, 
            內文, 
            主文,
            法官姓名
        ) VALUES (%s, %s, %s, %s, %s, %s, %s)
        """
        values = df.values.tolist()

        # 執行批量插入
        cursor.executemany(sql, values)
        db.commit()
        print("數據成功提交到資料庫。")

    except MySQLdb.Error as e:
        print(f"連接 MySQL 平台時出錯: {e}")

    finally:
        if cursor:
            cursor.close()
        if db:
            db.close()


# 插入數據到 SQL 資料庫
try:
    insert_data_to_sql(all_titles, all_dates, all_reasons, all_contents, all_main_texts, all_judge_names)
    print("資料成功寫入 MySQL 資料庫")
except Exception as e:
    print(f"寫入 MySQL 資料庫時發生錯誤: {e}")

service = Service(ChromeDriverManager().install())  # 替換為你的chromedriver路徑
driver = webdriver.Chrome(service=service)
driver.implicitly_wait(30)

startYear = 113


def conSearch(year):
    for month in range(1, 4):
        driver.get('https://judgment.judicial.gov.tw/FJUD/default.aspx')

        # 更多條件查詢
        driver.find_element(By.CSS_SELECTOR, 'a.btn.btn-warning').click()

        # 等待一下
        sleep(1)

        # 條件選取
        driver.find_element(By.CSS_SELECTOR, 'button[type="reset"]').click()  # 重置搜尋條件

        # 設置搜索條件
        driver.find_element(By.CSS_SELECTOR, 'input[value="V"]').click()  # 案件類別
        sleep(1)
        driver.find_element(By.CSS_SELECTOR, 'input#dy1').send_keys(year)  # 裁判期間－起年
        driver.find_element(By.CSS_SELECTOR, 'input#dm1').send_keys(month)  # 裁判期間－起月
        driver.find_element(By.CSS_SELECTOR, 'input#dd1').send_keys('1')  # 裁判期間－起日
        driver.find_element(By.CSS_SELECTOR, 'input#dy2').send_keys(year)  # 裁判期間－迄年
        driver.find_element(By.CSS_SELECTOR, 'input#dm2').send_keys(month)  # 裁判期間－迄月
        driver.find_element(By.CSS_SELECTOR, 'input#dd2').send_keys('31')  # 裁判期間－迄日
        sleep(1)
        driver.find_element(By.CSS_SELECTOR, 'input#jud_title').send_keys('返還墊款')
        # 清償借款、給付借款、返還借款、損害賠償、遷讓房屋、清償債務、分割共有物、侵權行為損害賠償、分割遺產、損害賠償(交通)、給付工程款、給付電信費
        # 返還不當得利、給付工資、侵權行為損害賠償(交通)、給付資遣費、返還信用卡消費款、返還消費借貸款、請求分割共有物、請求損害賠償、給付買賣價金
        # 給付貨款、給付承攬報酬、返還價金、請求給付資遣費、請求遷讓房屋、給付違約金、給付退休金差額、撤銷遺產分割登記、代位分割遺產
        # 請求侵權行為損害賠償、給付簽帳卡消費款、給付薪資、請求給付薪資、給付報酬、請求給付工資、返還土地、返還買賣價金、請求清償借款、返還所有物
        # 債務不履行損害賠償、給付保險金、給付租金、給付扣押款、給付分期買賣價金、請求分割遺產、返還代墊款、返還租賃物、返還房屋
        # 給付委任報酬、給付職業災害補償、請求給付職業災害補償、給付服務報酬、給付管理費
        # 履行遺產分割協議、返還股份、返還牌照、請求返還不當得利、拆除地上物返還土地、返還借名登記土地、返還不動產、返還押租金、國家損害賠償、請求職業災害損害賠償
        # 侵權配偶權損害賠償、遷讓房地、給付差額價金、給付應分擔款、給付薪資債權
        # 給付代收款項、給付預告工資、給付醫療費用、給付電費、給付必要費用、給付醫療費、給付信用卡消費款、給付加班費、給付票款
        # 給付款項、返還報酬、返還履約保證金、返還工程款
        # 請求返還買賣價金、返還本票、返還墊付款、返還服務費、返還金錢、返還不動產所有權、返還承攬報酬、返還款項、請求返還押租保證金
        # 請求返還土地、請求返還價金、返還投資金、返還獎金、返還房地、返還保管款、返還就學貸款、返還墊款
        sleep(1)
        driver.find_element(By.CSS_SELECTOR, 'input#jud_jmain').send_keys('判決')  # 裁判案由
        sleep(1)
        driver.switch_to.default_content()
        sleep(1)
        driver.find_element(By.CSS_SELECTOR, 'input.btn.btn-success').click()  # 送出搜尋條件

        print(f"目前搜尋月份 : {month}")
        # 等待一下
        sleep(1)

        # 切換到 iframe
        iframe = driver.find_element(By.CSS_SELECTOR, 'iframe#iframe-data')
        driver.switch_to.frame(iframe)

        # 提取資料
        extract_data(driver)

        driver.switch_to.default_content()


def conSearch_year():
    # 爬取個年度資料
    countYear = 0
    for d in range(1):
        year = startYear - d
        countYear += 1
        print(f"開始搜尋{year}年度")
        conSearch(year)
    print(f"所有年度搜尋完畢，共有{countYear}年的資料")


# 開始爬取資料
conSearch_year()

# 保存Excel工作簿並關閉瀏覽器
write_to_excel()
insert_data_to_sql(all_titles, all_dates, all_reasons, all_contents, all_main_texts, all_judge_names)
driver.quit()
