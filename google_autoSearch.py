import selenium
import openpyxl
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

search_string = input("検索語句は？: ")
search_number = input("検索数は？: ")
search_number_int = int(search_number)

#seleniumを使うための設定とGoogleの画面への遷移
INTERVAL = 2.5
URL = "https://www.google.com"
options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--lang=jp')
driver = webdriver.Chrome('chromedriver', options=options)
driver.maximize_window()
time.sleep(INTERVAL)
driver.get(URL)
time.sleep(INTERVAL)

#文字を入力して検索

query = driver.find_element(By.NAME, 'q')
print(query.get_attribute('outerHTML'))
query.click()
query.clear()
query.send_keys(search_string)
query.send_keys(Keys.RETURN)
print("Searching for: " + search_string)
time.sleep(INTERVAL)

#検索結果の一覧を取得する
results = []
flag = False
while True:
    g_ary = driver.find_elements(By.CLASS_NAME,'g')
    for g in g_ary:
        result = {}
        try:
          result['url'] = g.find_element(By.CLASS_NAME,'yuRUbf').find_element(By.TAG_NAME,'a').get_attribute('href')
          result['title'] = g.find_element(By.TAG_NAME,'h3').text
          results.append(result)
          if len(results) >= search_number_int:
              flag = True
              break
        except Exception as e:
          print(e)
          print(results)
          print(len(results))
    if flag:
      break
    driver.find_element(By.ID,'pnnext').click()
    time.sleep(INTERVAL)

#ワークブックの作成とヘッダー入力
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet['A1'].value = 'マーク'
sheet['B1'].value = 'タイトル'
sheet['C1'].value = 'URL'
sheet.column_dimensions['A'].width = 7
sheet.column_dimensions['B'].width = 60
sheet.column_dimensions['C'].width = 200
sheet.column_dimensions['D'].hidden= True

#シートにタイトルとURLの書き込み
for row, result in enumerate(results,2):
    sheet[f"B{row}"] = result['title']
    sheet[f"C{row}"].value = f"=HYPERLINK(D{row})"
    sheet[f"D{row}"] = result['url']

workbook.save(f"google_search_{search_string}.xlsx")
driver.close()