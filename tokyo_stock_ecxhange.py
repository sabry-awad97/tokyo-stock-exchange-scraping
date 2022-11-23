import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

BASE_URL = "https://www2.jpx.co.jp/tseHpFront/JJK020010Action.do"

s = Service(ChromeDriverManager().install())
chrome_options = Options()
browser = webdriver.Chrome(service=s,
                           options=chrome_options)
wait = WebDriverWait(browser, 60)
browser.maximize_window()
browser.get(BASE_URL)

selectElement = wait.until(EC.presence_of_element_located((By.XPATH,
                                                           '//*[@id="bodycontents"]/div[2]/form/div[1]/table/tbody/tr[1]/td/span/select')))

select = Select(selectElement)
select.select_by_value('200')

prime_checkbox = wait.until(EC.presence_of_element_located((By.XPATH,
                                                            '//*[@id="bodycontents"]/div[2]/form/div[1]/table/tbody/tr[4]/td/span/span/label[1]/input')))

prime_checkbox.click()

search_button = wait.until(EC.presence_of_element_located((By.XPATH,
                                                           '//*[@id="bodycontents"]/div[2]/form/p/input')))

search_button.click()

all_data = pd.DataFrame()

page_number = 1
while page_number <= 10:
    table = wait.until(EC.presence_of_element_located((By.XPATH,
                                                       '//*[@id="bodycontents"]/div[2]/form/table')))

    html_content = table.get_attribute("outerHTML")
    df, = pd.read_html(html_content)

    all_data = pd.concat([all_data, df], ignore_index=True)

    if page_number == 10:
        continue

    next_button = wait.until(EC.presence_of_element_located((By.XPATH,
                                                             '//*[@id="bodycontents"]/div[2]/form/div[1]/div[2]/a')))

    next_button.click()

    page_number += 1

all_data.to_excel("tokyo_stock_ecxhange.xlsx", index=False)

browser.close()
