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


def read_table(path: str):
    table = wait.until(EC.presence_of_element_located((By.XPATH, path)))
    html_content = table.get_attribute("outerHTML")
    df, = pd.read_html(html_content)
    return df


def get_button(path: str):
    return wait.until(EC.presence_of_element_located((By.XPATH, path)))


def click_button(path: str):
    get_button(path).click()


# select 200
Select(get_button(
    '//*[@id="bodycontents"]/div[2]/form/div[1]/table/tbody/tr[1]/td/span/select')).select_by_value('200')


# click prime checkbox
click_button(
    '//*[@id="bodycontents"]/div[2]/form/div[1]/table/tbody/tr[4]/td/span/span/label[1]/input')


# click search button
click_button('//*[@id="bodycontents"]/div[2]/form/p/input')

page_number = 1
while page_number <= 10:
    tokyo_stock_exchange = read_table(
        '//*[@id="bodycontents"]/div[2]/form/table')

    tokyo_stock_exchange.to_excel(f"tokyo_stock_ecxhange {page_number}.xlsx",
                                  index=False)

    all_basic_info = pd.DataFrame()
    for i in range(len(tokyo_stock_exchange)):
        # click basic information button
        click_button(
            f'//*[@id="bodycontents"]/div[2]/form/table/tbody/tr[{i + 3}]/td[7]/input')

        basic_info_df = read_table(
            '//*[@id="body_basicInformation"]/div/table[1]')

        all_basic_info = pd.concat([all_basic_info, basic_info_df],
                                   ignore_index=True)

        all_basic_info.to_excel("basic_info.xlsx", index=False)

        # click return button
        click_button(
            '//*[@id="body_basicInformation"]/div/table[5]/tbody/tr/td/input')

    all_basic_info.to_excel(f"basic_info_page {page_number}.xlsx", index=False)

    if page_number == 10:
        continue

    click_button('//*[@id="bodycontents"]/div[2]/form/div[1]/div[2]/a')
    page_number += 1

browser.close()
