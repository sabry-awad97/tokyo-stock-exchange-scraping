import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

BASE_URL = "https://www2.jpx.co.jp/tseHpFront/JJK020010Action.do"


class ChromeBrowser:
    def __init__(self):
        self.service = Service(ChromeDriverManager().install())
        self.options = Options()
        self.browser = webdriver.Chrome(
            service=self.service, options=self.options)
        self.wait = WebDriverWait(self.browser, 60)
        self.browser.maximize_window()

    def close(self):
        self.browser.close()


class WebPage(ChromeBrowser):
    def __init__(self, url):
        super().__init__()
        self.browser.get(url)

    def read_table(self, path):
        table = self.wait.until(
            EC.presence_of_element_located((By.XPATH, path)))
        html_content = table.get_attribute("outerHTML")
        df, = pd.read_html(html_content)
        return df

    def get_element(self, path):
        return self.wait.until(EC.presence_of_element_located((By.XPATH, path)))

    def click_element(self, path):
        self.get_element(path).click()

    def select_value_from_dropdown(self, path, value):
        Select(self.get_element(path)).select_by_value(value)


class TokyoStockExchangePage(WebPage):
    all_basic_info = pd.DataFrame()
    page_number = 1
    max_pages = 10

    def __init__(self):
        super().__init__(BASE_URL)

    def select_200(self):
        self.select_value_from_dropdown(
            '//*[@id="bodycontents"]/div[2]/form/div[1]/table/tbody/tr[1]/td/span/select', '200')

    def click_prime_checkbox(self):
        self.click_element(
            '//*[@id="bodycontents"]/div[2]/form/div[1]/table/tbody/tr[4]/td/span/span/label[1]/input')

    def click_search_button(self):
        self.click_element('//*[@id="bodycontents"]/div[2]/form/p/input')

    def navigate_to_next_page(self):
        self.click_element(
            '//*[@id="bodycontents"]/div[2]/form/div[1]/div[2]/a')
        self.page_number += 1

    def scrape_basic_information(self, row_number):
        # click basic information button
        self.click_element(
            f'//*[@id="bodycontents"]/div[2]/form/table/tbody/tr[{row_number + 3}]/td[7]/input')
        basic_info_df = self.read_table(
            '//*[@id="body_basicInformation"]/div/table[1]')
        self.all_basic_info = pd.concat(
            [self.all_basic_info, basic_info_df], ignore_index=True)
        self.all_basic_info.to_excel("basic_info.xlsx", index=False)

        # click return button
        self.click_element(
            '//*[@id="body_basicInformation"]/div/table[5]/tbody/tr/td/input')

    def scrape_all_basic_information(self, tokyo_stock_exchange):
        for i in range(len(tokyo_stock_exchange)):
            self.scrape_basic_information(i)

        self.all_basic_info.to_excel(
            f"basic_info_page {self.page_number}.xlsx", index=False)

    def scrape_data(self):
        self.select_200()
        self.click_prime_checkbox()
        self.click_search_button()

        while self.page_number <= self.max_pages:
            tokyo_stock_exchange = self.read_table(
                '//*[@id="bodycontents"]/div[2]/form/table')
            tokyo_stock_exchange.to_excel(f"tokyo_stock_ecxhange {self.page_number}.xlsx",
                                          index=False)
            self.scrape_all_basic_information(tokyo_stock_exchange)

            if self.page_number == self.max_pages:
                continue

            self.navigate_to_next_page()

    def run(self):
        self.scrape_data()
        self.browser.close()


tse = TokyoStockExchangePage()

tse.run()
