import pandas as pd
from bs4 import BeautifulSoup
import requests
import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from selenium.webdriver.support import expected_conditions as ec


# 取得目標公司名稱
def get_stock_names(data_file="data .xlsx"):
    base_path = os.path.dirname(__file__)
    data_file_path = os.path.join(base_path, data_file)

    if not os.path.exists(data_file_path):
        raise FileNotFoundError(f"文件 {data_file_path} 不存在")

    try:
        df1 = pd.read_excel(data_file_path, sheet_name="股票股利標的", usecols=["股票名稱"])
        df2 = pd.read_excel(data_file_path, sheet_name="高現金殖利率標的", usecols=["股票名稱"])
        combined_df = pd.concat([df1, df2])
        stock_names = combined_df["股票名稱"].tolist()
        return stock_names
    except Exception as e:
        raise ValueError(f"讀取文件失敗: {e}")


# 抓取MOP本日之網頁內容
def fetch_mops_today_page(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        return BeautifulSoup(response.text, "html.parser")
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"抓取資料失敗: {e}")


# 抓取MOP前一日之網頁內容
def fetch_mops_yesterday_page(url):
    current_directory = os.path.dirname(__file__)
    driver_path = os.path.join(current_directory, "chromedriver.exe")
    service = Service(executable_path=driver_path)
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    driver = webdriver.Chrome(service=service, options=options)

    try:
        driver.get(url)
        yesterday_button = driver.find_element(By.XPATH, "//input[@value='前一日資料（市場別：全體公司）']")
        yesterday_button.click()

        WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((By.XPATH, "//input[@value='前一日資料']"))
        )
        soup = BeautifulSoup(driver.page_source, "html.parser")
    except NoSuchElementException as e:
        driver.quit()
        raise NoSuchElementException(f"Element not found: {e}")
    except TimeoutException as e:
        driver.quit()
        raise TimeoutException(f"Loading timeout: {e}")
    except WebDriverException as e:
        driver.quit()
        raise WebDriverException(f"WebDriver error: {e}")
    finally:
        driver.quit()

    return soup


# 解析HTML 並回傳
def parse_mops_data(soup, stock_names):
    table_rows = soup.find_all("tr")
    result = []
    for row in table_rows:
        cells = row.find_all("td")
        if cells and len(cells) > 5:
            stock_short_name = cells[1].text.strip()
            if stock_short_name in stock_names:
                stock_code = cells[0].text.strip()
                stock_date = cells[2].text.strip()
                stock_time = cells[3].text.strip()
                stock_subject = cells[4].text.strip().replace("\r\n", "")
                result.append((stock_code, stock_short_name, stock_date, stock_time, stock_subject))
    return result


def parse_mops_yesterday_data(soup, stock_names):
    table_rows = soup.find_all("tr", class_="odd")
    result = []
    for row in table_rows:
        cells = row.find_all("td")
        if cells and len(cells) > 4:
            stock_date = cells[0].text.strip()
            stock_time = cells[1].text.strip()
            stock_code = cells[2].text.strip()
            stock_short_name = cells[3].text.strip()
            stock_subject = cells[4].text.strip().replace("\r\n", "")

            stock_subject = " ".join(stock_subject.split())

            if stock_short_name in stock_names:
                result.append((stock_code, stock_short_name, stock_date, stock_time, stock_subject))
    return result

# if __name__ == "__main__":
