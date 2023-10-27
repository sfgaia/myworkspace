from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.select import Select
import openpyxl
import time


def perform_actions(file_path):
    # Excelファイルの読み込み
    wb = openpyxl.load_workbook(file_path)
    sheet = wb['config']

    # WebDriverの初期化
    driver =webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

    # 処理開始（2行目から）
    for row in sheet.iter_rows(min_row=2, values_only=True):
        action_type = row[0]
        xpath = row[1]
        input_type = row[2]
        value = row[3]
        print(xpath)

        if action_type == "遷移":
            driver.get(value)
        elif action_type == "待機":
            wait_time = float(value)
            #driver.implicitly_wait(wait_time)
            time.sleep(wait_time)
        elif action_type == "入力":
            element = driver.find_element(By.XPATH, xpath)
            if input_type == "テキストボックス" or input_type == "テキストエリア":
                print(value)
                element.send_keys(value)
            elif input_type == "ラジオボタン" or input_type == "チェックボックス":
                element.click()
            elif input_type == "プルダウン":
                dropdown = Select(element)
                dropdown.select_by_visible_text(value)
        elif action_type == "押下":
            element = driver.find_element(By.XPATH, xpath)
            element.click()

    # 終了処理
    # driver.quit()


# 実行例
file_path = r"C:\myworkspace\python\browser_automation\excel\input.xlsx"
perform_actions(file_path)
