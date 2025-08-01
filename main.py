import pandas as pd
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium_driverless.sync import webdriver as webdriverless
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, NoSuchElementException
from selenium.webdriver.common.proxy import Proxy, ProxyType
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
import random
import requests
import os
from dotenv import load_dotenv
import json
import re
import traceback
import sys

class Scraper:
    def __init__(self):
        load_dotenv()
        self.surname_number = os.getenv('SURNAME_NUMBER')
        self.surnames = self.get_surnames()
        self.results = []
        self.driver = self.get_driver()
        self.wait = WebDriverWait(self.driver, 10)
        self.index_comune = 1
    def get_driver(self):
        driver = webdriver.Chrome()
        driver.maximize_window()
        return driver
    def get_chunk_surnames_filepath(self):
        chunk_file = f'./data/surnames_{self.surname_number}.xlsx'
        print(f'Use Surnames Chunk File: {chunk_file}')

        return chunk_file
    def get_surnames(self):
        surnames = []
        chunk_file = self.get_chunk_surnames_filepath()
        workbook = openpyxl.load_workbook(chunk_file)
        sheet = workbook.active
        for row in sheet.iter_rows(min_row=2, values_only=True):
            surnames.append(row[0])

        return surnames
    def delay(self, min=1, max=2):
        time.sleep(random.uniform(min, max))
    def export_results_excel(self):
        df = pd.DataFrame(self.results)
        df.to_excel(f'results/results_{self.surname_number}_{str(int(time.time()))}.xlsx', index=False)
    def start_get_data(self):
        search_url = "https://portale.fnomceo.it/cerca-prof/index.php"
        self.driver.get(search_url)

        for surname in self.surnames:
            self.wait.until(EC.visibility_of_element_located((By.ID, "cognomeID")))
            surname_text = self.driver.find_element(By.ID, "cognomeID")
            surname_text.send_keys(surname)
            submit_button = self.driver.find_element(By.ID, "submitButtonID")
            submit_button.click()
            break






scraper = Scraper()
try:
    scraper.start_get_data()
except Exception as e:
    print(f"An error occurred: {e}")
finally:
    scraper.delay(60,70)
    scraper.export_results_excel()