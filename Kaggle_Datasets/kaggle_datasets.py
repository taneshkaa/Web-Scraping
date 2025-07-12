from playwright.sync_api import Playwright, sync_playwright, expect
import re
import os
import multiprocessing
import pandas as pd
import datetime as dt, time
from time import sleep, perf_counter
from openpyxl.workbook import Workbook

url = 'https://www.kaggle.com/'

# Dataset Scraping, (can be changed to Competitions, Models, etc.)
navi = url + 'datasets'
data_tt = []
data_ll = []


dt_tm = dt.datetime.fromtimestamp(time.time())
dt_stamp = dt_tm.strftime("%d-%B-%Y")

def run(playwright: Playwright) -> None:
    # Browser [Chrome, Firefox, WebKit]
    browser = playwright.chromium.launch(headless=False)
    
    context = browser.new_context()
    page = context.new_page()
    page.goto(url)
    page.goto(navi)
    sleep(1)

    # Browsing Kaggle - Dataset Section (Aiming for Trending Datasets)
    page.locator('//h2[text()="Trending Datasets"]').scroll_into_view_if_needed()
    sleep(0.5)

    # Expand All
    page.locator('(//h2[text()="Trending Datasets"]/following::button/span[text()="See All"])[1]').click()
    sleep(1)

    list_cnt = page.locator('(//h2[text()="Trending Datasets"]/parent::div//following::ul)[1]//li').count()
    # print(list_cnt)
    int_cnt = int(list_cnt)

    for x in range(int_cnt):
        dt_name = page.locator('(//h2[text()="Trending Datasets"]/parent::div//following::ul)[1]//li//a[@role="link"]').nth(x).get_attribute('aria-label')
        data_tt.append(dt_name)
        sleep(0.5)
        dt_link = page.locator('(//h2[text()="Trending Datasets"]/parent::div//following::ul)[1]//li//a[@role="link"]').nth(x).get_attribute('href')
        # print(dt_link)
        data_ll.append(dt_link)
        sleep(1)

    my_dic = {
        "Dataset Title": data_tt,
        "Dataset Link": data_ll
    }
    df = pd.DataFrame(data=my_dic)
    folder_path = '/Users/taneshkamehta/Documents/RPA/Web_Scraping/Kaggle_Datasets/ '
    absolute_path = os.path.abspath(folder_path)
    directory_path = os.path.dirname(absolute_path)
    file_name = 'Extracted_Datasets_%s.xlsx' %dt_stamp
    file_path = os.path.join(directory_path, file_name)
    df.to_excel(file_path, index=True)

    # IF .csv is required...
    # df.to_csv(file_path, index=False)

    context.close()
    browser.close()


def fn_run():
    # information('function f')
    with sync_playwright() as playwright:
        run(playwright)

if __name__ == "__main__":
    start = perf_counter()
    fn_run()
    end = perf_counter()
    print(f'\n---------------\n Finished in {round(end-start, 2)} second(s)')
