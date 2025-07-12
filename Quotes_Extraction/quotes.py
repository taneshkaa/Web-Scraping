from playwright.sync_api import Playwright, sync_playwright, expect
import re
import os
import multiprocessing
import pandas as pd
import datetime as dt, time
from time import sleep, perf_counter
from openpyxl.workbook import Workbook

xl_fl = 'quotes_list.xlsx'

df = pd.read_excel(xl_fl)

# Excel Columns
# name = df["Name"].values.tolist()

url = 'https://quotes.toscrape.com/'
auth = []
quo = []

dt_tm = dt.datetime.fromtimestamp(time.time())
dt_stamp = dt_tm.strftime("%d-%B-%Y")

# def information(title):
#     print(title)
#     print('Module Nmae:', __name__)
#     print('Parent Process:', os.getppid())
#     print('Process ID:', os.getpid())

def run(playwright: Playwright) -> None:
    # Browser [Chrome, Firefox, WebKit]
    browser = playwright.chromium.launch(headless=False)
    
    context = browser.new_context()
    page = context.new_page()
    page.goto(url)
    sleep(0.5)

    # Browsing the website
    # page.reload()

    all_quotes = page.query_selector_all('.quote')
    for quote in all_quotes:
        text = quote.query_selector('.text').inner_text()
        author = quote.query_selector('.author').inner_text()
        print({'Author': author, 'Quote': text})
        auth.append(author)
        quo.append(text)
    page.wait_for_timeout(10000)


    my_dic = {
        "Author": auth,
        "Quote": quo
    }
    df = pd.DataFrame(data=my_dic)
    excel_path = '/Users/taneshkamehta/Documents/RPA/Web_Scraping/Quotes_Extraction/quotes_list.xlsx'
    absolute_path = os.path.abspath(excel_path)
    directory_path = os.path.dirname(absolute_path)
    file_name = 'Extracted_Quotes_%s.xlsx' %dt_tm
    file_path = os.path.join(directory_path, file_name)
    df.to_excel(file_path, index=True)
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
