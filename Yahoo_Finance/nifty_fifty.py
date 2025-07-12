from playwright.sync_api import Playwright, sync_playwright, expect
import re
import os
import multiprocessing
import pandas as pd
import datetime as dt, time
from time import sleep, perf_counter
from openpyxl.workbook import Workbook

# xl_fl = 'quotes_list.xlsx'

# df = pd.read_excel(xl_fl)

# Excel Columns
# name = df["Name"].values.tolist()

# https://finance.yahoo.com/quote/%5ENSEI/history/?guccounter=1&guce_referrer=aHR0cHM6Ly93d3cuZ29vZ2xlLmNvbS8&guce_referrer_sig=AQAAAD7EfT6diJPhH8ABBA7a-b9Dk7NC129zPEZvUIKiLwL-oKd79xpYP6UdtvzgVb4PFxK82PzMHtzT67Si-uHbrJWXJzWR9noXPu03o24sYjmif2LZaZuHgndbiTkeyEpCFezrmjaI8j09z74zTUAUg-ADRX8z4Ty0DmdKl0lFWQyt
# https://finance.yahoo.com/quote/%5ENSEI/history/

url = 'https://finance.yahoo.com/quote/%5ENSEI/history/?guccounter=1&guce_referrer=aHR0cHM6Ly93d3cuZ29vZ2xlLmNvbS8&guce_referrer_sig=AQAAAD7EfT6diJPhH8ABBA7a-b9Dk7NC129zPEZvUIKiLwL-oKd79xpYP6UdtvzgVb4PFxK82PzMHtzT67Si-uHbrJWXJzWR9noXPu03o24sYjmif2LZaZuHgndbiTkeyEpCFezrmjaI8j09z74zTUAUg-ADRX8z4Ty0DmdKl0lFWQyt'
dest_url = url + 'lookup/' 
date = []
open = []
close = []
high = []
low = []
adjc = []
vol = []

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
    # page.goto(dest_url)
    sleep(0.5)

    # Browsing the website
    # page.reload()

    # Download File
    with page.expect_download() as download_info:
        page.locator('//span[text()="Download"]').click()
    download = download_info.value
    # print(download)
    # print(download.path())

    folder_path = '/Users/taneshkamehta/Documents/RPA/Web_Scraping/Yahoo_Finance/ '

    # Downloading on Required PATH
    #  download.save_as('PATH' + download.suggested_filename)
    download.save_as(folder_path + download.suggested_filename)


    r_Cnt = page.locator('//tbody//tr').count()
    # print(r_Cnt)

    for i in range(r_Cnt):
        # print(i)
        n = i+1
        # print(n)

        # Table Structure - NIFTY 50
        # | Date | Open | High | Low | Close | Adj Close | Volume |

        # Date
        d = page.locator('((//tbody//tr)[%s]//td)[1]' %n).inner_text()
        date.append(str(d))
        # print(d)

        # Open
        o = page.locator('((//tbody//tr)[%s]//td)[2]' %n).inner_text()
        open.append(str(o))

        # High
        h = page.locator('((//tbody//tr)[%s]//td)[3]' %n).inner_text()
        high.append(str(h))

        # Low
        lw = page.locator('((//tbody//tr)[%s]//td)[4]' %n).inner_text()
        low.append(str(lw))

        # Close
        c = page.locator('((//tbody//tr)[%s]//td)[5]' %n).inner_text()
        close.append(str(c))

        # Adj Close
        ac = page.locator('((//tbody//tr)[%s]//td)[6]' %n).inner_text()
        adjc.append(str(ac))

        # Volume
        v = page.locator('((//tbody//tr)[%s]//td)[7]' %n).inner_text()
        vol.append(str(v))


    my_dic = {
        "Date": date,
        "Open": open,
        "High": high,
        "Low": low,
        "Close": close,
        "Adj Close": adjc,
        "Volume": vol
    }

    df = pd.DataFrame(data=my_dic)
    excel_path = '/Users/taneshkamehta/Documents/RPA/Web_Scraping/Yahoo_Finance/'
    absolute_path = os.path.abspath(excel_path)
    directory_path = os.path.dirname(absolute_path)
    file_name = 'Extracted_Trending_Stocks_Nifty_%s.xlsx' %dt_tm
    file_path = os.path.join(excel_path, file_name)
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
