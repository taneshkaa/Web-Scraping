from playwright.sync_api import Playwright, sync_playwright, expect
import re
import os
import multiprocessing
import pandas as pd
import datetime as dt, time
from time import sleep, perf_counter
from openpyxl.workbook import Workbook

# xl_fl = 'books_list.xlsx'

# df = pd.read_excel(xl_fl)

# Excel Columns
# name = df["Name"].values.tolist()

url = 'https://books.toscrape.com/'
gerne = url + 'catalogue/category/books/fantasy_19/'
auth = []
quo = []
st = []

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
    page.goto(gerne)
    sleep(1)

    # Browsing the website
    # page.reload()

    # Gerne -- Fantasy
    page.locator('(//div[@class="side_categories"]//ul//li)[19]').click()
    sleep(2)

    cnt = page.locator('//ol//li').count()
    print(cnt)

    for x in range(cnt):
        n = x+1
        
        # Book Name
        b = page.locator('(//ol//li)[%s]//h3/a' %n).inner_text()
        auth.append(str(b))

        # Book Price
        p = page.locator('(((//ol//li)[%s]//div)[2]/p)[1]' %n).inner_text()
        quo.append(str(p))

        # Book Availability
        a = page.locator('(((//ol//li)[%s]//div)[2]/p)[2]' %n).inner_text()
        st.append(str(a))


    my_dic = {
        "Book": auth,
        "Price": quo,
        "Availability": st
    }
    df = pd.DataFrame(data=my_dic)
    excel_path = '/Users/taneshkamehta/Documents/RPA/Web_Scraping/Books_Extraction/books_list.xlsx'
    absolute_path = os.path.abspath(excel_path)
    directory_path = os.path.dirname(absolute_path)
    file_name = 'Extracted_Books_%s.xlsx' %dt_stamp
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
