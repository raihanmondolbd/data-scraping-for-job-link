import time

from openpyxl.reader.excel import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import requests
from bs4 import BeautifulSoup
import pandas as pd
import datetime

job_search_topics = ['sqa test engineer', 'MANUAL QA', 'QA AUTOMATION', 'SOFTWARER TESTING', 'TEST AUTOMATION',
                     'MANUAL TESTING', 'PERFORMANCE TESTING', 'UNIT TESTING',
                     'Website scanning', 'Vulnerabilities finding & Reporting', 'Mobile Scanning', 'SQL Scanning',
                     'Metasploit', 'WordPress']

# job_search_topics = ['sqa test engineer', 'MANUAL QA', 'QA AUTOMATION']

'''Date Time set for excel file name'''
# dateTime = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
dateTime = datetime.datetime.now().strftime("%Y-%m-%d")
excelFileName = f'Virtual_Vacation_{dateTime}'

'''Added option for headless run'''
options = Options()
options.add_argument("--headless")
options.add_argument("--window-size=1920x1080")
'''Driver initialize'''
service = Service(executable_path=ChromeDriverManager().install())
# driver = webdriver.Chrome(options=options, service=service)
driver = webdriver.Chrome(service=service)
'''Open driver and search job in given domain'''
driver.get('https://www.virtualvocations.com/')
driver.maximize_window()
time.sleep(2)

for job_search in job_search_topics:
    ''' necessary variable define'''
    page = 1
    sum = 0
    global df
    data = {
        'Title': [],
        'Job Link': []

    }

    print(f'Job Searching for: {job_search}')
    driver.find_element(By.XPATH, '//input[@id="searchbox"]').clear()
    driver.find_element(By.XPATH, '//input[@id="searchbox"]').send_keys(job_search)
    driver.find_element(By.XPATH, '(//button[@type="submit"])[2]').click()
    url = driver.current_url
    # print(url)

    '''Using BeautifulSoup for data scraping'''
    while page != 10:  # use while condition for multiple pages
        response = requests.get(f'{url}/p-{page}')
        soup = BeautifulSoup(response.content, 'html.parser')
        jobs = soup.find('ul', class_="jobs-list list-unstyled").find_all('li')
        sum += len(jobs)

        for job in jobs:
            title = job.h2.a.get('title')
            job_link = job.h2.a.get('href')
            link = f'=HYPERLINK("{job_link}", "{job_link}")'

            data['Title'].append(title)
            # data['Job Link'].append(job_link)
            data['Job Link'].append(link)
        page += 1

    '''Convert Dictionary to a Dataframe'''
    df = pd.DataFrame(data, index=range(1, sum + 1))
    print(df)

    '''save data in excel'''
    try:
        with pd.ExcelWriter(f'{excelFileName}.xlsx', mode='a') as writer:
            df.to_excel(writer, sheet_name=job_search)
    except:
        with pd.ExcelWriter(f'{excelFileName}.xlsx') as writer:
            df.to_excel(writer, sheet_name=job_search)

    data.clear()
driver.quit()
