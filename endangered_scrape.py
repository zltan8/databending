# -*- coding: utf-8 -*-

from bs4 import BeautifulSoup
import time
from selenium import webdriver
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import xlwt
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import pykeyboard
import threading
from threading import Timer

# run chmod +x ./chromedriver first to enable run permissions
path_to_chromedriver = './chromedriver'
browser = webdriver.Chrome(executable_path=path_to_chromedriver)

wait10 = WebDriverWait(browser, 10)
timeout_time = 40

def get_page_source_tmall(url):
    browser.get(url)

browser.get("https://ecos.fws.gov/ecp0/reports/ad-hoc-species-report?kingdom=V&kingdom=I&status=E&status=T&status=EmE&status=EmT&status=EXPE&status=EXPN&status=SAE&status=SAT&fcrithab=on&fstatus=on&fspecrule=on&finvpop=on&fgroup=on&header=Listed+Animals")
time.sleep(2)

html = browser.page_source
soup = BeautifulSoup(html, "html.parser")
table = soup.find('table', attrs={'id': 'resultTable'})
rows = table.findAll('tr', attrs={'role': 'row'})

result1 = []
result2 = []

for row in rows:
    cols = row.findAll('td')
    if len(cols) == 5:
        print (cols[1])   
        result1.append(cols[1])
        result2.append(cols[2])

count = 0
with open('endangered.txt', 'w') as f:
    for item in result1:
        f.write("%s" % item)
        f.write("%s\n" % result2[count])
        count += 1
    