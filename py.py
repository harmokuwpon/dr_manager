
import time

import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from msedge.selenium_tools import Edge, EdgeOptions
from selenium import webdriver

path = r"C:\py\msedgedriver.exe"
       #ヘッドレス
#options.add_argument("disable-gpu")
driver = webdriver.Edge(executable_path='C:\py\msedgedriver.exe')#EDGEドライバー保管先　保管先によって変更すること。
driver.get('https://iab-bp.omron.co.jp/drppe/')#DRマネージャーURL



def driver_init():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    return webdriver.Chrome(options=options)


print(driver.find_element_by_xpath('/html/body/div[1]/div/section/div/div/h2').text)
print(driver.current_url)
driver.quit()
