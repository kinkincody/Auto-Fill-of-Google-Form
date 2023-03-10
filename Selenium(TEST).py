import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager as cdm
import time
from selenium.webdriver.support.ui import Select
import xlwings as xw
import os
from selenium.webdriver.common.keys import Keys

wb = xw.Book(r'1.xlsx')
ws = wb.sheets


for i in range(2,100):
    name = ws[0].cells(i,1).value
    clas = ws[0].cells(i,2).value
    clas2 = ws[0].cells(i,3).value
    sub = ws[0].cells(i,4).value
    li1 = ws[0].cells(i,5).value
    li2 = ws[0].cells(i,6).value
    mul = ws[0].cells(i,7).value
    
    options = webdriver.ChromeOptions().add_argument('--disable-notifications')
    #user-data-dir=C:\\Users\\user\\AppData\\Local\\Google\\Chrome\\User Data will add cookies
    chrome = webdriver.Chrome(cdm().install(), options = options)
    chrome.get('https://forms.gle/ZWKxZPXkcmoSzYiv9') #google form link
    
    chrome.maximize_window()
    time.sleep(3)

    q1 = chrome.find_element(By.XPATH, f'//input[@jsname="YPqjbf"]') #input normal text
    q1.send_keys(f'{name}')
    
    q2 = chrome.find_elements(By.XPATH, f'//div[@class="lLfZXe fnxRtf cNDBpf"]//div//div//div//div//div[@aria-label={clas}]') #radio button1
    q2[0].click()
    
    q3 = chrome.find_elements(By.XPATH, '//div[@class="lLfZXe fnxRtf cNDBpf"]//div//div//div//div//div[@aria-label="1"]') # radio button with similar answer
    q3[1].click()
    
    q4 = chrome.find_elements(By.XPATH, f'//div[@role="list"]//div//div[@aria-label={clas2}]') #checkbox
    q4[0].click()
    
    q5 = chrome.find_elements(By.XPATH, '//div[@role="listbox"]') #dropdown menu
    #google form drop down menu is a bit different, there is no <select> in source code, need to indicate the exact XPATH
    q5[0].click()
    time.sleep(1)
    
    q5a = chrome.find_elements(By.XPATH, '//div[@role="listbox"][@aria-labelledby="i60"]//div//div[@role="option"][@data-value="中文"]')
    q5a[0].click()
    
    q6 = chrome.find_elements(By.XPATH, '//div[@role="listbox"]')
    q6[1].click()
    time.sleep(1)
    
    q6a = chrome.find_elements(By.XPATH, '//div[@role="listbox"][@aria-labelledby="i64"]//div//div[@role="option"][@data-value="3"]')
    q6a[0].click()
    
    time.sleep(1)
    
    q7 = chrome.find_elements(By.XPATH, '//div[@aria-label="column1，對「row1」的回應"]')#multiple choice
    q7[0].click()
    
    submit = chrome.find_element(By.XPATH, '//span[contains(text(), "提交")]')
    #submit, if using English, please convert the Chinese into English
    submit.click()
    
    time.sleep(2)
    
    chrome.close()
    
    
    #the time.sleep is used for ensure the script does not run too fast and error. Wait can be used instead

