from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import openpyxl
import time
import xlrd

driver = webdriver.Chrome('chromedriver.exe')
driver.maximize_window()
driver.get('http://localhost:4200/#/')

def login_categorytype(email,password,name,status,description,amount,providers):

    user_input = driver.find_element_by_id('email')
    user_input.send_keys(email)

    pass_input = driver.find_element_by_id('pass')
    pass_input.send_keys(password)

    login_button = driver.find_element_by_id('lex')
    login_button.click()
    time.sleep(3)

    categorytype_button = driver.find_element_by_xpath('//*[@id="sidebar"]/ul/div[3]/li/a')
    categorytype_button.click()
    time.sleep(3)

    add_button = driver.find_element_by_xpath('//*[@id="body"]/div/div/div/div/div/app-content/app-c5-categorytype/mat-card[4]/div/div/button[1]')
    add_button.click()
    time.sleep(3)

    name_input = driver.find_element_by_id('name')
    name_input.send_keys(name)

    status_input = driver.find_element_by_id('status')
    status_input.send_keys(status)

    description_input = driver.find_element_by_id('description')
    description_input.send_keys(description)

    amount_input = driver.find_element_by_id('amount')
    amount_input.send_keys(amount)   

    providers_input = driver.find_element_by_id('providers')
    providers_input.send_keys(providers)

    accept_button = driver.find_element_by_id('accept')
    accept_button.click()
    time.sleep(3)


file_location = "D:\Test\Form.xlsx"
wb = xlrd.open_workbook(file_location)
sheet = wb.sheet_by_index(0)

lst = []
for rows in range(sheet.nrows):
	lst.append(sheet.cell_value(rows, 5))

email =lst[2]
password = lst[3]
name =lst[7]
status = lst[8]
description= lst[9]
amount = lst[10]
providers=lst[11]


login_categorytype(email,password,name,status,description,amount,providers)