from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By

from openpyxl import load_workbook
import time

wb = load_workbook(filename=r'C:\Users\Dafa Zulfikar\Documents\myWork\Data-Entry\data.xlsx')


sheetRange = wb['Sheet1']

driver = webdriver.Chrome()

driver.get('https://demoqa.com/webtables')
driver.maximize_window()
driver.implicitly_wait(10)

i = 2

while i <= len(sheetRange['A']):
    firstName = sheetRange['A' + str(i)].value
    lastName = sheetRange['B' + str(i)].value
    email = sheetRange['C' + str(i)].value
    age = sheetRange['D' + str(i)].value
    salary = sheetRange['E' + str(i)].value
    department = sheetRange['F' + str(i)].value

    driver.find_element(By.XPATH, '//*[@id="addNewRecordButton"]').click()
    
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/div/div')))
        driver.find_element(By.ID,'firstName').send_keys(firstName)
        driver.find_element(By.ID,'lastName').send_keys(lastName)
        driver.find_element(By.ID,'userEmail').send_keys(email)
        driver.find_element(By.ID,'age').send_keys(age)
        driver.find_element(By.ID,'salary').send_keys(salary)
        driver.find_element(By.ID,'department').send_keys(department)
        driver.find_element(By.ID,'submit').click()
    except TimeoutException:
        print('Form gamuncul bang!')
        pass
    
    time.sleep(1)
    i += 1

print('muncul nih bang!')
