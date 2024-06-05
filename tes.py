from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
import time

options = Options()
options.add_experimental_option("detach", True)

wb = load_workbook(filename="D:\Autmation\data.xlsx")

sheetRange = wb.active


driver = webdriver.Chrome(options=options)
url1='https://subsiditepatlpg.mypertamina.id/merchant/auth/login'
url2='https://subsiditepatlpg.mypertamina.id/merchant/app/verification-nik'
driver.maximize_window()
driver.implicitly_wait(10)
driver.get(url1)
driver.find_element(By.ID, "mantine-r0").send_keys("085337580232")
time.sleep(1)
driver.find_element(By.ID, "mantine-r1").send_keys("123456")
time.sleep(1)
driver.find_element(By.CLASS_NAME, "styles_btnLogin__wsKTT").click()
time.sleep(1)
driver.find_element(By.CLASS_NAME, "styles_iconClose__ZjGFM").click()
time.sleep(2)
#driver.get(url2)




i=2

while i <= len(sheetRange['A']):
    Nik = sheetRange['C'+str(i)].value
    
    
    print ("data ke "+str(i)+",nik = "+str(Nik))
    '''
    
    '''
    
    wait = WebDriverWait(driver, 30)
    typetextfirst = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='search']")))
    time.sleep(1)

    typetextfirst.click()
    try:
        
        typetextfirst.send_keys(Nik)        
        driver.find_element(By.CLASS_NAME, "styles_headerForm__t7P4g").click()
        time.sleep(2)
    
        driver.find_element(By.CLASS_NAME, "styles_btnBayar__o4O4A").click()
        time.sleep(2)
        
        driver.find_element(By.CLASS_NAME, "styles_btnBayar__blJ1W").click()
        time.sleep(2)
        driver.find_element(By.CLASS_NAME, "styles_btnBayar__moyir").click()
        time.sleep(2)
    
        driver.get("https://subsiditepatlpg.mypertamina.id/merchant/app/verification-nik")
        print("data berhasil masuk = "+str(i))

    except TimeoutException:
        print ("data tidak masuk")
        pass
    time.sleep(1)
    i = i+1
print ("Udahan")

'''
i=2

while i <= len(sheetRange['C']):
    Nik = sheetRange['C'+str(i)].value

    try:
        driver.find_element(By.ID, "mantine-ra").send_keys(Nik)
        time.sleep(5)
        driver.find_element(By.ID, "btnCheckNik").click()

    except TimeoutException:
        print ("data tidak masuk")
        pass

    time.sleep(5)
'''
