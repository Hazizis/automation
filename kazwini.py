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

wb = load_workbook(filename="D:\Autmation\datakazwani.xlsx")

sheetRange = wb.active


driver = webdriver.Chrome(options=options)
url1='https://subsiditepatlpg.mypertamina.id/merchant/auth/login'
url2='https://subsiditepatlpg.mypertamina.id/merchant/app/verification-nik'
driver.maximize_window()
driver.implicitly_wait(10)
driver.get(url1)
driver.find_element(By.ID, "mantine-r0").send_keys("081918104667")
time.sleep(2)
driver.find_element(By.ID, "mantine-r1").send_keys("123456")
time.sleep(5)
driver.find_element(By.CLASS_NAME, "styles_btnLogin__wsKTT").click()
time.sleep(1)
driver.find_element(By.CLASS_NAME, "styles_iconClose__ZjGFM").click()
time.sleep(2)

i=2
n = 0

while i <= len(sheetRange['A']):
    Nik = sheetRange['C'+str(i)].value
    Nama = sheetRange['B'+str(i)].value
       
    print ("Data ke "+str(i-1)+", Nama : "+str(Nama) +",nik = "+str(Nik))
    
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
        print("Data ke "+str(i-1)+" berhasil masuk.")
        n = n+1
        if n % 5 == 0 and n != 1:
            time.sleep(20)
    except TimeoutException:
        print ("Data tidak masuk")
        pass
    time.sleep(2)
    i = i+1
print ("Selesai Bos")
