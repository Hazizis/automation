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

wb = load_workbook(filename="D:\Autmation\datarijaliana.xlsx")

sheetRange = wb.active


driver = webdriver.Chrome(options=options)
url1='https://subsiditepatlpg.mypertamina.id/merchant/auth/login'
url2='https://subsiditepatlpg.mypertamina.id/merchant/app/verification-nik'
driver.maximize_window()
driver.implicitly_wait(10)
driver.get(url1)
time.sleep(2)
driver.find_element(By.ID, "mantine-r0").send_keys("081909053084")
time.sleep(2)
driver.find_element(By.ID, "mantine-r1").send_keys("271087")
time.sleep(5)
driver.find_element(By.CLASS_NAME, "styles_btnLogin__wsKTT").click()
time.sleep(1)
driver.find_element(By.CLASS_NAME, "styles_iconClose__ZjGFM").click()
time.sleep(2)
#driver.get(url2)

i=203
n = 0

while i <= len(sheetRange['A']):
    Nik = sheetRange['C'+str(i)].value
    Nama = sheetRange['B'+str(i)].value
       
    print ("Data ke "+str(i-1)+", Nama : "+str(Nama) +",nik = "+str(Nik))
    
    wait = WebDriverWait(driver, 30)
    typetextfirst = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='search']")))
    time.sleep(1)

    typetextfirst.click()
    typetextfirst.send_keys(Nik)        
    driver.find_element(By.CLASS_NAME, "styles_headerForm__t7P4g").click()
    time.sleep(2)   
    driver.find_element(By.CLASS_NAME, "styles_btnBayar__o4O4A").click()
    time.sleep(2)
    try:
        if  wait.until(EC.visibility_of_element_located((By.ID, "mantine-r5-body"))):
            time.sleep(1)
            driver.find_element(By.XPATH, "//*[@id='mantine-r5-body']/div/div[2]/button").click()
            print("Data NIK belum terdaftar")
            driver.get("https://subsiditepatlpg.mypertamina.id/merchant/app/verification-nik")

        elif wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "button[type='submit']" ))):
            driver.find_element(By.CLASS_NAME, "styles_btnBayar__blJ1W").click()
            time.sleep(2)
            driver.find_element(By.CLASS_NAME, "styles_btnBayar__moyir").click()
            time.sleep(2)
            print("Data ke "+str(i-1)+" berhasil masuk.")
            driver.get("https://subsiditepatlpg.mypertamina.id/merchant/app/verification-nik")
            
            n = n+1
            if n % 5 == 0 and n != 1:
                time.sleep(20)  
        elif wait.until(EC.visibility_of_element_located((By.ID, "mantine-r2m-body"))):
            driver.find_element(By.XPATH, "//*[@id='mantine-r2m-body']/label[2]/span[2]").click()
            time.sleep(1)
            driver.find_element(By.CLASS_NAME, "styles_btnModalStatusTrx__Hd0KY").click()

        else:
            print("tidak berjalan")
    except TimeoutException:
        print ("Data tidak masuk")
        pass
    i = i+1
print ("Selesai Bos")
