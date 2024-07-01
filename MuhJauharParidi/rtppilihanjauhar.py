from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
import time

options = Options()
options.add_experimental_option("detach", True)

wb = load_workbook(filename="D:\Automation\MuhJauharParidi\DataMuhJauharParidi.xlsx")

sheetRange = wb.active


driver = webdriver.Chrome(options=options)
url1='https://subsiditepatlpg.mypertamina.id/merchant/auth/login'
url2='https://subsiditepatlpg.mypertamina.id/merchant/app/verification-nik'
driver.maximize_window()
driver.implicitly_wait(10)
driver.get(url1)
driver.find_element(By.ID, "mantine-r0").send_keys("087777334929")
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
        # wait.until(EC.invisibility_of_element_located(By.XPATH, "//*[@id='mantine-rhd-body']/label[1]/span[1]")).click()
        time.sleep(1)
        wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "styles_btnModalStatusTrx__Hd0KY"))).click()
        time.sleep(1)
        driver.find_element(By.CLASS_NAME, "styles_btnBayar__blJ1W").click()
        time.sleep(2)
        driver.find_element(By.CLASS_NAME, "styles_btnBayar__moyir").click()
        time.sleep(2)
    
        driver.get("https://subsiditepatlpg.mypertamina.id/merchant/app/verification-nik")
        print("Data ke "+str(i-1)+" berhasil masuk.")
        n = n+1
        if n % 5 == 0 and n != 1:
            time.sleep(20)
        
        
    except NoSuchElementException:
        # WebDriverWait(driver,20).until(EC.presence_of_element_located(By.XPATH, "//*[@id='__next']/div[1]/div/main/div/div/div/div/div/div/div[2]/div[3]/div/span"))
        print("Data melebihi batas")
        driver.get("https://subsiditepatlpg.mypertamina.id/merchant/app/verification-nik")

    time.sleep(2)
    i = i+1
print ("Selesai Bos")


