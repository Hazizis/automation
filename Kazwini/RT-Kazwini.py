from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
import time

options = Options()
options.add_experimental_option("detach", True)

wb = load_workbook(filename="D:\Automation\Kazwini\datakazwani.xlsx")
sheetRange = wb.active

driver = webdriver.Chrome(options=options)
url1 = 'https://subsiditepatlpg.mypertamina.id/merchant/auth/login'
url2 = 'https://subsiditepatlpg.mypertamina.id/merchant/app/verification-nik'
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

i = 2
n = 0

while i <= len(sheetRange['A']):
    Nik = sheetRange['C' + str(i)].value
    Nama = sheetRange['B' + str(i)].value

    print("Data ke " + str(i-1) + ", Nama : " + str(Nama) + ", nik = " + str(Nik))

    wait = WebDriverWait(driver, 30)
    waitfaster= WebDriverWait(driver, 10)
    typetextfirst = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='search']")))
    time.sleep(1)

    typetextfirst.click()
    try:
        typetextfirst.send_keys(Nik)
        driver.find_element(By.CLASS_NAME, "styles_headerForm__t7P4g").click()
        time.sleep(2)
        driver.find_element(By.CLASS_NAME, "styles_btnBayar__o4O4A").click()
        time.sleep(2)

        # Check if the modal button is present
        modal_buttons = driver.find_elements(By.CLASS_NAME, "styles_btnModalStatusTrx__Hd0KY")
        if modal_buttons:
            # Click the "Rumah Tangga" radio button
            radio_button_xpath = '//label[contains(@class, "styles_container__vdRpf") and .//span[text()="Rumah Tangga"]]//input'
            radio_button = driver.find_element(By.XPATH, radio_button_xpath)
            driver.execute_script("arguments[0].click();", radio_button)
            time.sleep(1)
            
            # Click the "Lanjut Transaksi" button
            modal_buttons[0].click()
            time.sleep(2)
        else:
            print("Modal button not found, proceeding with next steps")

        # Wait until the element with text "Rp. 18.000" appears
        element_with_price = waitfaster.until(EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/div[1]/div/main/div/div/div/div/div/div/form/div[1]/div[3]/div/span[2]')))
        print("Element with price found:", element_with_price.text)
        driver.find_element(By.CLASS_NAME, "styles_btnBayar__blJ1W").click()
        time.sleep(2)
        driver.find_element(By.CLASS_NAME, "styles_btnBayar__moyir").click()
        time.sleep(2)
        driver.get(url2)
        print("Data ke " + str(i-1) + " berhasil masuk.")

        n += 1
        if n % 5 == 0 and n != 1:
            time.sleep(20)

    except NoSuchElementException:
        print("Data melebihi batas")
        driver.get(url2)

    time.sleep(2)
    i += 1

print("Selesai Bos")
