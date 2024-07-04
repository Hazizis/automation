from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
import time

options = Options()
options.add_experimental_option("detach", True)

try:
    wb = load_workbook(filename="D:\\Automation\\databaru.xlsx")
    sheetRange = wb.active
except PermissionError as e:
    print(f"PermissionError: {e}")
    print("Ensure the Excel file is not open or locked by another application.")
    exit()

driver = webdriver.Chrome(options=options)
url1 = 'https://subsiditepatlpg.mypertamina.id/merchant/auth/login'
url2 = 'https://subsiditepatlpg.mypertamina.id/merchant/app/verification-nik'
driver.maximize_window()
driver.implicitly_wait(10)
driver.get(url1)
driver.find_element(By.ID, "mantine-r0").send_keys("081803475364")
time.sleep(2)
driver.find_element(By.ID, "mantine-r1").send_keys("123456")
time.sleep(5)
driver.find_element(By.CLASS_NAME, "styles_btnLogin__wsKTT").click()
time.sleep(1)
driver.find_element(By.CLASS_NAME, "styles_iconClose__ZjGFM").click()
time.sleep(2)

i = 2

while i <= len(sheetRange['A']):
    Nik = sheetRange['C' + str(i)].value
    Nama = sheetRange['B' + str(i)].value

    print("Data ke " + str(i-1) + ", Nama : " + str(Nama) + ", nik = " + str(Nik))

    wait = WebDriverWait(driver, 20)
    typetextfirst = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='search']")))
    time.sleep(1)

    typetextfirst.click()
    try:
        typetextfirst.send_keys(Nik)
        driver.find_element(By.CLASS_NAME, "styles_headerForm__t7P4g").click()
        time.sleep(2)
        driver.find_element(By.CLASS_NAME, "styles_btnBayar__o4O4A").click()
        time.sleep(2)

        # Initialize outcome
        outcome = ""

        try:
            wait.until(EC.presence_of_element_located((By.CLASS_NAME, "styles_btnModalStatusTrx__Hd0KY")))

            # Check "Rumah Tangga" radio button
            try:
                RT_radio_button_xpath = '//label[contains(@class, "styles_container__vdRpf") and .//span[text()="Rumah Tangga"]]//input'
                RT_radio_button_presence = driver.find_element(By.XPATH, RT_radio_button_xpath)
                outcome += "R"
            except NoSuchElementException:
                pass

            # Check "Usaha Mikro" radio button
            try:
                UM_radio_button_xpath = '//label[contains(@class, "styles_container__vdRpf") and .//span[text()="Usaha Mikro"]]//input'
                UM_radio_button_presence = driver.find_element(By.XPATH, UM_radio_button_xpath)
                outcome += "M"
            except NoSuchElementException:
                pass

            # Check "Pengecer" radio button
            try:
                P_radio_button_xpath = '//label[contains(@class, "styles_container__vdRpf") and .//span[text()="Pengecer"]]//input'
                P_radio_button_presence = driver.find_element(By.XPATH, P_radio_button_xpath)
                outcome += "P"
            except NoSuchElementException:
                pass

        except TimeoutException:
            try:
                # Generalized check for the specific div element "Rumah Tangga"
                specific_div_rt_xpath = '//div[contains(@class, "styles_containerCustomerInfo__N3jX5")]//span[text()="Rumah Tangga"]'
                driver.find_element(By.XPATH, specific_div_rt_xpath)
                outcome = "R"
            except NoSuchElementException:
                pass

            try:
                # Generalized check for the specific div element "Usaha Mikro"
                specific_div_um_xpath = '//div[contains(@class, "styles_containerCustomerInfo__N3jX5")]//span[text()="Usaha Mikro"]'
                driver.find_element(By.XPATH, specific_div_um_xpath)
                outcome = "M"
            except NoSuchElementException:
                pass

            try:
                # Generalized check for the specific div element "Pengecer"
                specific_div_p_xpath = '//div[contains(@class, "styles_containerCustomerInfo__N3jX5")]//span[text()="Pengecer"]'
                driver.find_element(By.XPATH, specific_div_p_xpath)
                outcome = "P"
            except NoSuchElementException:
                pass

        # Update Excel cell with the outcome
        sheetRange[f"E{i}"] = outcome
        try:
            wb.save(filename="D:\\Automation\\databaru.xlsx")
        except PermissionError as e:
            print(f"PermissionError: {e}")
            print("Ensure the Excel file is not open or locked by another application.")
            exit()
        
        print("Data ke " + str(i-1) + " berhasil masuk. Outcome: " + outcome)

        driver.get(url2)

    except NoSuchElementException:
        print("NIK nlm terdftr 2")
        driver.get(url2)

    time.sleep(2)
    i += 1

print("Selesai Bos")
