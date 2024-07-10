from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
import time

# Initialize Chrome options
options = Options()
options.add_experimental_option("detach", True)

# Load Excel workbook
try:
    wb = load_workbook(filename="D:\\Automation\\databaru3.xlsx")
    sheetRange = wb.active
except PermissionError as e:
    print(f"PermissionError: {e}")
    print("Ensure the Excel file is not open or locked by another application.")
    exit()

# Initialize WebDriver
driver = webdriver.Chrome(options=options)
url1 = 'https://subsiditepatlpg.mypertamina.id/merchant/auth/login'
url2 = 'https://subsiditepatlpg.mypertamina.id/merchant/app/verification-nik'
driver.maximize_window()
driver.implicitly_wait(10)
driver.get(url1)

# Login process
driver.find_element(By.ID, "mantine-r0").send_keys("087777334929")
driver.find_element(By.ID, "mantine-r1").send_keys("123456")
time.sleep(5)
driver.find_element(By.CLASS_NAME, "styles_btnLogin__wsKTT").click()
time.sleep(1)
driver.find_element(By.CLASS_NAME, "styles_iconClose__ZjGFM").click()
time.sleep(2)

def check_radio_buttons(driver):
    outcome = ""
    radio_buttons = {
        "R": '//label[contains(@class, "styles_container__vdRpf") and .//span[text()="Rumah Tangga"]]//input',
        "M": '//label[contains(@class, "styles_container__vdRpf") and .//span[text()="Usaha Mikro"]]//input',
        "P": '//label[contains(@class, "styles_container__vdRpf") and .//span[text()="Pengecer"]]//input'
    }
    for key, xpath in radio_buttons.items():
        try:
            driver.find_element(By.XPATH, xpath)
            outcome += key
        except NoSuchElementException:
            pass
    return outcome

def check_specific_div_elements(driver):
    outcome = ""
    specific_divs = {
        "R": '//div[contains(@class, "styles_containerCustomerInfo__N3jX5")]//span[text()="Rumah Tangga"]',
        "M": '//div[contains(@class, "styles_containerCustomerInfo__N3jX5")]//span[text()="Usaha Mikro"]',
        "P": '//div[contains(@class, "styles_containerCustomerInfo__N3jX5")]//span[text()="Pengecer"]'
    }
    for key, xpath in specific_divs.items():
        try:
            driver.find_element(By.XPATH, xpath)
            outcome = key
        except NoSuchElementException:
            pass
    return outcome

i = 147

while i <= len(sheetRange['A']):
    Nik = sheetRange['C' + str(i)].value
    Nama = sheetRange['B' + str(i)].value

    print(f"Data ke {i-1}, Nama : {Nama}, nik = {Nik}")

    wait = WebDriverWait(driver, 20)
    typetextfirst = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='search']")))

    typetextfirst.click()
    try:
        typetextfirst.send_keys(Nik)
        driver.find_element(By.CLASS_NAME, "styles_headerForm__t7P4g").click()
        time.sleep(2)
        driver.find_element(By.CLASS_NAME, "styles_btnBayar__o4O4A").click()
        time.sleep(2)

        try:
            wait.until(EC.presence_of_element_located((By.CLASS_NAME, "styles_btnModalStatusTrx__Hd0KY")))
            outcome = check_radio_buttons(driver)
        except TimeoutException:
            outcome = check_specific_div_elements(driver)
            if not outcome:
                outcome = "nik blm trdftr 1"

        # Update Excel cell with the outcome
        sheetRange[f"E{i}"] = outcome
        try:
            wb.save(filename="D:\\Automation\\databaru3.xlsx")
        except PermissionError as e:
            print(f"PermissionError: {e}")
            print("Ensure the Excel file is not open or locked by another application.")
            exit()
        
        print(f"Data ke {i-1} berhasil masuk. Outcome: {outcome}")

        driver.get(url2)

    except NoSuchElementException:
        print("NIK nlm terdftr 2")
        driver.get(url2)

    time.sleep(2)
    i += 1

print("Selesai Bos")
