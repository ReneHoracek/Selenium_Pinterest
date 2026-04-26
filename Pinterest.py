from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import csv
import time
import os


# --- KONFIGURACE ---
CSV_FILE = './Python/Selenium/Pinterest/Data.csv'
BASE_URL = "https://cz.pinterest.com/"
SECOND_URL = "https://cz.pinterest.com/pin-creation-tool/"

# 1. Inicializace prohlížeče
driver = webdriver.Chrome()
driver.get(BASE_URL)


print("Ručně se příhlásit na Pinterest.")
print("Až budeš na stránce 'Mainbordu', stiskni v tomto terminálu ENTER...")
input()

driver.get(SECOND_URL)
print("Až budeš na stránce 'Vytvořit pin', stiskni v tomto terminálu ENTER...")
input()

try:
    with open(CSV_FILE, mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        for row in reader:
            wait = WebDriverWait(driver, 10)


            # Nahrání videa
            upload = wait.until(EC.presence_of_element_located((By.ID, "storyboard-upload-input")))
            upload.send_keys(row['video_cesta'])
            time.sleep(5)
            input()

            # Název
            title = wait.until(EC.presence_of_element_located((By.ID, "storyboard-selector-title")))
            title.clear()
            title.send_keys(row['nazev'])
            time.sleep(5)
            input()

            # Odkaz
            link = wait.until(EC.presence_of_element_located((By.ID, "WebsiteField")))
            link.clear()
            link.send_keys(row['odkaz'])
            time.sleep(5)
            input()

            # Nástěnka
            print(f"🔍 Hledám tlačítko nástěnky...")
            board_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@data-test-id='board-dropdown-select-button']")))
            print(f"✅ Tlačítko nástěnky nalezeno, klikám...")
            board_btn.click()
            time.sleep(5)
            input()

            # Vyber nástěnku podle jména z CSV
            print(f"🔍 Hledám nástěnku '{row['nastepka']}'...")
            board_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[@data-test-id='board-row-{row['nastepka']}']")))
            print(f"✅ Nástěnka nalezena, klikám...")
            board_option.click()
            time.sleep(5)
            input()

            print(f"✅ Zpracováno: {row['nazev']}")

            # Publikovat
            publish_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@data-test-id='storyboard-creation-nav-done']")))
            publish_btn.click()
            time.sleep(5)
            input()

            driver.get(SECOND_URL)
            time.sleep(5)
            input()

except FileNotFoundError:
    print(f"❌ Soubor {CSV_FILE} nebyl nalezen.")

print("\n--- HOTOVO ---")
"driver.quit()"
