from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import csv
import time
import os

# --- KONFIGURACE ---
CSV_FILE = './Data.csv'
BASE_URL = "https://cz.pinterest.com/"
PIN_URL = "https://cz.pinterest.com/pin-creation-tool/"

print("=" * 60)
print("Pinterest Pin Automation")
print("=" * 60)

# 1. Inicializace prohlížeče
driver = webdriver.Chrome()
driver.get(BASE_URL)

print("\n✅ Chrome otevřen")
print("📝 Přihlaš se ručně na Pinterest")
print("   Až budeš na mainbordu, zapiš: done")

while input("Zadej 'done': ").lower() != 'done':
    pass

driver.get(PIN_URL)
time.sleep(3)
print("✅ Jsem na stránce Vytvořit pin\n")

# 2. Čti CSV a vytvárej piny
try:
    with open(CSV_FILE, mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        pins = list(reader)
        
    print(f"📋 Nalezeno {len(pins)} pinů\n")
    
    for idx, row in enumerate(pins, 1):
        print(f"\n{'='*60}")
        print(f"📌 Pin {idx}/{len(pins)}: {row['nazev']}")
        print(f"{'='*60}")
        
        wait = WebDriverWait(driver, 15)
        
        try:
            # 1. NAHRÁNÍ VIDEA
            print("→ Nahrávám video...")
            upload = wait.until(EC.presence_of_element_located((By.ID, "storyboard-upload-input")))
            upload.send_keys(os.path.abspath(row['video_cesta']))
            time.sleep(3)
            print("  ✅ Video nahrává se...")
            
            # 2. NADPIS
            print("→ Vyplňuji nadpis...")
            title = wait.until(EC.presence_of_element_located((By.ID, "storyboard-selector-title")))
            title.clear()
            title.send_keys(row['nazev'])
            time.sleep(2)
            print(f"  ✅ Nadpis: {row['nazev']}")
            
            # 3. ODKAZ
            print("→ Vyplňuji odkaz...")
            link = wait.until(EC.presence_of_element_located((By.ID, "WebsiteField")))
            link.clear()
            link.send_keys(row['odkaz'])
            time.sleep(2)
            print(f"  ✅ Odkaz: {row['odkaz']}")
            
            # 4. NÁSTĚNKA
            print(f"→ Vybírám nástěnku: {row['nastepka']}")
            board_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@data-test-id='board-dropdown-select-button']")))
            board_btn.click()
            time.sleep(2)
            
            board_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[@data-test-id='board-row-{row['nastepka']}']")))
            board_option.click()
            time.sleep(2)
            print(f"  ✅ Nástěnka: {row['nastepka']}")
            
            # 5. TAGY
            print("→ Přidávám tagy...")
            tags_string = row.get('tagy', '')
            if tags_string:
                # Rozdělí tagy podle středníku
                tags_list = [tag.strip() for tag in tags_string.split(';')]
                
                for tag in tags_list:
                    try:
                        # Vyhledej input pole pro tagy
                        tag_inputs = driver.find_elements(By.XPATH, "//input[@placeholder='Add a tag' or contains(@class, 'tag')]")
                        if tag_inputs:
                            tag_input = tag_inputs[-1]  # Vezmi poslední input
                            tag_input.click()
                            tag_input.send_keys(tag)
                            time.sleep(0.5)
                            tag_input.send_keys('\n')
                            time.sleep(0.3)
                            print(f"  ✅ Tag: {tag}")
                    except Exception as e:
                        print(f"  ⚠️ Tag '{tag}' selhalo: {e}")
            
            time.sleep(1)
            
            # 6. PUBLIKOVÁNÍ
            print("→ Publikuji pin...")
            publish_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@data-test-id='storyboard-creation-nav-done']")))
            publish_btn.click()
            time.sleep(5)
            print("✅ Pin publikován!")
            
            # Návrat na úvodní stránku
            driver.get(PIN_URL)
            time.sleep(3)
            
        except Exception as e:
            print(f"❌ Chyba: {e}")
            print("   Pokračuji dál...")
            continue

except FileNotFoundError:
    print(f"❌ Soubor {CSV_FILE} nenalezen!")

print("\n" + "="*60)
print("✅ HOTOVO!")
print("="*60)
driver.quit()
