"""
Pinterest Pin Automation System
Automatizace vytváření pinů s validací, plánováním a inteligentním střídáním
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pandas as pd
import csv
import time
import os
import json
from datetime import datetime, timedelta
from collections import defaultdict
import logging

# ========== KONFIGURACE ==========
CSV_FILE = './Data.csv'
BASE_URL = "https://cz.pinterest.com/"
PIN_CREATION_URL = "https://cz.pinterest.com/pin-creation-tool/"
SCHEDULE_FILE = 'schedule.json'
VALIDATION_LOG = 'validation_report.txt'

# Časy pro plánování
SCHEDULE_TIMES = {
    0: [(7, 30), (12, 30), (19, 30)],      # Pondělí: 3 posty
    1: [(7, 30), (12, 30), (19, 30)],      # Úterý: 3 posty
    2: [(7, 30), (12, 30), (19, 30)],      # Středa: 3 posty
    3: [(7, 30), (12, 30), (19, 30)],      # Čtvrtek: 3 posty
    4: [(7, 30), (12, 30)],                # Pátek: 2 posty
    5: [(18, 30)],                          # Sobota: 1 post
    6: [(18, 30)]                           # Neděle: 1 post
}

# Nastavení loggingu
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('pinterest_automation.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ========== TŘÍDY ==========

class PinterestValidator:
    """Validace dat z CSV/Excelu"""
    
    def __init__(self):
        self.errors = []
        self.warnings = []
        self.data = []
    
    def validate_file(self, filepath):
        """Validace celého souboru"""
        try:
            # Zkus CSV nejdřív (jednodušší a spolehlivější)
            df = pd.read_csv(filepath, encoding='utf-8')
            logger.info(f"Soubor načten (CSV): {len(df)} řádků")
        except Exception as e:
            try:
                # Fallback na XLSX
                df = pd.read_excel(filepath, engine='openpyxl', dtype=str)
                logger.info(f"Soubor načten (XLSX): {len(df)} řádků")
            except Exception as e2:
                logger.error(f"Chyba při načítání souboru: {e} / {e2}")
                return False
        
        # Vyčisti sloupce - odstran ALL whitespace (včetně newline) a konvertuj na lowercase
        df.columns = df.columns.str.strip().str.replace('\n', '').str.replace('\r', '').str.lower()
        logger.info(f"Dostupné sloupce po vyčištění: {list(df.columns)}")
        
        # Mapování sloupců (toleruj různé názvy)
        column_mapping = {
            'nazev': 'nazev',
            'název': 'nazev',
            'video_cesta': 'video_cesta',
            'video cesta': 'video_cesta',
            'cesta': 'video_cesta',
            'odkaz': 'odkaz',
            'odkaz na web': 'odkaz',
            'nastepka': 'nastepka',
            'nástěnka': 'nastepka',
            'board': 'nastepka',
            'tagy': 'tagy',
            'tag': 'tagy',
            'tags': 'tagy',
            'barva_satu': 'barva_satu',
            'barva': 'barva_satu',
            'barva šatů': 'barva_satu',
            'popis': 'popis',  # Nepovinný
            'description': 'popis'
        }
        
        # Aplikuj mapování
        rename_dict = {}
        for col in df.columns:
            if col in column_mapping:
                rename_dict[col] = column_mapping[col]
        
        df = df.rename(columns=rename_dict)
        
        # Kontrola povinných sloupců
        required_columns = ['nazev', 'video_cesta', 'odkaz', 'nastepka']
        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            self.errors.append(f"Chybějící sloupce: {', '.join(missing_cols)}")
            logger.error(f"Dostupné sloupce: {list(df.columns)}")
            return False
        
        # Přidej chybějící nepovinné sloupce s výchozími hodnotami
        if 'tagy' not in df.columns:
            df['tagy'] = 'účetnictví'
            logger.warning("Sloupec 'tagy' chybí - přidáno výchozí 'účetnictví'")
        
        if 'barva_satu' not in df.columns:
            df['barva_satu'] = 'Různobarevná'
            logger.warning("Sloupec 'barva_satu' chybí - přidáno výchozí 'Různobarevná'")
        
        # Validace jednotlivých řádků
        seen_titles = set()
        for idx, row in df.iterrows():
            row_num = idx + 2  # +2 kvůli headeru a indexování od 1
            
            # Nadpis
            if pd.isna(row['nazev']) or not str(row['nazev']).strip():
                self.errors.append(f"Řádek {row_num}: Nadpis je povinný")
            elif str(row['nazev']) in seen_titles:
                self.errors.append(f"Řádek {row_num}: Duplicitní nadpis '{row['nazev']}'")
            else:
                seen_titles.add(str(row['nazev']))
            
            # Video cesta
            if pd.isna(row['video_cesta']) or not str(row['video_cesta']).strip():
                self.errors.append(f"Řádek {row_num}: Video cesta je povinná")
            # TODO: Kontrola existence videa se bude dělat později v Selenium botech
            # (aby se nemusely videa kopírovat pro testování validace)
            
            # Odkaz
            if pd.isna(row['odkaz']):
                self.errors.append(f"Řádek {row_num}: Odkaz je povinný")
            elif not self._is_valid_url(str(row['odkaz'])):
                self.warnings.append(f"Řádek {row_num}: Podivný formát odkazu: {row['odkaz']}")
            
            # Nástěnka
            valid_boards = [
                'daňové-poradenství',
                'podnikání-online',
                'rady-a-tipy',
                'uol-účetnictví',
                'účetnictví-online'
            ]
            if pd.isna(row['nastepka']):
                self.errors.append(f"Řádek {row_num}: Nástěnka je povinná")
            else:
                # Normalizuj nástěnku - změň mezery na pomlčky a smaž háčky/čárky
                board_normalized = str(row['nastepka']).lower().strip().replace(' ', '-')
                # Oprav diakritiku (jednoduchá verze)
                board_normalized = board_normalized.replace('á', 'a').replace('í', 'i').replace('ů', 'u').replace('ě', 'e')
                
                if board_normalized not in valid_boards:
                    self.warnings.append(
                        f"Řádek {row_num}: Nástěnka '{row['nastepka']}' -> '{board_normalized}' "
                        f"není v seznamu. Dostupné: {', '.join(valid_boards)}"
                    )
            
            # Tagy
            if pd.isna(row['tagy']):
                self.warnings.append(f"Řádek {row_num}: Tagy nejsou vyplněny")
            else:
                tags = str(row['tagy']).split(',')
                if not any('účetnictví' in tag.lower() for tag in tags):
                    self.warnings.append(f"Řádek {row_num}: Povinný tag 'účetnictví' chybí")
            
            # Barva šatů
            if pd.isna(row['barva_satu']):
                self.warnings.append(f"Řádek {row_num}: Barva šatů není vyplněna")
            
            self.data.append(row.to_dict())
        
        return len(self.errors) == 0
    
    @staticmethod
    def _is_valid_url(url):
        """Jednoduchá kontrola URL"""
        return url.startswith(('http://', 'https://'))
    
    def print_report(self):
        """Tisk validační zprávy"""
        with open(VALIDATION_LOG, 'w', encoding='utf-8') as f:
            f.write("=" * 60 + "\n")
            f.write("ZPRÁVA O VALIDACI\n")
            f.write("=" * 60 + "\n\n")
            
            if self.errors:
                f.write("❌ CHYBY (musí být opraveny):\n")
                for error in self.errors:
                    f.write(f"  - {error}\n")
                f.write("\n")
            
            if self.warnings:
                f.write("⚠️  VAROVÁNÍ (doporučujeme zkontrolovat):\n")
                for warning in self.warnings:
                    f.write(f"  - {warning}\n")
                f.write("\n")
            
            if not self.errors and not self.warnings:
                f.write("✅ VŠECHNA DATA JSOU V POŘÁDKU!\n\n")
            
            f.write(f"Celkem řádků k publikování: {len(self.data)}\n")
        
        logger.info(f"Zpráva o validaci uložena: {VALIDATION_LOG}")


class ScheduleOptimizer:
    """Optimalizuje plánování s ohledem na střídání barev a nástěnek"""
    
    def __init__(self, data):
        self.data = data
        self.schedule = []
    
    def optimize_schedule(self):
        """Vytvoří optimalizovaný plán publikování"""
        # Seřazení dat podle barvy (aby se střídala)
        sorted_data = self._sort_by_color_and_board(self.data)
        
        # Přiřazení časů
        time_slots = self._get_available_time_slots()
        
        for idx, pin_data in enumerate(sorted_data):
            if idx < len(time_slots):
                publish_time = time_slots[idx]
                pin_data['scheduled_time'] = publish_time.isoformat()
                self.schedule.append(pin_data)
        
        return self.schedule
    
    @staticmethod
    def _sort_by_color_and_board(data):
        """Seřadí piny podle barvy a nástěnky, aby se střídaly"""
        # Seskupení po barvě a nástěnce
        groups = defaultdict(list)
        for item in data:
            key = (item.get('barva_satu', 'unknown'), item.get('nastepka', 'unknown'))
            groups[key].append(item)
        
        # Round-robin přiřazení
        sorted_items = []
        active_groups = list(groups.values())
        
        while active_groups:
            next_round = []
            for group in active_groups:
                if group:
                    sorted_items.append(group.pop(0))
                    if group:
                        next_round.append(group)
            active_groups = next_round
        
        return sorted_items
    
    @staticmethod
    def _get_available_time_slots():
        """Generuje seznam všech dostupných časů"""
        slots = []
        now = datetime.now()
        
        # Sebereme dalších 30 dní
        for days_offset in range(30):
            current_date = now + timedelta(days=days_offset)
            weekday = current_date.weekday()
            
            # Pokud máme časy pro tento den v týdnu
            if weekday in SCHEDULE_TIMES:
                for hour, minute in SCHEDULE_TIMES[weekday]:
                    slot = current_date.replace(hour=hour, minute=minute, second=0)
                    # Jen budoucí časy
                    if slot > now:
                        slots.append(slot)
        
        return sorted(slots)
    
    def save_schedule(self):
        """Uloží plán do JSON"""
        with open(SCHEDULE_FILE, 'w', encoding='utf-8') as f:
            json.dump(
                [
                    {
                        'nazev': item.get('nazev'),
                        'nastepka': item.get('nastepka'),
                        'barva_satu': item.get('barva_satu'),
                        'scheduled_time': item.get('scheduled_time'),
                        'status': 'scheduled'
                    }
                    for item in self.schedule
                ],
                f,
                ensure_ascii=False,
                indent=2
            )
        logger.info(f"Plán uložen: {SCHEDULE_FILE}")


class PinterestBot:
    """Selenium bot pro tvorbu pinů"""
    
    def __init__(self):
        self.driver = None
        self.wait = None
    
    def initialize(self):
        """Inicializace prohlížeče"""
        chrome_options = Options()
        # chrome_options.add_argument("--headless")  # Odkomentuj pro headless mód
        self.driver = webdriver.Chrome(options=chrome_options)
        self.wait = WebDriverWait(self.driver, 15)
        
        logger.info("Prohlížeč inicializován")
    
    def login_manual(self):
        """Manuální přihlášení"""
        self.driver.get(BASE_URL)
        logger.info("Přejděte na Pinterest a přihlaste se")
        logger.info("Až budete na mainbordu, vložte do terminálu: done")
        
        while input("Zadejte 'done' až se přihlásíte: ").lower() != 'done':
            pass
    
    def create_pin(self, pin_data):
        """Vytvoří pin s danou strukturou"""
        try:
            self.driver.get(PIN_CREATION_URL)
            time.sleep(3)
            
            logger.info(f"Vytvářím pin: {pin_data['nazev']}")
            
            # 1. Nahrání videa
            logger.info("  → Nahrávám video...")
            upload = self.wait.until(EC.presence_of_element_located((By.ID, "storyboard-upload-input")))
            upload.send_keys(os.path.abspath(pin_data['video_cesta']))
            time.sleep(5)
            
            # 2. Nadpis
            logger.info("  → Vyplňuji nadpis...")
            title = self.wait.until(EC.presence_of_element_located((By.ID, "storyboard-selector-title")))
            title.clear()
            title.send_keys(pin_data['nazev'])
            time.sleep(2)
            
            # 3. Popis (pokud není prázdný)
            if pd.notna(pin_data.get('popis')):
                logger.info("  → Vyplňuji popis...")
                try:
                    description = self.driver.find_element(By.ID, "description-input")
                    description.clear()
                    description.send_keys(pin_data['popis'])
                    time.sleep(1)
                except:
                    logger.warning("  ⚠️ Pole popisu nenalezeno, pokračuji...")
            
            # 4. Odkaz
            logger.info("  → Vyplňuji odkaz...")
            link = self.wait.until(EC.presence_of_element_located((By.ID, "WebsiteField")))
            link.clear()
            link.send_keys(pin_data['odkaz'])
            time.sleep(2)
            
            # 5. Nástěnka
            logger.info(f"  → Vybírám nástěnku: {pin_data['nastepka']}")
            print(f"🔍 Hledám tlačítko nástěnky...")
            board_btn = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//div[@data-test-id='board-dropdown-select-button']"))
            )
            print(f"✅ Tlačítko nástěnky nalezeno, klikám...")
            board_btn.click()
            time.sleep(5)
            
            # Vyber nástěnku podle jména z CSV
            print(f"🔍 Hledám nástěnku '{pin_data['nastepka']}'...")
            board_option = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, f"//div[@data-test-id='board-row-{pin_data['nastepka']}']"))
            )
            print(f"✅ Nástěnka nalezena, klikám...")
            board_option.click()
            time.sleep(5)
            
            # 6. Tagy
            logger.info("  → Přidávám tagy...")
            self._add_tags(pin_data.get('tagy', ''))
            
            # 7. Plánování (pokud je nastaveno)
            if pd.notna(pin_data.get('scheduled_time')):
                logger.info(f"  → Plánuji na: {pin_data['scheduled_time']}")
                self._schedule_pin(pin_data['scheduled_time'])
            
            # 8. Publikování
            logger.info("  → Publikuji pin...")
            publish_btn = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//div[@data-test-id='storyboard-creation-nav-done']"))
            )
            publish_btn.click()
            time.sleep(5)
            
            logger.info(f"✅ Pin úspěšně vytvořen: {pin_data['nazev']}")
            return True
            
        except Exception as e:
            logger.error(f"❌ Chyba při vytváření pinu: {e}")
            return False
    
    def _add_tags(self, tags_string):
        """Přidá tagy k pinu"""
        if pd.isna(tags_string):
            return
        
        # Rozdělí tagy podle středníku a vyčistí whitespace
        tags_list = [tag.strip() for tag in str(tags_string).split(';') if tag.strip()]
        
        for tag in tags_list:
            try:
                # Vyhledej input pole pro tagy - zkus různé selektory
                tag_input = None
                try:
                    tag_input = self.driver.find_element(By.ID, "tags-input")
                except:
                    try:
                        tag_input = self.driver.find_element(By.XPATH, "//input[@placeholder='Add a tag']")
                    except:
                        try:
                            tag_input = self.driver.find_element(By.XPATH, "//input[contains(@class, 'tags')]")
                        except:
                            pass
                
                if tag_input:
                    tag_input.send_keys(tag)
                    time.sleep(0.3)
                    # Stiskni Enter/Tab pro potvrzení tagu
                    tag_input.send_keys('\t')
                    time.sleep(0.3)
                    logger.info(f"     ✓ Tag přidán: {tag}")
                else:
                    logger.warning(f"  ⚠️ Pole pro tagy nenalezeno")
                    
            except Exception as e:
                logger.warning(f"  ⚠️ Nepodařilo se přidat tag '{tag}': {e}")
    
    def _schedule_pin(self, scheduled_time):
        """Naplánuje pin na specifický čas"""
        try:
            # Pokud je dostupná možnost plánování na Pinterestu
            schedule_btn = self.driver.find_element(By.XPATH, "//button[contains(text(), 'Naplánovat')]")
            schedule_btn.click()
            time.sleep(1)
            # Zde by bylo potřeba vyplnit datum a čas dle UI Pinterestu
        except:
            logger.info("  ℹ️ Plánování není dostupné, publikuji nyní")
    
    def close(self):
        """Uzavře prohlížeč"""
        if self.driver:
            self.driver.quit()
            logger.info("Prohlížeč uzavřen")


# ========== MAIN ==========

def main():
    """Hlavní funkce"""
    logger.info("=" * 60)
    logger.info("Pinterest Pin Automation System")
    logger.info("=" * 60)
    
    # 1. Validace
    logger.info("\n📋 VALIDACE DAT...")
    validator = PinterestValidator()
    if not validator.validate_file(CSV_FILE):
        logger.error("❌ Validace selhala! Zkontrolujte validation_report.txt")
        validator.print_report()
        return
    
    validator.print_report()
    
    # 2. Optimalizace plánu
    logger.info("\n📅 OPTIMALIZACE PLÁNU PUBLIKOVÁNÍ...")
    optimizer = ScheduleOptimizer(validator.data)
    optimizer.optimize_schedule()
    optimizer.save_schedule()
    
    logger.info(f"✅ Plán vytvořen pro {len(optimizer.schedule)} pinů")
    
    # 3. Tvorba pinů
    logger.info("\n🚀 SPOUŠTĚNÍ TVORBY PINŮ...")
    bot = PinterestBot()
    
    try:
        bot.initialize()
        bot.login_manual()
        
        created = 0
        for idx, pin_data in enumerate(optimizer.schedule, 1):
            logger.info(f"\n📌 Pin {idx}/{len(optimizer.schedule)}")
            if bot.create_pin(pin_data):
                created += 1
            time.sleep(3)  # Čekání mezi piny
        
        logger.info(f"\n✅ Hotovo! Vytvořeno {created}/{len(optimizer.schedule)} pinů")
        
    finally:
        bot.close()


if __name__ == "__main__":
    main()
