"""
Převod Data.xlsx (pokud je poškozený) na správný Excel formát
Spusť v té samé složce jako tvůj Data.xlsx
"""
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# Zkus načíst existující data - nejdřív jako CSV, pak XLSX
data = None
try:
    # Zkus XLSX
    data = pd.read_excel('Data.xlsx', engine='openpyxl')
    print("✓ Načten XLSX formát")
except:
    try:
        # Zkus CSV
        data = pd.read_csv('Data.xlsx', encoding='utf-8')
        print("✓ Načten CSV formát")
    except:
        try:
            # Zkus s jiným oddělovačem
            data = pd.read_csv('Data.xlsx', sep='\t', encoding='utf-8')
            print("✓ Načten TSV (tab-separated) formát")
        except Exception as e:
            print(f"❌ Nepodařilo se načíst data: {e}")
            exit(1)

# Vyčisti sloupce
data.columns = data.columns.str.strip().str.replace('\n', '').str.replace('\r', '')
print(f"Sloupce: {list(data.columns)}")
print(f"Řádků: {len(data)}")

# Ulož do správného XLSX formátu
data.to_excel('Data_Fixed.xlsx', index=False, engine='openpyxl')
print("\n✅ Nový soubor vytvořen: Data_Fixed.xlsx")
print("   Přejmenuj ho na Data.xlsx a zkus znovu:")
print("   mv Data_Fixed.xlsx Data.xlsx")
