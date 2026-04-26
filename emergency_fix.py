"""
Nouzové řešení - vytvoření Data.xlsx od nuly z tvých dat
Spusť v té samé složce
"""
import pandas as pd
import os

# Zkus najít jakýkoli datový soubor
files_to_try = ['Data.csv', 'Data.txt', 'data.csv', 'data.txt', 'Data.xlsx']

data = None
for fname in files_to_try:
    if os.path.exists(fname):
        try:
            if fname.endswith('.csv') or fname.endswith('.txt'):
                # Zkus s různými oddělovači
                for sep in [',', ';', '\t', '|']:
                    try:
                        data = pd.read_csv(fname, sep=sep, encoding='utf-8')
                        if len(data.columns) > 2:
                            print(f"✓ Načten {fname} s oddělovačem '{sep}'")
                            break
                    except:
                        pass
            else:
                data = pd.read_excel(fname, engine='openpyxl')
                print(f"✓ Načten {fname}")
            
            if data is not None:
                break
        except Exception as e:
            print(f"✗ {fname}: {e}")

if data is None:
    print("❌ Nepodařilo se načíst žádný datový soubor!")
    print("Dostupné soubory:", [f for f in os.listdir('.') if f.endswith(('.csv', '.txt', '.xlsx', '.xls'))])
    exit(1)

print(f"Sloupce před: {list(data.columns)}")

# Vyčisti sloupce
data.columns = data.columns.str.strip().str.replace('\n', '').str.replace('\r', '')

# Odstraň duplikátní sloupce
data = data.loc[:, ~data.columns.duplicated(keep='first')]

print(f"Sloupce po: {list(data.columns)}")
print(f"Řádků: {len(data)}")

# Zkontroluj povinné sloupce
required = ['nazev', 'video_cesta', 'odkaz', 'nastepka']
for col in required:
    if col not in data.columns:
        print(f"⚠️ Chybí sloupec: {col}")

# Přidej chybějící sloupce
if 'tagy' not in data.columns:
    data['tagy'] = 'účetnictví'

if 'barva_satu' not in data.columns:
    data['barva_satu'] = 'Různobarevná'

# Ulož správně
output_file = 'Data_Clean.xlsx'
data.to_excel(output_file, index=False, engine='openpyxl', sheet_name='Piny')
print(f"\n✅ Čistý soubor: {output_file}")
print(f"   Přejmenuj na Data.xlsx:")
print(f"   mv {output_file} Data.xlsx")
