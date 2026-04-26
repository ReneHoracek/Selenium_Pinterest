"""
Konvertor z CSV na nový Excel formát
Hodí se, pokud máš starý CSV soubor a chceš ho přepracovat
"""

import pandas as pd
import os

def convert_csv_to_excel():
    """Konvertuje stávající CSV na nový Excel formát"""
    
    # Pokus se načíst starý CSV
    csv_files = [f for f in os.listdir('.') if f.endswith('.csv')]
    
    if not csv_files:
        print("❌ Žádný CSV soubor nenalezen v aktuální složce")
        return
    
    print(f"Nalezeny CSV soubory: {csv_files}")
    csv_file = csv_files[0]
    
    try:
        df = pd.read_csv(csv_file, encoding='utf-8')
        print(f"✅ Soubor načten: {csv_file}")
        print(f"   Počet řádků: {len(df)}")
        print(f"   Sloupce: {', '.join(df.columns)}")
        
        # Mapování starých názvů na nové
        mapping = {
            'nazev': 'nazev',
            'video_cesta': 'video_cesta',
            'odkaz': 'odkaz',
            'nastepka': 'nastepka',
        }
        
        # Zkontroluj dostupné sloupce
        available_mapping = {}
        for old, new in mapping.items():
            if old in df.columns:
                available_mapping[old] = new
        
        # Vezmi jen dostupné sloupce
        df_new = df[list(available_mapping.keys())].copy()
        df_new.columns = [available_mapping[col] for col in df_new.columns]
        
        # Přidej nové sloupce s výchozími hodnotami
        if 'popis' not in df_new.columns:
            df_new['popis'] = ''
        
        if 'tagy' not in df_new.columns:
            df_new['tagy'] = 'účetnictví'
        
        if 'barva_satu' not in df_new.columns:
            df_new['barva_satu'] = 'Různobarevná'
        
        # Ulož do Excelu
        output_file = 'Data_Converted.xlsx'
        df_new.to_excel(output_file, index=False, sheet_name='Piny')
        
        print(f"\n✅ Konverze úspěšná!")
        print(f"   Nový soubor: {output_file}")
        print(f"   Řádků: {len(df_new)}")
        
        # Upozornění
        print("\n⚠️  DŮLEŽITÉ:")
        print("   • Zkontroluj nový soubor - může potřebovat ruční úpravy")
        print("   • Vyplň barvu šatů pro každé video (důležité pro střídání)")
        print("   • Rozšiř tagy dle pokynů v README")
        print("   • Přejmenuj na 'Data.xlsx' až budeš připravená")
        
    except Exception as e:
        print(f"❌ Chyba při konverzi: {e}")


if __name__ == "__main__":
    convert_csv_to_excel()
