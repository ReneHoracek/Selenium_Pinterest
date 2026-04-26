"""
Monitorovací dashboard pro Pinterest automation
Umožňuje Pavlě vidět, jaké piny jsou naplánované a jejich stav
"""

import json
import os
from datetime import datetime
from pathlib import Path

def print_dashboard():
    """Vytiskne dashboard se stavem pinů"""
    
    print("\n" + "=" * 70)
    print(" " * 15 + "DASHBOARD PINTEREST AUTOMATION")
    print("=" * 70 + "\n")
    
    # 1. Validační zpráva
    if os.path.exists('validation_report.txt'):
        print("📋 VALIDACE DAT")
        print("-" * 70)
        with open('validation_report.txt', 'r', encoding='utf-8') as f:
            content = f.read()
            # Vezmi jen první 500 znaků
            if len(content) > 500:
                print(content[:500] + "\n[... více v validation_report.txt]")
            else:
                print(content)
        print()
    
    # 2. Plán publikování
    if os.path.exists('schedule.json'):
        print("📅 PLÁN PUBLIKOVÁNÍ")
        print("-" * 70)
        with open('schedule.json', 'r', encoding='utf-8') as f:
            schedule = json.load(f)
        
        if schedule:
            for idx, pin in enumerate(schedule[:10], 1):  # Ukaž jen prvních 10
                scheduled_time = pin.get('scheduled_time', 'N/A')
                status = pin.get('status', 'unknown')
                
                # Formátuj čas
                try:
                    dt = datetime.fromisoformat(scheduled_time)
                    time_str = dt.strftime("%d.%m.%Y %H:%M")
                except:
                    time_str = scheduled_time
                
                print(f"{idx}. {pin['nazev'][:40]}")
                print(f"   📌 Nástěnka: {pin['nastepka']}")
                print(f"   🎨 Barva: {pin['barva_satu']}")
                print(f"   📍 Čas: {time_str}")
                print(f"   ✓ Stav: {status}")
                print()
            
            if len(schedule) > 10:
                print(f"... a dalších {len(schedule) - 10} pinů\n")
        else:
            print("Žádné piny nejsou naplánované.\n")
    
    # 3. Log soubor
    if os.path.exists('pinterest_automation.log'):
        print("📝 POSLEDNÍ AKTIVITY")
        print("-" * 70)
        with open('pinterest_automation.log', 'r', encoding='utf-8') as f:
            lines = f.readlines()
            # Poslední 5 řádků
            for line in lines[-5:]:
                print(line.rstrip())
        print()
    
    # 4. Souhrnné statistiky
    if os.path.exists('schedule.json'):
        with open('schedule.json', 'r', encoding='utf-8') as f:
            schedule = json.load(f)
        
        print("📊 STATISTIKA")
        print("-" * 70)
        print(f"Celkem naplánovaných pinů: {len(schedule)}")
        
        # Počet pinů po dnech v týdnu
        days = {
            0: "PO", 1: "ÚT", 2: "ST", 3: "ČT", 
            4: "PÁ", 5: "SO", 6: "NE"
        }
        day_counts = {day: 0 for day in days.values()}
        
        for pin in schedule:
            try:
                dt = datetime.fromisoformat(pin.get('scheduled_time', ''))
                day_name = days[dt.weekday()]
                day_counts[day_name] += 1
            except:
                pass
        
        print("Rozložení dle dnů: ", end="")
        for day in ["PO", "ÚT", "ST", "ČT", "PÁ", "SO", "NE"]:
            count = day_counts.get(day, 0)
            print(f"{day}:{count} ", end="")
        print("\n")
    
    print("=" * 70)
    print("💡 Tip: Spusť 'python pinterest_automation.py' pro vytvoření pinů")
    print("=" * 70 + "\n")


if __name__ == "__main__":
    print_dashboard()
