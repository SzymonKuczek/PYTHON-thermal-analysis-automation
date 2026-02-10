import openpyxl
import os

# --- KONFIGURACJA ŚCIEŻEK ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(BASE_DIR)
DATA_DIR = os.path.join(PROJECT_ROOT, 'data')

os.makedirs(DATA_DIR, exist_ok=True)

def stworz_plik_projektowy(nazwa_pliku, warstwy):
    sciezka_zapisu = os.path.join(DATA_DIR, nazwa_pliku)
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sciana"

    ws.append(["Nazwa", "Grubosc [m]"])
    
    for material, grubosc in warstwy:
        ws.append([material, grubosc])
        
    wb.save(sciezka_zapisu)
    print(f"[GENERATOR] Utworzono: {sciezka_zapisu}")

# --- SCENARIUSZE ---
scenariusz_1 = [
    ("Tynk cementowo-wapienny", 0.02),
    ("Cegła pełna", 0.50),
    ("Tynk cementowo-wapienny", 0.02)
]

scenariusz_2 = [
    ("Tynk cementowo-wapienny", 0.015),
    ("Pustak ceramiczny", 0.25),
    ("Styropian EPS", 0.15),
    ("Tynk cienkowarstwowy", 0.01)
]

scenariusz_3 = [
    ("Beton zwykły", 0.20),
    ("Styropian Grafitowy", 0.25),
    ("Tynk cienkowarstwowy", 0.01)
]

if __name__ == "__main__":
    print(f"[INFO] Folder danych: {DATA_DIR}")
    stworz_plik_projektowy("projekt_kamienica.xlsx", scenariusz_1)
    stworz_plik_projektowy("projekt_standard.xlsx", scenariusz_2)
    stworz_plik_projektowy("projekt_pasywny.xlsx", scenariusz_3)