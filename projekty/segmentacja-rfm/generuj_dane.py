"""
Generuje przykładową historię zamówień sklepu internetowego (2 lata).

Klienci mają różne "charaktery" zakupowe – stali, nowi, jednorazowi,
odchodzący – tak, żeby analiza miała co wykryć. Dane są fikcyjne.

Uruchomienie:  python generuj_dane.py   →  zamowienia.csv
"""
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

rng = np.random.default_rng(42)
START, KONIEC = date(2024, 10, 1), date(2026, 9, 30)
DNI = (KONIEC - START).days

TYPY = {                    # udział, zamówień (śr.), wartość koszyka (śr.), kiedy aktywny (od, do jako ułamek okresu)
    "staly":        (0.12, 14, 165, (0.0, 1.0)),
    "regularny":    (0.18, 6, 120, (0.1, 1.0)),
    "nowy":         (0.15, 1.6, 110, (0.85, 1.0)),
    "jednorazowy":  (0.25, 1, 95, (0.0, 0.8)),
    "odchodzacy":   (0.15, 7, 150, (0.0, 0.6)),
    "zgubiony_vip": (0.05, 12, 260, (0.0, 0.55)),
    "okazjonalny":  (0.10, 2.5, 100, (0.2, 0.9)),
}
N_KLIENTOW = 1800
wiersze = []
nr = 100000
typy = rng.choice(list(TYPY), size=N_KLIENTOW, p=[t[0] for t in TYPY.values()])
for i, typ in enumerate(typy, start=1):
    _, sr_zam, sr_wart, (od, do) = TYPY[typ]
    n = max(1, int(rng.poisson(sr_zam))) if typ != "jednorazowy" else 1
    a, b = int(DNI * od), int(DNI * do)
    if typ in ("odchodzacy", "zgubiony_vip"):
        b = int(DNI * rng.uniform(do - 0.25, do))
    dni = np.sort(rng.integers(a, max(a + 1, b), size=n))
    for d in dni:
        nr += 1
        wartosc = round(float(rng.gamma(4, sr_wart / 4)), 2)
        wiersze.append((f"Z{nr}", f"K{i:05d}", START + timedelta(days=int(d)), max(wartosc, 19.0)))

df = pd.DataFrame(wiersze, columns=["id_zamowienia", "id_klienta", "data", "wartosc_zl"]).sort_values("data")
df.to_csv(Path(__file__).parent / "zamowienia.csv", index=False)
print(f"Zapisano {len(df)} zamówień od {df.id_klienta.nunique()} klientów → zamowienia.csv")
