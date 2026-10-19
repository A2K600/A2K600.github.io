"""
Eksportuje oczyszczone dane z projektu "raport-sprzedazy" do pliku dane.js,
z którego korzysta dashboard (index.html).

Uruchomienie (z folderu dashboard):  python eksport_danych.py
"""
import json
import sys
from pathlib import Path

PROJEKT_RAPORT = Path(__file__).resolve().parent.parent / "raport-sprzedazy"
sys.path.insert(0, str(PROJEKT_RAPORT))

import pandas as pd  # noqa: E402
import raport  # noqa: E402  (używamy tej samej logiki czyszczenia)

pliki = sorted(p for p in raport.FOLDER_WEJSCIE.glob("*") if p.suffix.lower() in {".xlsx", ".xls", ".csv"})
dane = pd.concat([raport.oczysc(p) for p in pliki], ignore_index=True).sort_values("data")

rekordy = [
    [r.data.strftime("%Y-%m-%d"), r.oddzial, r.produkt, int(r.ilosc), round(float(r.ilosc * r.cena), 2), r.handlowiec]
    for r in dane.itertuples()
]
wyjscie = Path(__file__).parent / "dane.js"
wyjscie.write_text(
    "// Wygenerowane przez eksport_danych.py – nie edytuj ręcznie.\n"
    "// Kolumny: data, oddział, produkt, ilość, wartość (zł), handlowiec\n"
    f"window.DANE = {json.dumps(rekordy, ensure_ascii=False, separators=(',', ':'))};\n",
    encoding="utf-8",
)
print(f"Zapisano {len(rekordy)} transakcji do {wyjscie.name}")
