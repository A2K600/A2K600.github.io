"""
Generuje przykładowe, celowo "brudne" dane sprzedażowe z trzech oddziałów.

Symuluje typowe problemy, które spotyka się u prawdziwych klientów:
  - każdy oddział ma inne nazwy kolumn,
  - daty w różnych formatach (2026-03-05, 05.03.2026, 5/3/2026),
  - kwoty jako tekst z "zł", spacjami i przecinkiem dziesiętnym,
  - literówki i różna wielkość liter w nazwach produktów,
  - zduplikowane wiersze i puste wiersze,
  - jeden oddział wysyła CSV ze średnikami zamiast Excela.

Uruchomienie:  python generuj_dane_przykladowe.py
"""
import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

random.seed(42)

FOLDER = Path(__file__).parent / "dane_wejsciowe"
FOLDER.mkdir(exist_ok=True)

PRODUKTY = {
    "Kawa ziarnista 1kg": 89.0,
    "Kawa mielona 250g": 24.0,
    "Herbata zielona": 18.5,
    "Ekspres przelewowy": 349.0,
    "Młynek ręczny": 129.0,
    "Kubek termiczny": 59.0,
}

# Warianty nazw, jakie "wpisują" pracownicy oddziałów
LITEROWKI = {
    "Kawa ziarnista 1kg": ["kawa ziarnista 1kg", "Kawa ziarnista 1 kg", "KAWA ZIARNISTA 1KG"],
    "Kawa mielona 250g": ["Kawa mielona 250 g", "kawa mielona 250g"],
    "Herbata zielona": ["herbata zielona", "Herbata Zielona "],
    "Ekspres przelewowy": ["Ekspres przelew.", "ekspres przelewowy"],
    "Młynek ręczny": ["Mlynek reczny", "młynek ręczny"],
    "Kubek termiczny": ["Kubek termiczny ", "kubek termiczny"],
}

HANDLOWCY = {
    "Poznań": ["Anna Nowak", "Piotr Kowalski", "Marta Wiśniewska"],
    "Warszawa": ["Tomasz Lewandowski", "Kasia Zielińska", "Jan Wójcik", "Ola Kamińska"],
    "Wrocław": ["Michał Szymański", "Ewa Dąbrowska"],
}

START = date(2026, 1, 1)
DNI = 273  # styczeń–wrzesień 2026


def nazwa_z_bledami(produkt: str) -> str:
    if random.random() < 0.3:
        return random.choice(LITEROWKI[produkt])
    return produkt


def transakcje(oddzial: str, n: int):
    wiersze = []
    for _ in range(n):
        d = START + timedelta(days=random.randint(0, DNI - 1))
        produkt = random.choice(list(PRODUKTY))
        ilosc = random.choices([1, 2, 3, 4, 5, 10], weights=[40, 25, 15, 8, 7, 5])[0]
        cena = PRODUKTY[produkt]
        # sezonowość: więcej sprzedaży kawy jesienią i zimą
        if d.month in (1, 2, 9) and "Kawa" in produkt and random.random() < 0.4:
            ilosc += random.randint(1, 3)
        wiersze.append((d, nazwa_z_bledami(produkt), ilosc, cena, random.choice(HANDLOWCY[oddzial])))
    return wiersze


# --- Poznań: porządny Excel, ale nagłówki po swojemu i duplikaty --------------
poz = transakcje("Poznań", 420)
df_poz = pd.DataFrame(poz, columns=["Data sprzedaży", "Produkt", "Ilość", "Cena jedn.", "Sprzedawca"])
df_poz["Wartość"] = df_poz["Ilość"] * df_poz["Cena jedn."]
df_poz = pd.concat([df_poz, df_poz.sample(12, random_state=1)])  # duplikaty
df_poz.to_excel(FOLDER / "sprzedaz_poznan_2026.xlsx", index=False)

# --- Warszawa: daty jako tekst DD.MM.YYYY, kwoty jako tekst "1 234,50 zł" ------
waw = transakcje("Warszawa", 610)
df_waw = pd.DataFrame(
    {
        "DATA": [d.strftime("%d.%m.%Y") for d, *_ in waw],
        "NAZWA TOWARU": [p for _, p, *_ in waw],
        "SZT": [q for _, _, q, *_ in waw],
        "KWOTA": [f"{q * c:,.2f} zł".replace(",", " ").replace(".", ",") for _, _, q, c, _ in waw],
        "HANDLOWIEC": [h for *_, h in waw],
    }
)
# puste wiersze w środku (ktoś zostawił przerwy)
puste = pd.DataFrame([[None] * 5] * 6, columns=df_waw.columns)
df_waw = pd.concat([df_waw.iloc[:200], puste.iloc[:3], df_waw.iloc[200:450], puste.iloc[3:], df_waw.iloc[450:]])
df_waw.to_excel(FOLDER / "WAW_sprzedaz.xlsx", index=False, sheet_name="Arkusz1")

# --- Wrocław: CSV ze średnikami, daty D/M/YYYY, brak kolumny wartości ----------
wro = transakcje("Wrocław", 300)
df_wro = pd.DataFrame(
    {
        "dt": [f"{d.day}/{d.month}/{d.year}" for d, *_ in wro],
        "towar": [p for _, p, *_ in wro],
        "ilosc": [q for _, _, q, *_ in wro],
        "cena": [str(c).replace(".", ",") for *_, c, _ in wro],
        "osoba": [h for *_, h in wro],
    }
)
# kilka wierszy z błędną ilością (0 albo pusta) – do odrzucenia
for i in random.sample(range(len(df_wro)), 5):
    df_wro.loc[i, "ilosc"] = random.choice([0, None])
df_wro.to_csv(FOLDER / "wroclaw-export.csv", sep=";", index=False, encoding="utf-8-sig")

print(f"Zapisano przykładowe dane w: {FOLDER}")
for f in sorted(FOLDER.iterdir()):
    print("  -", f.name)
