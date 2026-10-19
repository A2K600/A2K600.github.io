# 📊 Automatyczny raport sprzedaży z wielu plików Excel

> **EN:** Python tool that merges messy sales files from several branches (different column names, date formats, typos, duplicates, CSV/Excel), cleans them and generates a formatted Excel report with KPIs, charts, a sales-rep ranking and a full log of every correction — in one command.

## Problem

Firma ma kilka oddziałów. Każdy co miesiąc wysyła sprzedaż w swoim formacie:

| Oddział | Format | Co jest nie tak |
|---|---|---|
| Poznań | Excel | inne nazwy kolumn, zduplikowane wiersze |
| Warszawa | Excel | daty jako tekst `05.03.2026`, kwoty jako tekst `1 234,50 zł`, puste wiersze, brak ceny jednostkowej |
| Wrocław | CSV ze średnikami | daty `5/3/2026`, przecinek dziesiętny, brak kolumny wartości, wiersze z zerową ilością |

Wszędzie literówki w nazwach produktów (`Mlynek reczny`, `KAWA ZIARNISTA 1KG`, `Ekspres przelew.`).

Ktoś w biurze **ręcznie skleja to co miesiąc przez kilka godzin**, a i tak zdarzają się błędy.

## Rozwiązanie

Jeden skrypt, który:

1. **Wczytuje dowolną liczbę plików** `.xlsx` / `.csv` z folderu – sam rozpoznaje kolumny po synonimach i separator CSV.
2. **Czyści dane** – ujednolica daty i kwoty, dopasowuje literówki do katalogu produktów, usuwa duplikaty i puste wiersze, dopełnia brakujące kolumny.
3. **Generuje raport Excel** z arkuszami:
   - **Podsumowanie** – KPI, sprzedaż wg oddziałów, przychód miesięczny z wykresem, ranking produktów z wykresem,
   - **Handlowcy** – ranking sprzedawców,
   - **Dane** – czysta, przefiltrowana tabela,
   - **Log czyszczenia** – co dokładnie zostało poprawione i w ilu wierszach.

Wszystkie liczby w raporcie to **formuły Excela**, a nie wklejone wartości – raport da się dalej edytować, a każdą liczbę można sprawdzić.

## Efekt

⏱ Z kilku godzin ręcznej pracy → **kilka sekund**.
✅ Zero błędów przy przepisywaniu, pełna przejrzystość zmian.

```
Wczytuję 3 plik(i):
  ✓ WAW_sprzedaz.xlsx                604 wierszy  (Warszawa)
  ✓ sprzedaz_poznan_2026.xlsx        417 wierszy  (Poznań)
  ✓ wroclaw-export.csv               292 wierszy  (Wrocław)

Gotowe → raport_sprzedazy.xlsx  (1313 transakcji, przychód 374 384,00 zł)
```

## Uruchomienie

```bash
pip install pandas openpyxl
python generuj_dane_przykladowe.py   # tylko do demo – tworzy przykładowe "brudne" pliki
python raport.py                     # tworzy raport_sprzedazy.xlsx
```

## Dostosowanie pod nową firmę

Wszystko, co zależy od klienta, jest na górze `raport.py` w sekcji **KONFIGURACJA**:

- `NAZWA_FIRMY` – nagłówek raportu,
- `ODDZIALY` – jak rozpoznać oddział po nazwie pliku,
- `KATALOG` – oficjalne nazwy produktów,
- `KOLUMNY` – synonimy nazw kolumn.

## Technologie

Python · pandas · openpyxl (natywne wykresy i formuły Excela) · difflib (dopasowanie literówek)

---
*Dane w repozytorium są fikcyjne i wygenerowane skryptem `generuj_dane_przykladowe.py`.*
