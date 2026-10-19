# 👥 Segmentacja klientów sklepu internetowego (RFM)

> **EN:** Customer segmentation of an online shop using RFM (Recency, Frequency, Monetary). Python script scores 1,800 customers, assigns them to 9 segments with marketing recommendations, and generates an HTML report with charts plus an Excel file with a ready-to-use win-back campaign list.

## Problem

Sklep internetowy ma 1 800 klientów i ok. 9 500 zamówień z dwóch lat. Wszystkim wysyła te same newslettery i rabaty. Przychód spada, a nie wiadomo dlaczego ani do kogo kierować kampanie.

## Rozwiązanie

Analiza **RFM**, czyli trzy pytania o każdego klienta:
- **R (Recency):** kiedy ostatnio kupił?
- **F (Frequency):** jak często kupuje?
- **M (Monetary):** ile łącznie wydał?

Każdy wymiar dostaje ocenę 1–5, a z kombinacji ocen powstaje 9 segmentów: od „Najlepszych klientów” po „Utraconych”. Każdy segment ma konkretną rekomendację działań.

## Wyniki

Kilka wniosków z [raportu](index.html):
- **19% klientów przynosi 40% przychodu**, a jednocześnie **16% najcenniejszych klientów przestało kupować**. Średnio nie kupili nic od ponad roku.
- Średni przychód miesięczny **spadł o 53%**, co pokrywa się z odejściem tej grupy.
- Excel zawiera **gotową listę klientów do kampanii odzyskującej**, posortowaną według tego, ile wydali.

## Pliki

| Plik | Co to jest |
|---|---|
| [`index.html`](index.html) | raport z wnioskami, wykresami i rekomendacjami |
| [`segmenty_klientow.xlsx`](segmenty_klientow.xlsx) | podsumowanie segmentów (formuły), lista klientów z ocenami, lista do kampanii |
| `analiza_rfm.py` | cała analiza i generowanie raportu |
| `generuj_dane.py` | generator przykładowych danych |

## Uruchomienie

```bash
pip install pandas openpyxl
python generuj_dane.py   # tylko do demo
python analiza_rfm.py    # tworzy index.html i segmenty_klientow.xlsx
```

Dla prawdziwego sklepu wystarczy eksport zamówień z kolumnami: id zamówienia, id klienta, data, wartość. Większość platform (WooCommerce, Shopify, PrestaShop) ma taki eksport.

## Technologie

Python · pandas · openpyxl · wykresy SVG generowane w Pythonie

---
*Dane są fikcyjne, wygenerowane skryptem `generuj_dane.py`.*
