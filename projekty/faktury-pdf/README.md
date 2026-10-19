# 🧾 Faktury PDF → zestawienie w Excelu

> **EN:** Python tool that reads a folder of PDF invoices from different suppliers (each with its own layout), extracts supplier, tax ID, invoice number, date and net/VAT/gross amounts, validates them (sum check, tax ID checksum, duplicates, scanned PDFs without text) and builds an Excel summary with a "needs review" list.

## Problem

Biuro co miesiąc dostaje kilkadziesiąt faktur PDF od różnych dostawców. Każdy dostawca ma inny układ faktury:

| Dostawca | Numer | Data | Kwoty |
|---|---|---|---|
| Hurtownia Arabika | `Faktura VAT nr FV/2026/03/013` | `2026-03-19` | `1 068,87 zł` |
| EKO-BOX | `FAKTURA NR 6/03/2026` | `27.03.2026` | `1052.88 PLN` w jednej linii |
| Barista-Tech | `Numer faktury: BT-2026-0043` | `10 kwietnia 2026` | `Do zapłaty: 578,10 PLN` |

Ktoś przepisuje je ręcznie do Excela. Zajmuje to godziny i łatwo przeoczyć błąd na fakturze.

## Rozwiązanie

`wyciagnij_dane.py` dla każdego PDF-a:

1. **Odczytuje tekst** (pdfplumber). Dane sprzedawcy bierze z lewej kolumny strony, żeby nie pomylić ich z nabywcą.
2. **Wyciąga dane** wzorcami dopasowanymi do różnych etykiet i formatów dat oraz kwot. Nowego dostawcę dodaje się jedną linijką.
3. **Sprawdza poprawność:**
   - ✅ czy netto + VAT = brutto,
   - ✅ czy NIP ma poprawną sumę kontrolną,
   - ✅ czy ta sama faktura nie przyszła dwa razy (np. raz z maila, raz ze skanu),
   - ✅ czy PDF ma tekst. Skan bez tekstu jest oznaczany jako „wymaga OCR”.
4. **Tworzy Excel** z trzema arkuszami:
   - **Podsumowanie:** koszty netto/VAT/brutto wg dostawców i miesięcy (formuły Excela),
   - **Faktury:** pełna lista, błędne wiersze podświetlone na czerwono,
   - **Do sprawdzenia:** konkretna lista problemów do wyjaśnienia z dostawcą.

## Efekt

```
Przetworzono 24 faktur: 20 OK, 4 do sprawdzenia
  ! EKOBOX faktura 8.pdf: Błędny NIP 8970002611 (zła suma kontrolna)
  ! arabika_014.pdf: Netto + VAT = 2027.04, a na fakturze brutto 2037.04
  ! skan_mail_46.pdf: Duplikat – ta sama faktura co w pliku BT-2026-0046.pdf
  ! skan_paragon_0912.pdf: Brak tekstu w PDF (skan) – wymaga OCR albo ręcznego przepisania
Gotowe → zestawienie_faktur.xlsx
```

Kilka godzin przepisywania zamienia się w kilka sekund, a błędy na fakturach wychodzą, zanim trafią do księgowości.

📥 [Pobierz wynikowy Excel](zestawienie_faktur.xlsx) · 📄 [Przykładowa faktura](faktury/arabika_011.pdf)

## Uruchomienie

```bash
pip install pdfplumber openpyxl reportlab pillow
python generuj_faktury.py   # tylko do demo – tworzy 24 przykładowe faktury
python wyciagnij_dane.py    # tworzy zestawienie_faktur.xlsx
```

## Technologie

Python · pdfplumber (odczyt PDF) · wyrażenia regularne · openpyxl (Excel z formułami i formatowaniem warunkowym) · reportlab (generowanie przykładowych faktur)

---
*Wszystkie faktury są fikcyjne, oznaczone jako „dokument przykładowy” i wygenerowane skryptem `generuj_faktury.py`.*
