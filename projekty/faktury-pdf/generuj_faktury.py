"""
Generuje przykładowe faktury PDF od trzech fikcyjnych dostawców.

Każdy dostawca ma INNY układ faktury (inne etykiety, format daty i kwot),
tak jak w prawdziwym życiu. Celowo dodane są też problemy do wykrycia:
  - faktura, na której brutto nie zgadza się z netto + VAT,
  - faktura z błędnym NIP-em (zła suma kontrolna),
  - ta sama faktura przysłana dwa razy (duplikat),
  - "skan" – PDF bez warstwy tekstowej (sam obrazek).

Wszystkie dokumenty są oznaczone jako PRZYKŁADOWE – dane są fikcyjne.

Uruchomienie:  python generuj_faktury.py
"""
import random
from datetime import date, timedelta
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

random.seed(7)
FOLDER = Path(__file__).parent / "faktury"
FOLDER.mkdir(exist_ok=True)

# --- czcionka z polskimi znakami (Linux / macOS) ---
KANDYDACI = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/Library/Fonts/Arial Unicode.ttf",
    "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
    "/System/Library/Fonts/Supplemental/Arial.ttf",
    "C:/Windows/Fonts/arial.ttf",
]
FONT_PATH = next(p for p in KANDYDACI if Path(p).exists())
pdfmetrics.registerFont(TTFont("F", FONT_PATH))
BOLD_PATH = FONT_PATH.replace("DejaVuSans.ttf", "DejaVuSans-Bold.ttf")
pdfmetrics.registerFont(TTFont("FB", BOLD_PATH if Path(BOLD_PATH).exists() else FONT_PATH))


def nip_z_suma(prefix9: str) -> str:
    wagi = [6, 5, 7, 2, 3, 4, 5, 6, 7]
    k = sum(int(c) * w for c, w in zip(prefix9, wagi)) % 11
    if k == 10:
        return nip_z_suma(str(int(prefix9) + 1).zfill(9))
    return prefix9 + str(k)


def zl(v, styl):
    if styl == "pl":   # 1 234,56
        return f"{v:,.2f}".replace(",", " ").replace(".", ",")
    return f"{v:.2f}"  # 1234.56


NABYWCA = ("Kawiarnia & Palarnia „Ziarno” sp. z o.o.", "ul. Przykładowa 1, 60-001 Poznań", nip_z_suma("778000123"))

DOSTAWCY = {
    "arabika": dict(nazwa="Hurtownia Kawy Arabika sp. z o.o.", adres="ul. Fikcyjna 8, 02-001 Warszawa",
                    nip=nip_z_suma("525000417"),
                    towary=[("Kawa ziarnista Brazylia 1 kg", 62.0), ("Kawa ziarnista Etiopia 1 kg", 79.0),
                            ("Kawa ziarnista Kolumbia 1 kg", 71.0)]),
    "ekobox": dict(nazwa="Pakownia EKO-BOX s.c.", adres="ul. Testowa 3, 50-001 Wrocław",
                   nip=nip_z_suma("897000261"),
                   towary=[("Kubek papierowy 250 ml (100 szt.)", 38.0), ("Pokrywka do kubka (100 szt.)", 21.0),
                           ("Torba kraft na kawę 1 kg (50 szt.)", 64.0)]),
    "barista": dict(nazwa="Serwis Ekspresów Barista-Tech", adres="ul. Demo 15, 61-001 Poznań",
                    nip=nip_z_suma("779000585"),
                    towary=[("Przegląd ekspresu ciśnieniowego", 280.0), ("Wymiana uszczelek grupy", 140.0),
                            ("Odkamienianie ekspresu", 190.0)]),
}
MIESIACE = ["stycznia", "lutego", "marca", "kwietnia", "maja", "czerwca", "lipca", "sierpnia", "września"]


def pozycje(dostawca):
    out = []
    for nazwa, cena in random.sample(DOSTAWCY[dostawca]["towary"], k=random.randint(1, 3)):
        ilosc = random.randint(1, 12) if dostawca != "barista" else 1
        out.append((nazwa, ilosc, cena, round(ilosc * cena, 2)))
    return out


def znak_wodny(c):
    c.saveState()
    c.setFont("FB", 9)
    c.setFillGray(0.55)
    c.drawString(20 * mm, 10 * mm, "DOKUMENT PRZYKŁADOWY – DANE FIKCYJNE, WYGENEROWANE DO PORTFOLIO")
    c.restoreState()


def strony_firm(c, y, sprzedawca, nip_sprzedawcy):
    c.setFont("FB", 10); c.drawString(20 * mm, y, "Sprzedawca"); c.drawString(110 * mm, y, "Nabywca")
    c.setFont("F", 10)
    for i, (l, r) in enumerate([(sprzedawca["nazwa"], NABYWCA[0]), (sprzedawca["adres"], NABYWCA[1]),
                                (f"NIP: {nip_sprzedawcy}", f"NIP: {NABYWCA[2]}")]):
        c.drawString(20 * mm, y - (i + 1) * 5 * mm, l); c.drawString(110 * mm, y - (i + 1) * 5 * mm, r)


def tabela(c, y, poz, styl):
    c.setFont("FB", 9)
    for x, t in [(20, "Lp."), (30, "Nazwa"), (125, "Ilość"), (145, "Cena netto"), (172, "Wartość netto")]:
        c.drawString(x * mm, y, t)
    c.line(20 * mm, y - 2 * mm, 190 * mm, y - 2 * mm)
    c.setFont("F", 9)
    for i, (n, q, cena, w) in enumerate(poz, 1):
        yy = y - (i * 6 + 2) * mm
        c.drawString(20 * mm, yy, str(i)); c.drawString(30 * mm, yy, n); c.drawRightString(135 * mm, yy, str(q))
        c.drawRightString(165 * mm, yy, zl(cena, styl)); c.drawRightString(190 * mm, yy, zl(w, styl))
    return y - (len(poz) * 6 + 10) * mm


def faktura_arabika(plik, nr, d, poz, blad_brutto=False):
    s = DOSTAWCY["arabika"]; c = canvas.Canvas(str(plik), pagesize=A4)
    c.setFont("FB", 16); c.drawString(20 * mm, 270 * mm, f"Faktura VAT nr {nr}")
    c.setFont("F", 10); c.drawString(20 * mm, 262 * mm, f"Data wystawienia: {d:%Y-%m-%d}")
    c.drawString(20 * mm, 257 * mm, f"Data sprzedaży: {d:%Y-%m-%d}")
    strony_firm(c, 242 * mm, s, s["nip"])
    y = tabela(c, 210 * mm, poz, "pl")
    netto = round(sum(p[3] for p in poz), 2); vat = round(netto * 0.23, 2); brutto = round(netto + vat, 2)
    if blad_brutto:
        brutto = round(brutto + 10, 2)
    c.setFont("F", 10)
    c.drawRightString(190 * mm, y, f"Razem netto: {zl(netto, 'pl')} zł")
    c.drawRightString(190 * mm, y - 6 * mm, f"VAT 23%: {zl(vat, 'pl')} zł")
    c.setFont("FB", 11); c.drawRightString(190 * mm, y - 13 * mm, f"Razem brutto: {zl(brutto, 'pl')} zł")
    c.setFont("F", 9); c.drawString(20 * mm, y - 25 * mm, "Termin płatności: 14 dni, przelew")
    znak_wodny(c); c.save()


def faktura_ekobox(plik, nr, d, poz, zly_nip=False):
    s = DOSTAWCY["ekobox"]; c = canvas.Canvas(str(plik), pagesize=A4)
    nip = s["nip"] if not zly_nip else s["nip"][:-1] + str((int(s["nip"][-1]) + 3) % 10)
    c.setFillColorRGB(0.12, 0.45, 0.3); c.rect(0, 280 * mm, 210 * mm, 17 * mm, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1); c.setFont("FB", 15); c.drawString(20 * mm, 286 * mm, "EKO-BOX")
    c.setFillColorRGB(0, 0, 0)
    c.setFont("FB", 13); c.drawString(20 * mm, 268 * mm, f"FAKTURA NR {nr}")
    c.setFont("F", 10); c.drawString(130 * mm, 268 * mm, f"Wystawiono: {d:%d.%m.%Y}")
    strony_firm(c, 252 * mm, s, nip)
    y = tabela(c, 220 * mm, poz, "en")
    netto = round(sum(p[3] for p in poz), 2); vat = round(netto * 0.23, 2); brutto = round(netto + vat, 2)
    c.setFont("FB", 10)
    c.drawString(20 * mm, y, f"SUMA   netto {zl(netto, 'en')}   VAT {zl(vat, 'en')}   brutto {zl(brutto, 'en')} PLN")
    znak_wodny(c); c.save()


def faktura_barista(plik, nr, d, poz):
    s = DOSTAWCY["barista"]; c = canvas.Canvas(str(plik), pagesize=A4)
    c.setFont("FB", 12); c.drawString(20 * mm, 272 * mm, s["nazwa"])
    c.setFont("F", 10); c.drawRightString(190 * mm, 272 * mm, f"Numer faktury: {nr}")
    c.drawRightString(190 * mm, 266 * mm, f"Data: {d.day} {MIESIACE[d.month - 1]} {d.year}")
    strony_firm(c, 250 * mm, s, s["nip"])
    y = tabela(c, 218 * mm, poz, "pl")
    netto = round(sum(p[3] for p in poz), 2); vat = round(netto * 0.23, 2); brutto = round(netto + vat, 2)
    c.setFont("F", 10)
    c.drawString(110 * mm, y, f"Wartość netto: {zl(netto, 'pl')} PLN")
    c.drawString(110 * mm, y - 6 * mm, f"Kwota VAT (23%): {zl(vat, 'pl')} PLN")
    c.setFont("FB", 12); c.drawString(110 * mm, y - 14 * mm, f"Do zapłaty: {zl(brutto, 'pl')} PLN")
    znak_wodny(c); c.save()


def skan(plik, tekst):
    """PDF zawierający tylko obrazek – jak zeskanowany papier (brak warstwy tekstowej)."""
    img = Image.new("L", (1240, 1754), 245)
    dr = ImageDraw.Draw(img)
    f = ImageFont.truetype(FONT_PATH, 34)
    for i, linia in enumerate(tekst):
        dr.text((110, 140 + i * 60), linia, fill=40, font=f)
    img = img.rotate(0.8, fillcolor=245)
    tmp = FOLDER / "_tmp_skan.png"; img.save(tmp)
    c = canvas.Canvas(str(plik), pagesize=A4)
    c.drawImage(str(tmp), 0, 0, width=210 * mm, height=297 * mm)
    c.save(); tmp.unlink()


start = date(2026, 1, 8)
licznik = {"arabika": 10, "ekobox": 3, "barista": 40}
for i in range(22):
    dostawca = ["arabika", "ekobox", "barista"][i % 3]
    d = start + timedelta(days=i * 11 + random.randint(0, 4))
    licznik[dostawca] += 1
    n = licznik[dostawca]
    poz = pozycje(dostawca)
    if dostawca == "arabika":
        nr = f"FV/{d.year}/{d.month:02d}/{n:03d}"
        faktura_arabika(FOLDER / f"arabika_{n:03d}.pdf", nr, d, poz, blad_brutto=(i == 9))
    elif dostawca == "ekobox":
        nr = f"{n}/{d.month:02d}/{d.year}"
        faktura_ekobox(FOLDER / f"EKOBOX faktura {n}.pdf", nr, d, poz, zly_nip=(i == 13))
    else:
        nr = f"BT-{d.year}-{n:04d}"
        faktura_barista(FOLDER / f"BT-{d.year}-{n:04d}.pdf", nr, d, poz)
        if i == 17:  # ta sama faktura przysłana drugi raz, pod inną nazwą pliku
            faktura_barista(FOLDER / f"skan_mail_{n}.pdf", nr, d, poz)

skan(FOLDER / "skan_paragon_0912.pdf", ["Faktura nr 77/09/2026", "Sprzedawca: Drukarnia Etykiet", "Razem brutto: 455,10 zł"])
print(f"Zapisano {len(list(FOLDER.glob('*.pdf')))} plików PDF w {FOLDER}")
