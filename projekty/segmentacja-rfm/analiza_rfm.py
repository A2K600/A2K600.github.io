"""
Segmentacja klientów metodą RFM
===============================

RFM ocenia każdego klienta w trzech wymiarach:
  R (Recency)   – ile dni minęło od ostatniego zakupu,
  F (Frequency) – ile razy kupił,
  M (Monetary)  – ile łącznie wydał.

Każdy wymiar dostaje ocenę 1–5 (kwintyle), a z kombinacji ocen
powstają segmenty, np. „Najlepsi klienci” albo „Zagrożeni odejściem”.

Uruchomienie:  python analiza_rfm.py
Wynik:  raport.html (raport z wykresami) i segmenty_klientow.xlsx
"""
from __future__ import annotations

import html
import json
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

TU = Path(__file__).parent
DANE = TU / "zamowienia.csv"
DZIEN_ANALIZY = pd.Timestamp("2026-09-30")

# kolejność = priorytet (pierwsza pasująca reguła wygrywa)
SEGMENTY = [
    ("Najlepsi klienci", lambda r, f, m: r >= 4 and f >= 4,
     "Program lojalnościowy, wcześniejszy dostęp do nowości, prośba o opinię."),
    ("Nie możemy ich stracić", lambda r, f, m: r <= 2 and f >= 4 and m >= 4,
     "Osobisty kontakt i mocna oferta powrotu – to byli najcenniejsi klienci."),
    ("Zagrożeni odejściem", lambda r, f, m: r <= 2 and f >= 3,
     "Kampania przypominająca z rabatem na ulubione produkty."),
    ("Lojalni", lambda r, f, m: r >= 3 and f >= 4,
     "Sprzedaż droższych wariantów i zestawów (upselling)."),
    ("Obiecujący", lambda r, f, m: r >= 4 and f >= 2,
     "Zachęta do kolejnego zakupu: darmowa dostawa przy następnym zamówieniu."),
    ("Nowi klienci", lambda r, f, m: r >= 4 and f == 1,
     "Seria powitalnych maili i rabat na drugi zakup."),
    ("Wymagają uwagi", lambda r, f, m: r == 3,
     "Personalizowane rekomendacje, zanim przestaną kupować."),
    ("Uśpieni", lambda r, f, m: r == 2,
     "Tania kampania reaktywacyjna, np. newsletter z bestsellerami."),
    ("Utraceni", lambda r, f, m: True,
     "Nie inwestować dużo – jedna próba odzyskania, potem pomijać w kampaniach."),
]
OPIS = {nazwa: rek for nazwa, _, rek in SEGMENTY}


def ocena(seria: pd.Series, odwrotnie=False) -> pd.Series:
    """Ocena 1–5 wg kwintyli; rank() rozwiązuje problem wielu identycznych wartości."""
    r = seria.rank(method="first", ascending=not odwrotnie)
    return pd.qcut(r, 5, labels=[1, 2, 3, 4, 5]).astype(int)


def policz() -> pd.DataFrame:
    z = pd.read_csv(DANE, parse_dates=["data"])
    k = z.groupby("id_klienta").agg(ostatni=("data", "max"), pierwszy=("data", "min"),
                                    zamowien=("id_zamowienia", "count"), wydal=("wartosc_zl", "sum"))
    k["dni_od_zakupu"] = (DZIEN_ANALIZY - k["ostatni"]).dt.days
    k["R"] = ocena(k["dni_od_zakupu"], odwrotnie=True)   # mniej dni = lepiej
    k["F"] = ocena(k["zamowien"])
    k["M"] = ocena(k["wydal"])
    k["segment"] = [next(n for n, reg, _ in SEGMENTY if reg(r, f, m)) for r, f, m in zip(k.R, k.F, k.M)]
    k["sredni_koszyk"] = k["wydal"] / k["zamowien"]
    return z, k.reset_index()


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------
def excel(k: pd.DataFrame, plik: Path):
    F_ = "Arial"; G = "1F3A5F"
    H = Font(name=F_, bold=True, color="FFFFFF"); N = Font(name=F_, size=10); B = Font(name=F_, size=10, bold=True)
    fill = PatternFill("solid", fgColor=G)
    wb = Workbook()

    def nagl(ws, r, teksty):
        for i, t in enumerate(teksty, 1):
            c = ws.cell(r, i, t); c.font, c.fill = H, fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Klienci
    kl = wb.active; kl.title = "Klienci"
    kol = ["ID klienta", "Pierwszy zakup", "Ostatni zakup", "Dni od zakupu", "Zamówień", "Wydał (zł)",
           "Średni koszyk (zł)", "R", "F", "M", "Segment"]
    nagl(kl, 1, kol)
    k2 = k.sort_values(["segment", "wydal"], ascending=[True, False])
    for i, r in enumerate(k2.itertuples(), start=2):
        vals = [r.id_klienta, r.pierwszy.date(), r.ostatni.date(), r.dni_od_zakupu, r.zamowien,
                round(r.wydal, 2), None, r.R, r.F, r.M, r.segment]
        for j, v in enumerate(vals, 1):
            kl.cell(i, j, v).font = N
        kl.cell(i, 7, f"=F{i}/E{i}")
        for c, fmt in ((2, "yyyy-mm-dd"), (3, "yyyy-mm-dd"), (6, '#,##0.00'), (7, '#,##0.00')):
            kl.cell(i, c).number_format = fmt
    last = len(k2) + 1
    t = Table(displayName="Klienci", ref=f"A1:K{last}"); t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    kl.add_table(t); kl.freeze_panes = "A2"
    for j, w in enumerate([12, 14, 14, 12, 11, 13, 15, 6, 6, 6, 24], 1):
        kl.column_dimensions[get_column_letter(j)].width = w

    # Segmenty (formuły liczone z arkusza Klienci)
    sg = wb.create_sheet("Segmenty", 0); sg.sheet_view.showGridLines = False
    sg["A1"] = "Segmentacja klientów RFM"; sg["A1"].font = Font(name=F_, size=16, bold=True, color=G)
    sg["A2"] = f"Stan na {DZIEN_ANALIZY:%d.%m.%Y}. Wszystkie liczby to formuły z arkusza „Klienci”."
    sg["A2"].font = Font(name=F_, size=9, italic=True, color="666666")
    nagl(sg, 4, ["Segment", "Klientów", "% klientów", "Przychód (zł)", "% przychodu", "Śr. dni od zakupu",
                 "Śr. liczba zamówień", "Co robić"])
    R = lambda c: f"Klienci!${c}$2:${c}${last}"  # noqa: E731
    for i, (nazwa, _, rek) in enumerate(SEGMENTY):
        r = 5 + i
        sg.cell(r, 1, nazwa).font = B
        sg.cell(r, 2, f'=COUNTIFS({R("K")},A{r})')
        sg.cell(r, 3, f'=B{r}/SUM($B$5:$B${4 + len(SEGMENTY)})').number_format = "0.0%"
        sg.cell(r, 4, f'=SUMIFS({R("F")},{R("K")},A{r})').number_format = '#,##0'
        sg.cell(r, 5, f'=D{r}/SUM($D$5:$D${4 + len(SEGMENTY)})').number_format = "0.0%"
        sg.cell(r, 6, f'=IFERROR(AVERAGEIFS({R("D")},{R("K")},A{r}),0)').number_format = "0"
        sg.cell(r, 7, f'=IFERROR(AVERAGEIFS({R("E")},{R("K")},A{r}),0)').number_format = "0.0"
        sg.cell(r, 8, rek)
        for j in range(2, 9):
            sg.cell(r, j).font = N
        sg.cell(r, 8).alignment = Alignment(wrap_text=True, vertical="top")
    for j, w in enumerate([24, 10, 11, 14, 12, 14, 14, 60], 1):
        sg.column_dimensions[get_column_letter(j)].width = w

    # Do kampanii
    dk = wb.create_sheet("Do kampanii")
    dk["A1"] = "Klienci do odzyskania – priorytet wg wydanej kwoty"
    dk["A1"].font = Font(name=F_, size=14, bold=True, color=G)
    nagl(dk, 3, ["ID klienta", "Segment", "Ostatni zakup", "Dni od zakupu", "Zamówień", "Wydał (zł)"])
    kamp = k[k.segment.isin(["Nie możemy ich stracić", "Zagrożeni odejściem"])].sort_values("wydal", ascending=False)
    for i, r in enumerate(kamp.itertuples(), start=4):
        for j, v in enumerate([r.id_klienta, r.segment, r.ostatni.date(), r.dni_od_zakupu, r.zamowien, round(r.wydal, 2)], 1):
            dk.cell(i, j, v).font = N
        dk.cell(i, 3).number_format = "yyyy-mm-dd"; dk.cell(i, 6).number_format = '#,##0.00'
    for j, w in enumerate([12, 24, 14, 13, 11, 13], 1):
        dk.column_dimensions[get_column_letter(j)].width = w
    wb.save(plik)


# ---------------------------------------------------------------------------
# Raport HTML
# ---------------------------------------------------------------------------
def fmt(v, d=0):
    return f"{v:,.{d}f}".replace(",", " ").replace(".", ",")


def raport(z: pd.DataFrame, k: pd.DataFrame, plik: Path):
    s = k.groupby("segment").agg(klientow=("id_klienta", "count"), przychod=("wydal", "sum"),
                                 dni=("dni_od_zakupu", "mean"), zam=("zamowien", "mean"))
    s = s.reindex([n for n, *_ in SEGMENTY]).fillna(0)
    s["pk"] = s.klientow / s.klientow.sum(); s["pp"] = s.przychod / s.przychod.sum()
    total = k.wydal.sum()

    naj = s.loc["Najlepsi klienci"]; ryz = s.loc[["Nie możemy ich stracić", "Zagrożeni odejściem"]]
    nowi = s.loc["Nowi klienci"]
    jedno = (k.zamowien == 1).mean()
    wnioski = [
        f"<b>{fmt(naj.pk * 100, 0)}% klientów przynosi {fmt(naj.pp * 100, 0)}% przychodu.</b> To segment „Najlepsi klienci”: "
        f"kupują średnio {fmt(naj.zam, 1)} razy i byli w sklepie średnio {fmt(naj.dni)} dni temu.",
        f"<b>{fmt(ryz.klientow.sum())} klientów wartych łącznie {fmt(ryz.przychod.sum())} zł przestaje kupować</b> "
        f"(segmenty „Nie możemy ich stracić” i „Zagrożeni odejściem”). Ostatni zakup zrobili średnio "
        f"{fmt((ryz.dni * ryz.klientow).sum() / ryz.klientow.sum())} dni temu. To pierwsza grupa do kampanii odzyskującej.",
        f"<b>{fmt(jedno * 100, 0)}% wszystkich klientów kupiło tylko raz.</b> W tym {fmt(nowi.klientow)} to nowi klienci "
        f"z ostatnich miesięcy. Seria powitalna i rabat na drugi zakup mogą zamienić część z nich w stałych klientów.",
    ]

    # --- wykres 1: % klientów vs % przychodu (słupki poziome, 2 serie) ---
    W, rowH, L, Rm = 760, 46, 190, 70
    H_ = 30 + len(s) * rowH
    mx = max(s.pk.max(), s.pp.max())
    sx = lambda v: L + v / mx * (W - L - Rm)  # noqa: E731
    bars = []
    for i, (nazwa, r) in enumerate(s.iterrows()):
        y = 30 + i * rowH
        bars.append(f'<text x="{L - 12}" y="{y + 22}" text-anchor="end" class="cat">{html.escape(nazwa)}</text>')
        for j, (v, cls, lab) in enumerate([(r.pk, "s1", "klientów"), (r.pp, "s2", "przychodu")]):
            yy = y + 6 + j * 16; w = max(sx(v) - L, 1)
            bars.append(f'<g><title>{html.escape(nazwa)}: {fmt(v * 100, 1)}% {lab}</title>'
                        f'<path class="{cls}" d="M{L},{yy}h{w - 3}a3,3 0 0 1 3,3v6a3,3 0 0 1 -3,3h-{w - 3}z"/>'
                        f'<text x="{L + w + 6}" y="{yy + 10}" class="val">{fmt(v * 100, 0)}%</text></g>')
    for p in (0, 0.1, 0.2, 0.3, 0.4, 0.5):
        if p <= mx:
            bars.insert(0, f'<line x1="{sx(p)}" x2="{sx(p)}" y1="22" y2="{H_}" class="grid"/>'
                           f'<text x="{sx(p)}" y="14" text-anchor="middle" class="tick">{int(p * 100)}%</text>')
    svg1 = f'<svg viewBox="0 0 {W} {H_}" role="img" aria-label="Udział segmentów w liczbie klientów i przychodzie">{"".join(bars)}</svg>'

    # --- wykres 2: mapa R × F (liczba klientów) ---
    grid = k.groupby(["R", "F"]).size().unstack(fill_value=0).reindex(index=[5, 4, 3, 2, 1], columns=[1, 2, 3, 4, 5], fill_value=0)
    gmax = grid.values.max()
    ramp = ["#cde2fb", "#9ec5f4", "#6da7ec", "#3987e5", "#256abf", "#184f95", "#0d366b"]
    cell, ox, oy = 64, 70, 16
    cells = []
    for i, rv in enumerate(grid.index):
        cells.append(f'<text x="{ox - 10}" y="{oy + i * cell + cell / 2 + 4}" text-anchor="end" class="tick">R{rv}</text>')
        for j, fv in enumerate(grid.columns):
            n = int(grid.loc[rv, fv]); idx = min(len(ramp) - 1, int(n / gmax * (len(ramp) - 1) + 0.5)) if gmax else 0
            col = ramp[idx]; txt = "#fff" if idx >= 3 else "#1a1a19"
            cells.append(f'<g><title>R={rv}, F={fv}: {n} klientów</title>'
                         f'<rect x="{ox + j * cell + 1}" y="{oy + i * cell + 1}" width="{cell - 2}" height="{cell - 2}" rx="4" fill="{col}"/>'
                         f'<text x="{ox + j * cell + cell / 2}" y="{oy + i * cell + cell / 2 + 4}" text-anchor="middle" fill="{txt}" class="cnt">{n}</text></g>')
    for j, fv in enumerate(grid.columns):
        cells.append(f'<text x="{ox + j * cell + cell / 2}" y="{oy + 5 * cell + 18}" text-anchor="middle" class="tick">F{fv}</text>')
    svg2 = (f'<svg viewBox="0 0 {ox + 5 * cell + 10} {oy + 5 * cell + 28}" role="img" '
            f'aria-label="Mapa klientów wg ocen R i F">{"".join(cells)}</svg>')

    # --- wykres 3: przychód miesięczny ---
    mies = z.set_index("data").resample("MS")["wartosc_zl"].sum()
    W3, H3, l3, b3, t3 = 760, 220, 56, 26, 10
    m3 = mies.max(); step = 25000 if m3 > 80000 else (10000 if m3 > 30000 else 5000)
    top = (int(m3 / step) + 1) * step
    xs = lambda i: l3 + i * (W3 - l3 - 20) / (len(mies) - 1)  # noqa: E731
    ys = lambda v: t3 + (H3 - t3 - b3) * (1 - v / top)  # noqa: E731
    parts = []
    for v in range(0, top + 1, step):
        parts.append(f'<line x1="{l3}" x2="{W3 - 20}" y1="{ys(v)}" y2="{ys(v)}" class="{"axis" if v == 0 else "grid"}"/>'
                     f'<text x="{l3 - 8}" y="{ys(v) + 4}" text-anchor="end" class="tick">{v // 1000} tys.</text>')
    for i, d in enumerate(mies.index):
        if d.month in (1, 4, 7, 10):
            parts.append(f'<text x="{xs(i)}" y="{H3 - 6}" text-anchor="middle" class="tick">{d:%m.%Y}</text>')
    path = "".join(f'{"M" if i == 0 else "L"}{xs(i):.1f},{ys(v):.1f}' for i, v in enumerate(mies.values))
    parts.append(f'<path d="{path}" class="line"/>')
    for i, (d, v) in enumerate(mies.items()):
        parts.append(f'<g><title>{d:%m.%Y}: {fmt(v)} zł</title><circle cx="{xs(i)}" cy="{ys(v)}" r="9" fill="transparent"/>'
                     f'<circle cx="{xs(i)}" cy="{ys(v)}" r="3" class="dot"/></g>')
    svg3 = f'<svg viewBox="0 0 {W3} {H3}" role="img" aria-label="Przychód miesięczny sklepu">{"".join(parts)}</svg>'

    wiersze = "".join(
        f"<tr><td><b>{html.escape(n)}</b></td><td class='n'>{fmt(r.klientow)}</td><td class='n'>{fmt(r.pk * 100, 1)}%</td>"
        f"<td class='n'>{fmt(r.przychod)} zł</td><td class='n'>{fmt(r.pp * 100, 1)}%</td><td class='n'>{fmt(r.dni)}</td>"
        f"<td>{html.escape(OPIS[n])}</td></tr>" for n, r in s.iterrows())

    pocz, kon = mies.iloc[:6].mean(), mies.iloc[-6:].mean()
    spadek = 1 - kon / pocz
    trend_opis = (f"Średni przychód miesięczny spadł z {fmt(pocz)} zł (pierwsze 6 miesięcy) do {fmt(kon)} zł "
                  f"(ostatnie 6 miesięcy), czyli o {fmt(spadek * 100, 0)}%. Spadek pokrywa się z odejściem klientów "
                  f"z segmentu „Nie możemy ich stracić”.") if spadek > 0.05 else "Przychód miesięczny jest stabilny."
    dane = dict(klienci=len(k), zamowien=len(z), przychod=total, od=z.data.min(), do=z.data.max())
    tpl = (TU / "szablon_raportu.html").read_text(encoding="utf-8")
    out = (tpl.replace("{{KLIENCI}}", fmt(dane["klienci"])).replace("{{ZAMOWIEN}}", fmt(dane["zamowien"]))
              .replace("{{PRZYCHOD}}", fmt(dane["przychod"])).replace("{{OKRES}}", f"{dane['od']:%d.%m.%Y} – {dane['do']:%d.%m.%Y}")
              .replace("{{WNIOSKI}}", "".join(f"<li>{w}</li>" for w in wnioski))
              .replace("{{SVG_SEGMENTY}}", svg1).replace("{{SVG_MAPA}}", svg2).replace("{{SVG_TREND}}", svg3)
              .replace("{{TABELA}}", wiersze).replace("{{TREND_OPIS}}", trend_opis).replace("{{DZIEN}}", f"{DZIEN_ANALIZY:%d.%m.%Y}"))
    plik.write_text(out, encoding="utf-8")
    return s


def main():
    z, k = policz()
    excel(k, TU / "segmenty_klientow.xlsx")
    s = raport(z, k, TU / "index.html")
    print(f"{len(k)} klientów, {len(z)} zamówień, przychód {fmt(k.wydal.sum())} zł")
    print(s[["klientow", "pk", "pp"]].round(3).to_string())
    print("Gotowe → index.html (raport) i segmenty_klientow.xlsx")


if __name__ == "__main__":
    main()
