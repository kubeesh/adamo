# -*- coding: utf-8 -*-
"""Test na 100 częściach: 5 folderów wydania po 20 części (BOM + DXF + PDF), z czego
52 poprawne i 48 z celowo wprowadzonym błędem DXF (po 4 z każdego z 12 typów).
Rysunki PDF mają prawdziwy widok wektorowy części w skali (jak z Inventora: grube linie
widoczne, cienkie wymiarowe, widok z boku, ramka), więc działa też nakładka DXF 1:1.
Liczy, ile błędów program wykrył, ile przeoczył i ile zgłosił fałszywych alarmów.

Części poprawne celowo zawierają rzeczy, które NIE są błędem, a mogłyby dać fałszywy alarm:
zaokrąglone narożniki (polilinia z łukami), krawędź podzieloną na dwa odcinki (jak w eksporcie
Inventora), linie gięcia, linie ukryte i osie otworów na osobnych warstwach, otwory pod gwint,
fazy zwymiarowane liczbą zamiast notki, szum numeryczny współrzędnych, części z ukosowaniem
(krawędź ukosu narysowana na rysunku, w DXF jej nie ma), DXF obrócony o 90° względem rysunku.

Uruchom:  python testy/test_100_czesci.py [folder_na_pliki]
"""
import os
import random
import re
import sys
import tempfile

import math

import ezdxf
import ezdxf.math
import fitz
import openpyxl

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import sprawdz_hybrydowo_GEO as S  # noqa: E402

FONT = next((f for f in ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", r"C:\Windows\Fonts\arial.ttf")
             if os.path.isfile(f)), None)
WID, UKR, OSIE = "Widoczne (ISO)", "Ukryte (ISO)", "Osie (ISO)"
WADY = ["zla_faza", "brak_fazy", "nieprzycieta_faza", "przerwa", "podwojny_kontur",
        "otwor_bez_wymiaru", "pogłębienie", "zly_gabaryt", "wystajaca_linia", "zly_otwor",
        "przesuniety_otwor", "zly_ksztalt"]
# czego szukać w opisie wiersza raportu, żeby uznać wadę za wykrytą (nakładka 1:1 łapie każdą
# różnicę konturu, więc dla wad geometrii też się liczy)
NAK = r"|nakładka — DXF odbiega"
WYKRYCIE = {
    "zla_faza": r"fazy —", "brak_fazy": r"fazy —", "nieprzycieta_faza": r"kontur — .*rozgałęzień",
    "przerwa": r"przerw", "podwojny_kontur": r"podwójne linie", "otwor_bez_wymiaru": r"dorysowany otwór",
    "pogłębienie": r"współśrodkowe", "zly_gabaryt": r"geometria DXF", "wystajaca_linia": r"wystaje",
    "zly_otwor": r"otwory —", "przesuniety_otwor": r"nakładka — DXF odbiega", "zly_ksztalt": r"nakładka — DXF odbiega",
}
for _t in ("zla_faza", "brak_fazy", "nieprzycieta_faza", "zly_gabaryt", "zly_otwor"):
    WYKRYCIE[_t] += NAK


def losuj_czesc(R, nr, wada):
    W = R.choice([150, 200, 240, 300, 420, 500, 640, 800, 1000, 1226, 1500, 2000])
    H = R.choice([100, 150, 180, 250, 300, 400, 600, 900])
    c = dict(part=f"PG{nr:05d}", t=R.choice([6, 8, 10, 12, 15, 20, 25, 30]), W=W, H=H, qty=R.choice([1, 2, 4]),
             wada=wada, fazy={}, otwory=[], flags="", styl=R.choice(["linie", "poly"]))
    potrzebne_fazy = wada in ("zla_faza", "brak_fazy", "nieprzycieta_faza")
    if potrzebne_fazy or R.random() < 0.5:
        rozm = R.choice([5, 10, 15, 20, 25])
        for k in R.sample(range(4), R.choice([1, 2, 2, 4])):
            c["fazy"][k] = rozm
    c["fazy_notka"] = potrzebne_fazy or R.random() < 0.7   # inaczej faza zwymiarowana samą liczbą
    potrzebne_otwory = wada in ("pogłębienie", "zly_otwor", "przesuniety_otwor")
    if potrzebne_otwory or R.random() < 0.6:
        d = R.choice([11, 13, 17.5, 22, 26])
        for _ in range(R.choice([1, 2, 4, 6])):
            c["otwory"].append((R.uniform(45, W - 45), R.uniform(45, H - 45), d))
        if R.random() < 0.3:
            c["otwory"].append((W / 2, H / 2, 14.0))    # pod gwint M16
            c["gwint"] = 16
    c["zaokr"] = c["styl"] == "poly" and R.random() < 0.5    # narożniki bez fazy zaokrąglone R10
    c["podzial"] = c["styl"] == "linie" and R.random() < 0.5   # dolna krawędź z dwóch odcinków
    if wada is None and R.random() < 0.2:
        c["flags"] = "p"                                      # gięta, z linią gięcia
    elif wada is None and R.random() < 0.2:
        c["flags"] = "u"                                      # ukosowana: krawędź ukosu na rysunku
    c["obrot_dxf"] = R.random() < 0.2                         # DXF obrócony o 90° względem rysunku
    if wada == "przerwa":
        c["styl"] = "linie"
    return c


def kontur(c, W=None):
    """Lista elementów obrysu: ("L", p1, p2) albo ("A", p1, p2, bulge) — po kolei wokół części."""
    W = W or c["W"]
    H = c["H"]
    rogi = [(0, 0), (W, 0), (W, H), (0, H)]
    pkt = []   # (punkt, bulge do następnego)
    for k, (x, y) in enumerate(rogi):
        poprz, nast = rogi[k - 1], rogi[(k + 1) % 4]
        def ku(p, d):
            dx, dy = p[0] - x, p[1] - y
            n = (dx * dx + dy * dy) ** 0.5
            return (x + dx / n * d, y + dy / n * d)
        if k in c["fazy"]:
            pkt += [(ku(poprz, c["fazy"][k]), 0), (ku(nast, c["fazy"][k]), 0)]
        elif c.get("zaokr"):
            pkt += [(ku(poprz, 10), 0.41421356), (ku(nast, 10), 0)]
        else:
            pkt.append(((x, y), 0))
    el = []
    for i, (p, b) in enumerate(pkt):
        q = pkt[(i + 1) % len(pkt)][0]
        el.append(("A", p, q, b) if b else ("L", p, q))
    return el


def zapisz_dxf(c, sciezka, R):
    doc = ezdxf.new("R2010")
    msp = doc.modelspace()
    szum = lambda p: (p[0] + R.uniform(-1e-7, 1e-7), p[1] + R.uniform(-1e-7, 1e-7))  # noqa: E731
    w = c["wada"]
    W = c["W"] + 5 if w == "zly_gabaryt" else c["W"]
    geo = dict(c, fazy=dict(c["fazy"]))
    if w == "zla_faza":
        k = next(iter(geo["fazy"]))
        geo["fazy"][k] += 3
    if w in ("brak_fazy", "nieprzycieta_faza"):
        k = next(iter(geo["fazy"]))
        del geo["fazy"][k]
    el = kontur(geo, W)
    if c["styl"] == "poly":
        pts = [(*szum(e[1]), e[3] if e[0] == "A" else 0) for e in el]
        msp.add_lwpolyline(pts, format="xyb", close=True, dxfattribs={"layer": "IV_OUTER_PROFILE"})
    else:
        linie = [(szum(e[1]), szum(e[2])) for e in el]
        if w == "przerwa":
            (a, b) = linie[1]
            dl = ((b[0] - a[0]) ** 2 + (b[1] - a[1]) ** 2) ** 0.5
            g = R.uniform(0.3, 1.5) / dl
            linie[1] = (a, (b[0] - (b[0] - a[0]) * g, b[1] - (b[1] - a[1]) * g))
        if c["podzial"]:
            a, b = linie[0]
            m = ((a[0] + b[0]) / 2, (a[1] + b[1]) / 2)
            linie[0:1] = [(a, m), (m, b)]
        for a, b in linie:
            msp.add_line(a, b, dxfattribs={"layer": WID})
        if w == "podwojny_kontur":
            for a, b in linie:
                msp.add_line(b, a, dxfattribs={"layer": WID})
    if w == "podwojny_kontur" and c["styl"] == "poly":
        msp.add_lwpolyline(pts, format="xyb", close=True, dxfattribs={"layer": "IV_OUTER_PROFILE"})
    if w == "nieprzycieta_faza":
        k = next(iter(c["fazy"]))
        f = [e for e in kontur(c) if e[0] == "L"]
        rog = [(0, 0), (c["W"], 0), (c["W"], c["H"]), (0, c["H"])][k]
        fz = min(f, key=lambda e: abs(e[1][0] - rog[0]) + abs(e[1][1] - rog[1]) + abs(e[2][0] - rog[0])
                 + abs(e[2][1] - rog[1]))   # odcinek fazy przy tym narożniku w geometrii z rysunku
        msp.add_line(fz[1], fz[2], dxfattribs={"layer": "0"})
    warstwa_otw = "IV_INTERIOR_PROFILES" if c["styl"] == "poly" else WID
    if w == "zly_ksztalt":     # prawy górny narożnik cofnięty o 4 mm — gabaryt się nie zmienia
        for e in msp.query("LINE LWPOLYLINE"):
            if e.dxftype() == "LINE":
                for a in ("start", "end"):
                    v = e.dxf.get(a)
                    if abs(v.y - c["H"]) < 1e-3 and v.x > W - 30:
                        e.dxf.set(a, (v.x - 4, v.y, 0))
            else:
                pts = [(x - 4 if abs(y - c["H"]) < 1e-3 and x > W - 30 else x, y, sw, ew, b)
                       for x, y, sw, ew, b in e.get_points()]
                e.set_points(pts)
    for i, (x, y, d) in enumerate(c["otwory"]):
        if w == "zly_otwor" and i == 0:
            d += 2
        if w == "przesuniety_otwor" and i == 0:
            x += 5
        msp.add_circle(szum((x, y)), d / 2, dxfattribs={"layer": warstwa_otw})
        msp.add_line((x - d, y), (x + d, y), dxfattribs={"layer": OSIE})
        msp.add_line((x, y - d), (x, y + d), dxfattribs={"layer": OSIE})
        if w == "pogłębienie" and i == 0:
            msp.add_circle((x, y), d / 2 + 2, dxfattribs={"layer": warstwa_otw})
    if w == "otwor_bez_wymiaru":
        msp.add_circle((c["W"] / 2, c["H"] / 3), 7.5, dxfattribs={"layer": "0"})
    if w == "wystajaca_linia":
        msp.add_line((c["W"] + 40, 0), (c["W"] + 40, c["H"] / 2), dxfattribs={"layer": "0"})
    if "p" in c["flags"]:
        msp.add_line((W / 2, 0), (W / 2, c["H"]), dxfattribs={"layer": "IV_BEND"})
    if c["styl"] == "linie" and R.random() < 0.4:
        msp.add_line((0, c["H"] / 2), (W, c["H"] / 2), dxfattribs={"layer": UKR})
    if c["obrot_dxf"]:
        for e in msp:
            e.transform(ezdxf.math.Matrix44.z_rotate(math.pi / 2))
    doc.saveas(sciezka)


def mm(x):
    return f"{x:g}".replace(".", ",")


def zapisz_pdf(c, sciezka, R):
    """Rysunek A4 poziomo: widok części w skali 1:k grubą linią (0,54 pt, jak Inventor), wymiary
    cienką (0,36), widok z boku (grubość), ramka, tabliczka i notki jako tekst."""
    doc = fitz.open()
    strona = doc.new_page(width=842, height=595)
    k = next(k for k in (1, 2, 2.5, 5, 7.5, 10, 15, 20, 25) if max(c["W"] / 520, c["H"] / 300) * 72 / 25.4 <= k)
    s = 72 / 25.4 / k
    x0, y0 = 120, 120 + c["H"] * s                       # lewy dolny róg widoku na stronie
    P = lambda p: fitz.Point(x0 + p[0] * s, y0 - p[1] * s)  # noqa: E731
    gr, cien = 0.54, 0.36
    for e in kontur(c):
        if e[0] == "L":
            strona.draw_line(P(e[1]), P(e[2]), width=gr)
        else:   # łuk 90° z bulge -> krzywa: przybliżenie odcinkami (Inventor daje Beziera, oba są wektorowe)
            (x1, y1), (x2, y2) = e[1], e[2]
            cx, cy = (x1 + x2) / 2 + (y2 - y1) / 2, (y1 + y2) / 2 - (x2 - x1) / 2
            r = math.dist((cx, cy), (x1, y1))
            a1, a2 = math.atan2(y1 - cy, x1 - cx), math.atan2(y2 - cy, x2 - cx)
            if a2 < a1:
                a2 += 2 * math.pi
            pts = [P((cx + r * math.cos(t), cy + r * math.sin(t))) for t in [a1 + (a2 - a1) * i / 12 for i in range(13)]]
            strona.draw_polyline(pts, width=gr)
    for x, y, d in c["otwory"]:
        strona.draw_circle(P((x, y)), d / 2 * s, width=gr)
        strona.draw_line(P((x - d, y)), P((x + d, y)), width=cien)
    if "u" in c["flags"]:      # krawędź ukosu 7 mm od dolnej krawędzi
        f0, f1 = c["fazy"].get(0, 0), c["fazy"].get(1, 0)
        strona.draw_line(P((max(f0, 7), 7)), P((c["W"] - max(f1, 7), 7)), width=gr)
    if "p" in c["flags"]:
        strona.draw_line(P((c["W"] / 2, 0)), P((c["W"] / 2, c["H"])), width=cien, dashes="[6 2 1 2] 0")
    # wymiary (cienkie, bez dotykania konturu)
    for i in range(30):
        strona.draw_line(P((0, -8 - i % 3)), P((c["W"], -8 - i % 3)), width=cien)
    strona.draw_line(P((-8, 0)), P((-8, c["H"])), width=cien)
    # widok z boku: prostokąt W x t, i ramka arkusza
    strona.draw_rect(fitz.Rect(x0, y0 + 40, x0 + c["W"] * s, y0 + 40 + max(c["t"] * s, 1)), width=gr)
    strona.draw_rect(fitz.Rect(20, 20, 822, 575), width=0.72)
    tekst = [f"VIEW1 ( 1 : {mm(k)} )", f"PL {c['t']} x {c['W']} x {c['H']}", "S355", f"Ilość: {c['qty']}",
             mm(c["W"]), mm(c["H"])]
    rozm = {}
    for f in c["fazy"].values():
        rozm[f] = rozm.get(f, 0) + 1
    for f, n in rozm.items():
        tekst.append(((f"{n}x " if n > 1 else "") + f"{f} x 45°") if c["fazy_notka"] else mm(f))
    if "u" in c["flags"]:
        tekst.append("7,00 X 45° Chamfer")       # faza krawędzi (ukos) — nie ma jej w konturze
    otw = {}
    for _, _, d in c["otwory"]:
        if d != 14.0:
            otw[d] = otw.get(d, 0) + 1
    for d, n in otw.items():
        tekst.append((f"{n}x " if n > 1 else "") + R.choice(["n", "Ø"]) + mm(d))
    if c.get("gwint"):
        tekst.append(f"M{c['gwint']}")
    for x, y, _ in c["otwory"][:3]:
        tekst += [mm(round(x)), mm(round(y))]
    for i, l in enumerate(tekst):
        strona.insert_text((560, 60 + 13 * i), l, fontname="dv", fontfile=FONT, fontsize=8)
    doc.save(sciezka)


def zbuduj(baza, R):
    wady = [w for w in WADY for _ in range(4)] + [None] * (100 - 4 * len(WADY))
    R.shuffle(wady)
    czesci = []
    for f in range(5):
        nazwa = f"GE91{f:02d}"
        folder = os.path.join(baza, nazwa)
        os.makedirs(folder, exist_ok=True)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["POS", "PART NUMBER", "TITLE", "", "BOM STRUCTURE", "QTY", "", "DESCRIPTION", "", "",
                   "MATERIAL", "M", "O", "P", "U", "C"])
        for i in range(20):
            nr = f * 20 + i + 1
            c = losuj_czesc(R, nr, wady[nr - 1])
            c["folder"] = nazwa
            flagi = [1 if fl in c["flags"] else None for fl in "mopuc"]
            ws.append([i + 1, c["part"], "PLATE", "", "Normal", c["qty"], "",
                       f"PL {c['t']} x {c['W']} x {c['H']}", "", "", "S355", *flagi])
            ogon = f"{c['t']}mm__S355__{c['qty']}" + "".join(f"__{fl}" for fl in c["flags"])
            zapisz_dxf(c, os.path.join(folder, f"{nazwa}__{c['part']}__{ogon}.dxf"), R)
            zapisz_pdf(c, os.path.join(folder, f"{c['part']}__{ogon}.pdf"), R)
            czesci.append(c)
        wb.save(os.path.join(folder, f"{nazwa}.xlsx"))
    return czesci


def ocen(czesci, baza):
    S.NAKLADKI_PDF.clear()
    S.WSZYSTKIE_WIERSZE.clear()
    S.PODSUMOWANIE_FOLDEROW.clear()
    for f in sorted({c["folder"] for c in czesci}):
        S.sprawdz(os.path.join(baza, f))
    wiersze = {}
    for w in S.WSZYSTKIE_WIERSZE:
        wiersze.setdefault(w["part"], []).append(w)
    wynik = dict(bl=[], uw=[], przeocz=[], falsz_bl=[], falsz_uw=[], zly_typ=[])
    for c in czesci:
        ws = wiersze.get(c["part"], [])
        bl = " | ".join(w["opis"] for w in ws if w["poziom"] == "BŁĄD")
        uw = " | ".join(w["opis"] for w in ws if w["poziom"] == "UWAGA")
        if c["wada"] is None:
            if bl:
                wynik["falsz_bl"].append((c, bl))
            if uw:
                wynik["falsz_uw"].append((c, uw))
            continue
        wzor = WYKRYCIE[c["wada"]]
        if re.search(wzor, bl):
            wynik["bl"].append(c)
        elif re.search(wzor, uw):
            wynik["uw"].append((c, uw))
        elif bl:
            wynik["zly_typ"].append((c, bl))
        else:
            wynik["przeocz"].append((c, uw))
    return wynik


def main():
    baza = sys.argv[1] if len(sys.argv) > 1 else tempfile.mkdtemp(prefix="test100_")
    R = random.Random(2026)
    czesci = zbuduj(baza, R)
    stdout = sys.stdout
    sys.stdout = open(os.devnull, "w", encoding="utf-8")   # konsola sprawdz() — wyciszona
    try:
        wynik = ocen(czesci, baza)
    finally:
        sys.stdout.close()
        sys.stdout = stdout
    raport = os.path.join(baza, "RAPORT_TEST_100.xlsx")
    S.zapisz_raport_xlsx(raport)

    print(f"Pliki testowe: {baza}\nRaport programu: {raport}\n")
    n_wad = 4 * len(WADY)
    print(f"Części z wadą: {n_wad}, bez wady: {100 - n_wad}")
    print(f"  wykryte jako BŁĄD:            {len(wynik['bl'])}")
    print(f"  wykryte tylko jako UWAGA:     {len(wynik['uw'])}")
    print(f"  BŁĄD, ale innego rodzaju:     {len(wynik['zly_typ'])}")
    print(f"  przeoczone:                   {len(wynik['przeocz'])}")
    print(f"  fałszywe BŁĘDY (dobre części): {len(wynik['falsz_bl'])}")
    print(f"  UWAGI na dobrych częściach:   {len(wynik['falsz_uw'])}")
    print("\nWg typu wady (wykryte jako BŁĄD / 4):")
    for t in WADY:
        print(f"  {t:20s} {sum(1 for c in wynik['bl'] if c['wada'] == t)}/4")
    if S.NAKLADKI_PDF:
        print(f"\nPodgląd nakładek: {S.zapisz_nakladki_pdf(os.path.join(baza, 'NAKLADKI_TEST_100.pdf'))}")
    for klucz, tytul in (("uw", "TYLKO UWAGA"), ("zly_typ", "INNY RODZAJ"), ("przeocz", "PRZEOCZONE"),
                         ("falsz_bl", "FAŁSZYWE BŁĘDY"), ("falsz_uw", "UWAGI NA DOBRYCH")):
        for c, opis in wynik[klucz]:
            print(f"  [{tytul}] {c['part']} {c['wada'] or ''} styl={c['styl']} fazy={c['fazy']} -> {opis[:300]}")
    return wynik


def test_100_czesci():
    w = main()
    assert len(w["bl"]) == 4 * len(WADY) and not w["falsz_bl"], w


if __name__ == "__main__":
    w = main()
    sys.exit(0 if len(w["bl"]) == 4 * len(WADY) and not w["falsz_bl"] else 1)
