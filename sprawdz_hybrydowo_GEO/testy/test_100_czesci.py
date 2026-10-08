# -*- coding: utf-8 -*-
"""Test na 100 częściach: 5 folderów wydania po 20 części (BOM + DXF + PDF), z czego
44 poprawne i 56 z celowo wprowadzoną wadą DXF (po 4 z każdego z 14 typów).

Rysunki PDF mają prawdziwy widok wektorowy części w skali, jak z Inventora: grube linie widoczne,
cienkie wymiarowe, widok z boku, ramka. Każda część losuje też ODMIANY rysunku/DXF, na których
program ma działać tak samo (żadna z nich nie jest błędem):
  - widok obrócony na arkuszu o dowolny kąt, strona PDF obrócona o 90°,
  - brak opisu skali widoku "( 1 : k )" (skala do odgadnięcia),
  - linie ukryte (kreskowane) tej samej grubości w widoku, grube osie otworów,
  - długa linia przekroju doklejona do widoku (skupisko linii dużo większe niż część),
  - inne grubości linii (widoczne 0,7 pt, wymiary 0,25 pt),
  - DXF: geometria w bloku (INSERT), DXF obrócony o 90°, krawędź podzielona na dwa odcinki,
    zaokrąglone narożniki (polilinia z łukami), linie gięcia/ukryte/osie na osobnych warstwach,
    szum numeryczny współrzędnych,
  - części z ukosowaniem (krawędź ukosu na rysunku, w DXF jej nie ma) i gięte.

Zaliczenie: każda wada wykryta właściwym komunikatem (BŁĄD, a odbicie lustrzane jako UWAGA ze statusem
NIEZWERYFIKOWANY), żadna część z wadą nie dostaje statusu "ZGODNY 1:1", a każda poprawna część ma
status "ZGODNY 1:1" i ani jednego błędu ani uwagi.

Uruchom:  python testy/test_100_czesci.py [folder_na_pliki]
"""
import math
import os
import random
import re
import sys
import tempfile

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
        "przesuniety_otwor", "zly_ksztalt", "brak_otworu_w_dxf", "dxf_lustro"]
ODMIANY = ["obrot_widoku", "strona_obrocona", "bez_etykiety", "ukryte_kreskowane", "osie_grube",
           "przekroj_dlugi", "inne_grubosci", "insert_dxf"]
# czego szukać w komunikatach części, żeby uznać wadę za wykrytą, i na jakim poziomie
NAK = r"|nakładka — DXF odbiega"
WYKRYCIE = {
    "zla_faza": r"fazy —" + NAK, "brak_fazy": r"fazy —" + NAK, "nieprzycieta_faza": r"kontur — .*rozgałęzień",
    "przerwa": r"przerw", "podwojny_kontur": r"podwójne linie", "otwor_bez_wymiaru": r"dorysowany otwór" + NAK,
    "pogłębienie": r"współśrodkowe", "zly_gabaryt": r"geometria DXF" + NAK, "wystajaca_linia": r"wystaje",
    "zly_otwor": r"otwory —" + NAK, "przesuniety_otwor": r"nakładka — DXF odbiega",
    "zly_ksztalt": r"nakładka — DXF odbiega", "brak_otworu_w_dxf": r"nakładka — na rysunku jest .* zamknięty",
    "dxf_lustro": r"ODBICIU LUSTRZANYM",
}
POZIOM = {"dxf_lustro": "UWAGA"}   # lustro to nie błąd sam w sobie — ale DXF nie może być "ZGODNY 1:1"


def losuj_czesc(R, nr, wada):
    W = R.choice([150, 200, 240, 300, 420, 500, 640, 800, 1000, 1226, 1500, 2000])
    H = R.choice([100, 150, 180, 250, 300, 400, 600, 900])
    if wada == "dxf_lustro":
        while H == W:
            H = R.choice([100, 180, 250, 400])
    c = dict(part=f"PG{nr:05d}", t=R.choice([6, 8, 10, 12, 15, 20, 25, 30]), W=W, H=H, qty=R.choice([1, 2, 4]),
             wada=wada, fazy={}, otwory=[], flags="", styl=R.choice(["linie", "poly"]))
    potrzebne_fazy = wada in ("zla_faza", "brak_fazy", "nieprzycieta_faza")
    if wada == "dxf_lustro":
        c["fazy"][R.randrange(4)] = R.choice([10, 15, 20])         # jedna faza: część niesymetryczna
    elif potrzebne_fazy or R.random() < 0.5:
        rozm = R.choice([5, 10, 15, 20, 25])
        for k in R.sample(range(4), R.choice([1, 2, 2, 4])):
            c["fazy"][k] = rozm
    c["fazy_notka"] = potrzebne_fazy or R.random() < 0.7   # inaczej faza zwymiarowana samą liczbą
    potrzebne_otwory = wada in ("pogłębienie", "zly_otwor", "przesuniety_otwor", "brak_otworu_w_dxf")
    if potrzebne_otwory or R.random() < 0.6:
        d = R.choice([11, 13, 17.5, 22, 26])
        for _ in range(R.choice([1, 2, 4, 6])):
            c["otwory"].append((R.uniform(45, W - 45), R.uniform(45, H - 45), d))
        if R.random() < 0.3:
            c["otwory"].append((W / 2, H / 2, 14.0))    # pod gwint M16
            c["gwint"] = 16
    c["odmiany"] = {o for o in ODMIANY if R.random() < 0.15}
    if wada == "przerwa" or "insert_dxf" in c["odmiany"]:
        c["styl"] = "linie"
    c["zaokr"] = c["styl"] == "poly" and R.random() < 0.5 and wada != "zly_ksztalt"   # łuki tylko w polilinii
    c["podzial"] = c["styl"] == "linie" and R.random() < 0.5   # krawędź z dwóch odcinków
    if wada is None and R.random() < 0.2:
        c["flags"] = "p"                                      # gięta, z linią gięcia
    elif wada is None and R.random() < 0.2:
        c["flags"] = "u"                                      # ukosowana: krawędź ukosu na rysunku
    c["obrot_dxf"] = R.random() < 0.2                         # DXF obrócony o 90° względem rysunku
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
    doc = ezdxf.new("R2010", units=ezdxf.units.MM)   # jak Inventor: $INSUNITS=4
    msp = doc.modelspace()
    cel = doc.blocks.new("CZESC") if "insert_dxf" in c["odmiany"] else msp
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
    warstwa = "0" if cel is not msp else WID    # w bloku: warstwa "0" dziedziczy warstwę wstawienia
    if c["styl"] == "poly":
        pts = [(*szum(e[1]), e[3] if e[0] == "A" else 0) for e in el]
        cel.add_lwpolyline(pts, format="xyb", close=True, dxfattribs={"layer": "IV_OUTER_PROFILE"})
        if w == "podwojny_kontur":
            cel.add_lwpolyline(pts, format="xyb", close=True, dxfattribs={"layer": "IV_OUTER_PROFILE"})
    else:
        linie = [(szum(e[1]), szum(e[2])) for e in el]
        if w == "przerwa":
            (a, b) = linie[1]
            dl = ((b[0] - a[0]) ** 2 + (b[1] - a[1]) ** 2) ** 0.5
            g = R.uniform(0.3, 1.5) / dl
            linie[1] = (a, (b[0] - (b[0] - a[0]) * g, b[1] - (b[1] - a[1]) * g))
        if c["podzial"]:
            k = max(range(len(linie)), key=lambda i: math.dist(*linie[i]))   # najdłuższa krawędź
            a, b = linie[k]
            m = ((a[0] + b[0]) / 2, (a[1] + b[1]) / 2)
            linie[k:k + 1] = [(a, m), (m, b)]
        for a, b in linie:
            cel.add_line(a, b, dxfattribs={"layer": warstwa})
        if w == "podwojny_kontur":
            for a, b in linie:
                cel.add_line(b, a, dxfattribs={"layer": warstwa})
    if w == "nieprzycieta_faza":
        k = next(iter(c["fazy"]))
        f = [e for e in kontur(c) if e[0] == "L"]
        rog = [(0, 0), (c["W"], 0), (c["W"], c["H"]), (0, c["H"])][k]
        fz = min(f, key=lambda e: abs(e[1][0] - rog[0]) + abs(e[1][1] - rog[1]) + abs(e[2][0] - rog[0])
                 + abs(e[2][1] - rog[1]))   # odcinek fazy przy tym narożniku w geometrii z rysunku
        msp.add_line(fz[1], fz[2], dxfattribs={"layer": "0"})
    if w == "zly_ksztalt":     # prawy górny narożnik cofnięty o 4 mm — gabaryt się nie zmienia
        for e in cel.query("LINE LWPOLYLINE"):
            if e.dxftype() == "LINE":
                for a in ("start", "end"):
                    v = e.dxf.get(a)
                    if abs(v.y - c["H"]) < 1e-3 and v.x > W - 30:
                        e.dxf.set(a, (v.x - 4, v.y, 0))
            else:
                pts = [(x - 4 if abs(y - c["H"]) < 1e-3 and x > W - 30 else x, y, sw, ew, b)
                       for x, y, sw, ew, b in e.get_points()]
                e.set_points(pts)
    warstwa_otw = "IV_INTERIOR_PROFILES" if c["styl"] == "poly" else warstwa
    for i, (x, y, d) in enumerate(c["otwory"]):
        if w == "brak_otworu_w_dxf" and i == 0:
            continue
        if w == "zly_otwor" and i == 0:
            d += 2
        if w == "przesuniety_otwor" and i == 0:
            x += 5
        cel.add_circle(szum((x, y)), d / 2, dxfattribs={"layer": warstwa_otw})
        msp.add_line((x - d, y), (x + d, y), dxfattribs={"layer": OSIE})
        msp.add_line((x, y - d), (x, y + d), dxfattribs={"layer": OSIE})
        if w == "pogłębienie" and i == 0:
            cel.add_circle((x, y), d / 2 + 2, dxfattribs={"layer": warstwa_otw})
    if w == "otwor_bez_wymiaru":
        msp.add_circle((c["W"] / 2, c["H"] / 3), 7.5, dxfattribs={"layer": "0"})
    if w == "wystajaca_linia":
        msp.add_line((c["W"] + 40, 0), (c["W"] + 40, c["H"] / 2), dxfattribs={"layer": "0"})
    if "p" in c["flags"]:
        msp.add_line((W / 2, 0), (W / 2, c["H"]), dxfattribs={"layer": "IV_BEND"})
    if c["styl"] == "linie" and R.random() < 0.4:
        msp.add_line((0, c["H"] / 2), (W, c["H"] / 2), dxfattribs={"layer": UKR})
    if cel is not msp:
        msp.add_blockref("CZESC", (0, 0), dxfattribs={"layer": WID})
    if c["obrot_dxf"]:
        for e in msp:
            e.transform(ezdxf.math.Matrix44.z_rotate(math.pi / 2))
    if w == "dxf_lustro":
        for e in msp:
            e.transform(ezdxf.math.Matrix44.scale(-1, 1, 1))
    doc.saveas(sciezka)


def mm(x):
    return f"{x:g}".replace(".", ",")


def zapisz_pdf(c, sciezka, R):
    """Rysunek A4 poziomo: widok części w skali 1:k grubą linią, wymiary cienką, widok z boku,
    ramka, tabliczka i notki jako tekst — plus odmiany z c["odmiany"]."""
    od = c["odmiany"]
    gr, cien = (0.7, 0.25) if "inne_grubosci" in od else (0.54, 0.36)
    kat = math.radians(R.uniform(15, 75)) if "obrot_widoku" in od else 0.0
    zasieg = (abs(c["W"] * math.cos(kat)) + abs(c["H"] * math.sin(kat)),
              abs(c["W"] * math.sin(kat)) + abs(c["H"] * math.cos(kat)))
    k = next(k for k in (1, 2, 2.5, 5, 7.5, 10, 15, 20, 25, 30, 40)
             if max(zasieg[0] / 480, zasieg[1] / 300) * 72 / 25.4 <= k)
    s = 72 / 25.4 / k
    doc = fitz.open()
    strona = doc.new_page(width=842, height=595)
    x0, y0 = 90 + c["H"] * math.sin(kat) * s, 110 + zasieg[1] * s     # obraz punktu (0, 0) części

    def P(p):
        x, y = p
        xr, yr = x * math.cos(kat) - y * math.sin(kat), x * math.sin(kat) + y * math.cos(kat)
        return fitz.Point(x0 + xr * s, y0 - yr * s)

    for e in kontur(c):
        if e[0] == "L":
            strona.draw_line(P(e[1]), P(e[2]), width=gr)
        else:   # łuk 90° z bulge -> przybliżenie odcinkami (Inventor daje Beziera, oba są wektorowe)
            (x1, y1), (x2, y2) = e[1], e[2]
            cx, cy = (x1 + x2) / 2 - (y2 - y1) / 2, (y1 + y2) / 2 + (x2 - x1) / 2   # środek po stronie części
            r = math.dist((cx, cy), (x1, y1))
            a1, a2 = math.atan2(y1 - cy, x1 - cx), math.atan2(y2 - cy, x2 - cx)
            if a2 < a1:
                a2 += 2 * math.pi
            pts = [P((cx + r * math.cos(t), cy + r * math.sin(t))) for t in [a1 + (a2 - a1) * i / 12 for i in range(13)]]
            strona.draw_polyline(pts, width=gr)
    for x, y, d in c["otwory"]:
        strona.draw_circle(P((x, y)), d / 2 * s, width=gr)
        if "osie_grube" in od:            # oś otworu tą samą grubą linią, dotyka okręgu
            strona.draw_line(P((x - d, y)), P((x + d, y)), width=gr)
            strona.draw_line(P((x, y - d)), P((x, y + d)), width=gr)
        else:
            strona.draw_line(P((x - d, y)), P((x + d, y)), width=cien)
    if "u" in c["flags"]:      # krawędź ukosu 7 mm od dolnej krawędzi
        f0, f1 = c["fazy"].get(0, 0), c["fazy"].get(1, 0)
        strona.draw_line(P((max(f0, 7), 7)), P((c["W"] - max(f1, 7), 7)), width=gr)
    if "p" in c["flags"]:
        strona.draw_line(P((c["W"] / 2, 0)), P((c["W"] / 2, c["H"])), width=cien, dashes="[6 2 1 2] 0")
    if "ukryte_kreskowane" in od:   # kieszeń ukryta po drugiej stronie — kreskowana, ta sama grubość
        x1, y1, x2, y2 = c["W"] * 0.3, c["H"] * 0.3, c["W"] * 0.6, c["H"] * 0.6
        for a, b in (((x1, y1), (x2, y1)), ((x2, y1), (x2, y2)), ((x2, y2), (x1, y2)), ((x1, y2), (x1, y1))):
            strona.draw_line(P(a), P(b), width=gr, dashes="[4 2] 0")
    if "przekroj_dlugi" in od:      # linia przekroju A-A: gruba, przecina obrys, wystaje daleko poza widok
        strona.draw_line(P((-0.6 * c["W"], c["H"] * 0.5)), P((1.6 * c["W"], c["H"] * 0.5)), width=gr)
    # wymiary (cienkie, bez dotykania konturu)
    for i in range(30):
        strona.draw_line(P((0, -8 - i % 3)), P((c["W"], -8 - i % 3)), width=cien)
    strona.draw_line(P((-8, 0)), P((-8, c["H"])), width=cien)
    # widok z boku: prostokąt W x t, i ramka arkusza
    strona.draw_rect(fitz.Rect(500, 520, 500 + min(c["W"] * s, 300), 520 + max(c["t"] * s, 1)), width=gr)
    strona.draw_rect(fitz.Rect(20, 20, 822, 575), width=0.72)
    tekst = ([] if "bez_etykiety" in od else [f"VIEW1 ( 1 : {mm(k)} )"]) + [
        f"PL {c['t']} x {c['W']} x {c['H']}", "S355", f"Ilość: {c['qty']}", mm(c["W"]), mm(c["H"])]
    rozm = {}
    for f in c["fazy"].values():
        rozm[f] = rozm.get(f, 0) + 1
    for f, n in rozm.items():
        if c["fazy_notka"]:           # notka przy pierwszym narożniku z tą fazą (jak odnośnik w Inventorze)
            rog = next(r for r, v in c["fazy"].items() if v == f)
            x, y = [(0, 0), (c["W"], 0), (c["W"], c["H"]), (0, c["H"])][rog]
            pkt = P((x, y))
            strona.insert_text((pkt.x + 8, pkt.y - 8), (f"{n}x " if n > 1 else "") + f"{f} x 45°",
                               fontname="dv", fontfile=FONT, fontsize=7)
        else:
            tekst.append(mm(f))
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
        strona.insert_text((640, 50 + 12 * i), l, fontname="dv", fontfile=FONT, fontsize=8)
    if "strona_obrocona" in od:
        strona.set_rotation(90)
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
    S.WERYFIKACJA_DXF.clear()
    for f in sorted({c["folder"] for c in czesci}):
        S.sprawdz(os.path.join(baza, f))
    wiersze, status = {}, {w["part"]: w for w in S.WERYFIKACJA_DXF}
    for w in S.WSZYSTKIE_WIERSZE:
        wiersze.setdefault(w["part"], []).append(w)
    wynik = dict(wykryte=[], tylko_uwaga=[], przeocz=[], zly_typ=[], wada_zgodna=[],
                 falsz_bl=[], falsz_uw=[], dobre_niezweryf=[])
    for c in czesci:
        ws = wiersze.get(c["part"], [])
        bl = " | ".join(w["opis"] for w in ws if w["poziom"] == "BŁĄD")
        uw = " | ".join(w["opis"] for w in ws if w["poziom"] == "UWAGA")
        c["status"] = status.get(c["part"], {}).get("status", "?")
        powod = status.get(c["part"], {}).get("powod", "")
        if c["wada"] is None:
            if bl:
                wynik["falsz_bl"].append((c, bl))
            if uw:
                wynik["falsz_uw"].append((c, uw))
            if c["status"] != "ZGODNY 1:1":
                wynik["dobre_niezweryf"].append((c, f"{c['status']}: {powod}"))
            continue
        if c["status"] == "ZGODNY 1:1":
            wynik["wada_zgodna"].append((c, "część z wadą dostała status ZGODNY 1:1"))
        wzor = WYKRYCIE[c["wada"]]
        tam = uw if POZIOM.get(c["wada"]) == "UWAGA" else bl
        if re.search(wzor, tam):
            wynik["wykryte"].append(c)
        elif re.search(wzor, uw):
            wynik["tylko_uwaga"].append((c, uw))
        elif bl:
            wynik["zly_typ"].append((c, bl))
        else:
            wynik["przeocz"].append((c, f"{c['status']}: {powod} | {uw}"))
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

    n_wad = 4 * len(WADY)
    print(f"Pliki testowe: {baza}\nRaport programu: {raport}\n")
    print(f"Części z wadą: {n_wad}, bez wady: {100 - n_wad}")
    print(f"  wady wykryte (właściwy komunikat):    {len(wynik['wykryte'])}/{n_wad}")
    print(f"  wykryte tylko jako UWAGA:             {len(wynik['tylko_uwaga'])}")
    print(f"  BŁĄD, ale innego rodzaju:             {len(wynik['zly_typ'])}")
    print(f"  przeoczone:                           {len(wynik['przeocz'])}")
    print(f"  część z wadą ze statusem ZGODNY 1:1:  {len(wynik['wada_zgodna'])}")
    print(f"  fałszywe BŁĘDY (dobre części):        {len(wynik['falsz_bl'])}")
    print(f"  UWAGI na dobrych częściach:           {len(wynik['falsz_uw'])}")
    print(f"  dobre części bez statusu ZGODNY 1:1:  {len(wynik['dobre_niezweryf'])}")
    print("\nWg typu wady (wykryte / 4):")
    for t in WADY:
        print(f"  {t:20s} {sum(1 for c in wynik['wykryte'] if c['wada'] == t)}/4")
    print("\nOdmiany rysunku/DXF na częściach poprawnych (ZGODNY 1:1 / ile):")
    for o in ODMIANY:
        dobre = [c for c in czesci if c["wada"] is None and o in c["odmiany"]]
        print(f"  {o:20s} {sum(1 for c in dobre if c['status'] == 'ZGODNY 1:1')}/{len(dobre)}")
    if S.NAKLADKI_PDF:
        print(f"\nPodgląd nakładek: {S.zapisz_nakladki_pdf(os.path.join(baza, 'NAKLADKI_TEST_100.pdf'))}")
    for klucz, tytul in (("tylko_uwaga", "TYLKO UWAGA"), ("zly_typ", "INNY RODZAJ"), ("przeocz", "PRZEOCZONE"),
                         ("wada_zgodna", "WADA A ZGODNY"), ("falsz_bl", "FAŁSZYWE BŁĘDY"),
                         ("falsz_uw", "UWAGI NA DOBRYCH"), ("dobre_niezweryf", "DOBRA NIEZWERYF.")):
        for c, opis in wynik[klucz]:
            print(f"  [{tytul}] {c['part']} {c['wada'] or ''} styl={c['styl']} odmiany={sorted(c['odmiany'])} "
                  f"fazy={c['fazy']} -> {opis[:300]}")
    return wynik


def zaliczony(w):
    # Wada zgłoszona tylko jako UWAGA (strefa szara nakładki: tolerancja..2 x tolerancja) jest dopuszczalna —
    # taka część i tak nie dostaje statusu "ZGODNY 1:1" (sprawdza to "wada_zgodna"). Najważniejsze:
    # żadna wada nie przechodzi jako zgodna, a dobre części nie mają fałszywych alarmów.
    return (len(w["wykryte"]) + len(w["tylko_uwaga"]) == 4 * len(WADY) and not w["przeocz"] and not w["zly_typ"]
            and not w["falsz_bl"] and not w["falsz_uw"] and not w["dobre_niezweryf"] and not w["wada_zgodna"])


def test_100_czesci():
    w = main()
    assert zaliczony(w), w


if __name__ == "__main__":
    sys.exit(0 if zaliczony(main()) else 1)
