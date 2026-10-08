# -*- coding: utf-8 -*-
"""
sprawdz_hybrydowo_GEO.py v5 (v4 + kontur DXF vs rysunek) — jak v3 (geometria DXF), PLUS kontrola geometrii 3D z pliku STP
(grubość i gabaryt każdej części z rzeczywistej bryły, nie tylko z opisu/nazwy pliku) ORAZ
czytelny, strukturalny raport zbiorczy .xlsx (kolorowane błędy/uwagi, sortowanie po pozycji BOM
= dokładne miejsce błędu w strukturze łyżki), zamiast samej ściany tekstu w konsoli.

Porównuje krzyżowo: BOM (xlsx) <-> nazwy plików DXF/PDF <-> zawartość PDF
(tabliczka + stempel "Ilość:" z PdfRenamera) <-> geometria DXF (2D, rozwinięcie/kontur)
<-> geometria STP (3D, rzeczywista bryła każdej części w zespole, jeśli plik STP istnieje).

Reguły wyprowadzone z CAŁEGO archiwum S:\\14_Gotowe_Projekty\\GJERSTAD
(154 wydania hybrydowe, ~4900 pozycji BOM; 2026-07-03), plus:
2026-07-18 — dodano kontrolę STP i raport xlsx na prośbę użytkownika ("połącz sprawdzenie
STP-BOM ze sprawdz hybrydowo GEO, żeby wszystkie błędy były w jednym raporcie i były
czytelniejsze / pokazywały gdzie jest błąd").
2026-09-30 — różnice z pliku .stp nie są już BŁĘDAMI, tylko osobnym poziomem "INFO STP" (niebieskie
wiersze): gabaryt bryły z STP zależy od ułożenia części w złożeniu (obrót, pochylenie, gięcie), więc
na GE98022-S2 dał 6 "błędów" przy w 100% poprawnej dokumentacji (każdy odtworzony samym obrotem
płaskiego konturu z DXF, błąd <= 0,7 mm). O zgodności decydują DXF, PDF i BOM. Na końcu arkusza
raportu jest podsumowanie: czy wszystkie DXF są zgodne z BOM + objaśnienie wierszy STP.
2026-10-08 (v5) — kontrola KONTURU DXF vs RYSUNEK PDF (zgłoszenie: "DXF nie pokrywa się z rysunkiem,
fazy zepsute" — v4 porównywała tylko gabaryt z BOM, więc zła faza przechodziła): przerwy i wolne końce
konturu, rozgałęzienia (nieprzycięty narożnik po fazie, linie fazy/ukosu krawędzi), podwójne linie,
elementy dorysowane na warstwie 0, fazy narożników vs notki faz na rysunku ("4x 10 x 45°"), otwory vs
Ø na rysunku, współśrodkowe okręgi (pogłębienie otworu w DXF), gabaryt DXF vs tabliczka PDF, gdy BOM
nie ma wymiarów. Szczegóły i progi: sekcja "kontur DXF vs rysunek" niżej. Testy: testy/test_kontur.py.

  Kategorie części (z DESCRIPTION + flag + struktury BOM):
  - blacha "PL t x a x b" (Normal)          -> DXF + PDF (99% archiwum),
      ALE z flagą C (cięta z płaskownika/lemiesza) -> tylko PDF (76%), DXF dopuszczalny;
      krawędź typu FRONT/CUTTING EDGE z wąskiego paska (szer. <=400) bez DXF -> uwaga.
  - "PL t" / przygotówka PRZ                -> DXF wymagany, PDF opcjonalny (48/47%).
  - lemiesz LM / pręt PR / rura RU/RØR / Round Bar / stock BL -> tylko PDF; DXF dopuszczalny
      (lemiesze HALF ARROW L/P mają DXF, środkowe nie).
  - złożenie (flaga M)                      -> PDF (__X__UNKNOWN__m); DXF przy złożeniu = uwaga
      (loadery G BOLTON mają DXF złożenia - szablon).
  - ramka tabliczki (DEK-*/PG4104777, title FRAME..SIGN) -> PDF bywa albo nie; tylko uwaga.
  - Purchased / STICKER                     -> pliki opcjonalne (93% bez plików).

  Zgodność metadanych:
  - grubość w nazwie musi być JEDNYM z wymiarów DESCRIPTION (kolejność t/a/b bywa różna,
    np. "PL 123 x 90 x 3" = grubość 3; "PL 60 x 186 x 35" = grubość 60);
  - ilość w nazwie/stemplu = QTY wiersza LUB ilość całkowita (QTY x rodzice wg POS,
    np. poz. 7.3.2 przy rodzicu x2 -> pliki nazwane 2 przy QTY=1);
  - flagi w nazwie = kolumny M/O/P/U/C (m=montaż, o=obróbka, p=gięcie, u=ukosowanie, c=cięcie);
  - stempel "Ilość: N" != QTY jest OK przy adnotacji o odbiciu lustrzanym
    ("DRUGA SZTUKA ODBICIE LUSTRZANE");
  - DXF ma prefiks głównego złożenia (numer z BOM, NIE nazwa folderu), PDF bez prefiksu;
  - pliki "P10.01 WYTYCZNE GJERSTAD" itp. ignorowane.

  Geometria STP (opcjonalna, wymaga Inventora zainstalowanego lokalnie + pywin32):
  - szuka pliku "<nazwa_folderu_bez_dopiskow>.stp"/".step" w folderze projektu, albo
    dowolnego *.stp/*.step jeśli nazwa nie pasuje (znaleziono 182 takie pliki w archiwum
    GJERSTAD 2026-07-18, w większości nazwane jak folder projektu, np. GE6725\\GE6725.stp);
  - otwiera plik przez COM (win32com) w Inventorze (używa już uruchomionej instancji, jeśli
    jest, inaczej odpala nową w tle); pliki STP z tego archiwum zachowują strukturę zespołu
    (PRODUCT('PG4104413','PG4104413','PL 12 x 1633 x 1406',...) w danych STEP) — Inventor
    powinien zaimportować to jako ZESPÓŁ z osobnymi komponentami, nie jedną zlaną bryłę;
  - dla każdej liściastej części czyta LOKALNY (nieobrócony) RangeBox definicji części
    (PartComponentDefinition.RangeBox — w układzie współrzędnych samej części, NIE zespołu,
    więc obrót/pozycja części w złożeniu nie zniekształca wyniku, w przeciwieństwie do
    RangeBox samego wystąpienia w przestrzeni zespołu);
  - 3 wymiary lokalnego pudełka porównywane wprost (posortowane) z 3 wymiarami z DESCRIPTION
    BOM, tą samą tolerancją co geometria DXF (GEO_TOL_BLAD/GEO_TOL_INFO), z tym samym
    złagodzeniem dla części giętych (flaga P) i obrabianych (flagi O/U) — patrz komentarz
    przy geo_porownaj_3d.
  - ZWERYFIKOWANE NA ŻYWO 2026-07-18 (COM przez Inventor 2026.1, plik GE75206-S1.stp z tego
    archiwum): Inventor faktycznie importuje ten STEP jako ZESPÓŁ (DocumentType=12291) z 64
    osobnymi wystąpieniami, iProperty "Part Number" poprawnie = numer części z BOM (np.
    'PG4104333'). Jedyna pułapka: `Documents.Open()` zwraca ogólny obiekt "Document" bez
    `.ComponentDefinition` (early-bound gen_py nie widzi interfejsu AssemblyDocument) —
    rozwiązanie: `win32com.client.CastTo(oDoc, "AssemblyDocument")` (oficjalny mechanizm
    pywin32 na ten dokładnie przypadek). Sprawdzone wyniki dla GE75206-S1 są SPÓJNE z
    niezależnie znalezionym błędem tabliczki PDF: bryła STP `PG4104333` (BOTTOM PLATE) =
    1226.0 mm, dokładnie zgadza się z tabliczką PDF (1226), a NIE z BOM (1230) — czyli to
    naprawdę BOM ma nieaktualną wartość, nie model/PDF. Części gięte (flaga P) mają lokalny
    RangeBox mniejszy niż DESCRIPTION (bo bryła jest już wygięta, a DESCRIPTION to zwykle
    rozwinięcie) - stąd złagodzenie do UWAGI dla flagi P/O/U, tak samo jak przy DXF.

Użycie: sprawdz_hybrydowo_GEO.py <folder> [folder2 ...]  (albo przeciągnij folder(y) na .bat)
Wymaga: pip install openpyxl pymupdf ezdxf pywin32
Na końcu zapisuje jeden zbiorczy raport .xlsx (obok pierwszego podanego folderu) i otwiera go.
"""
import atexit
import datetime
import math
import os
import re
import sys

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    import fitz  # pymupdf
except ImportError as e:
    print("Brak biblioteki:", e, "->  pip install openpyxl pymupdf")
    sys.exit(2)

try:
    import ezdxf
    import numpy as np  # zależność ezdxf, więc jest zawsze razem z nim
    from ezdxf import bbox as ezbbox
    from ezdxf import path as ezpath
    GEO = True
except ImportError:
    GEO = False  # bez ezdxf dziala jak zwykla wersja v2 (bez kontroli geometrii DXF)

try:
    import win32com.client
    STP_MODULE_OK = True
except ImportError:
    STP_MODULE_OK = False  # bez pywin32 dziala bez kontroli geometrii STP

# tolerancje kontroli geometrii (bbox konturu DXF/bryly STP vs wymiary DESCRIPTION)
GEO_TOL_BLAD = 3.0   # mm — odchylka powyzej -> BLAD
GEO_TOL_INFO = 1.5   # mm — strefa szara -> UWAGA

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

FLAG_LETTERS = set("mopuc")
JUNK = re.compile(r"WYTYCZNE|^__?P10\.01", re.I)

# Aliasy nazw kolumn BOM w roznych wariantach jezykowych/eksportow Inventora - zbudowane
# 2026-07-19 ze skanu 38 realnych plikow xlsx w S:\14_Gotowe_Projekty\GJERSTAD (12 wariantow
# naglowka: angielski ITEM/PART NUMBER/QTY/..., polski POZYCJA/NUMER CZĘŚCI/ILOŚĆ/...).
# Kolumna "TEMAT" (widziana obok TYTUŁ w wariancie polskim) to INNE pole niz TITLE
# (najprawdopodobniej "Subject") - celowo pominieta jako alias.
ALIASY_KOLUMN = {
    "POS": ("POS", "ITEM", "POZYCJA"),
    "PART": ("PART NUMBER", "NUMER CZĘŚCI", "NUMER CZESCI"),
    "STRUCT": ("BOM STRUCTURE", "TYP W STRUKTURZE ZESTAWIENIA BOM"),
    "QTY": ("QTY", "ILOŚĆ", "ILOSC"),
    "DESC": ("DESCRIPTION", "OPIS"),
    "MAT": ("MATERIAL", "MATERIAŁ"),
    "TITLE": ("TITLE", "TYTUŁ", "TYTUL"),
}

# zbiorczy raport strukturalny (wszystkie foldery podane w jednym uruchomieniu)
WSZYSTKIE_WIERSZE = []   # list of dict: folder, pos, part, tytul, poziom, plik, rodzaj, opis
PODSUMOWANIE_FOLDEROW = []  # list of dict: folder, ok, bledy, uwagi


# ---------------------------------------------------------------- BOM
def desc_dims(desc):
    """Zbior wymiarow z DESCRIPTION 'PL t x a x b' (kolejnosc bywa dowolna)."""
    m = re.match(r"PL\.?\s*([\d,\.]+(?:\s*[xX]\s*[\d,\.]+)*)", desc.strip())
    if not m:
        return []
    dims = []
    for x in re.split(r"\s*[xX]\s*", m.group(1)):
        x = x.replace(",", ".").strip(".")
        try:
            dims.append(float(x))
        except ValueError:
            pass
    return dims


def _znajdz_kolumne_m(hdr):
    """Kolumna startowa flag M/O/P/U/C. Dopasowuje 'M' oraz warianty z dopiskiem typu
    'M\\n(MONTAŻ)' (znaleziono w archiwum), ale NIE 'MATERIAL'/'MATERIAŁ'/'MINIATURA'."""
    for i, h in enumerate(hdr):
        if re.match(r'^M(\s|\(|$)', h.strip()):
            return i
    return None


def czytaj_bom(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows = []
    for r in ws.iter_rows(values_only=True):
        rows.append(list(r))
        if len(rows) > 500:
            break
    wb.close()
    hdr_i = None
    for i, r in enumerate(rows[:6]):
        up = [str(c or "").strip().upper() for c in r]
        if (any(a in up for a in ALIASY_KOLUMN["PART"])
                and any(a in up for a in ALIASY_KOLUMN["QTY"])):
            hdr_i = i
            break
    if hdr_i is None:
        return None, "nie znaleziono nagłówka BOM (PART NUMBER/QTY, w żadnym znanym wariancie językowym)"
    hdr = [str(c or "").strip().upper() for c in rows[hdr_i]]

    def col(field, default=None):
        for alias in ALIASY_KOLUMN[field]:
            if alias in hdr:
                return hdr.index(alias)
        return default

    c_pos = col("POS", default=0)
    c_part = col("PART", default=1)
    c_struct = col("STRUCT", default=4)
    c_qty = col("QTY", default=5)
    c_desc = col("DESC", default=7)
    c_mat = col("MAT", default=10)
    c_title = col("TITLE")
    c_m = _znajdz_kolumne_m(hdr)
    if c_m is None:
        c_m = 11  # fallback jak w wersji bez rozpoznanych aliasow
    bom = []
    for r in rows[hdr_i + 1:]:
        if c_part >= len(r) or not r[c_part]:
            continue
        desc = str(r[c_desc] or "").strip() if c_desc < len(r) else ""
        bom.append(dict(
            pos=str(r[c_pos]).strip() if r[c_pos] is not None else "?",
            part=str(r[c_part]).strip(),
            struct=str(r[c_struct] or "").strip() if c_struct < len(r) else "",
            qty=r[c_qty] if c_qty < len(r) else None,
            desc=desc, dims=desc_dims(desc),
            mat=str(r[c_mat] or "").strip() if c_mat < len(r) else "",
            title=str(r[c_title] or "").strip() if (c_title is not None and c_title < len(r)) else "",
            flags={f for off, f in enumerate("mopuc") if c_m + off < len(r) and r[c_m + off]},
        ))
    # ilosc calkowita: iloczyn QTY rodzicow wg POS (poz. "7.3.2" -> rodzice "7.3", "7")
    qty_by_pos = {b["pos"]: b["qty"] for b in bom}
    for b in bom:
        total = b["qty"] if isinstance(b["qty"], (int, float)) else None
        if total is not None and "." in b["pos"]:
            czesci = b["pos"].split(".")
            for k in range(1, len(czesci)):
                q = qty_by_pos.get(".".join(czesci[:k]))
                if isinstance(q, (int, float)):
                    total *= q
        b["qty_total"] = total
    return bom, None


def kategoria(b, proj_root):
    d = b["desc"].upper()
    t = (b["title"] or "").upper()
    part = b["part"].upper()
    if b["struct"].lower() == "purchased" or "STICKER" in part or "STICKER" in t:
        return "purchased"
    if part.startswith("DEK-") and "FRAME" in t or ("FRAME" in t and "SIGN" in t) or part == "PG4104777":
        return "ramka"
    if "m" in b["flags"] or part == proj_root:
        return "zlozenie"
    if d.startswith("LM") or re.match(r"(PR\d|PR\s|RU\d|RUND|ROUND\s*BAR|RØR|ROR\b|AKSEL)", d) or d.startswith("BL"):
        return "material_ciety"      # lemiesz/pret/rura/stock: PDF, DXF opcjonalny
    if re.match(r"PL\.?\s*[\d,\.]+\s*[xX]\s*[\d,\.]+\s*[xX]", d):
        s = sorted(b["dims"])
        if len(s) == 3 and s[0] == s[1] and s[2] >= 3 * s[0]:
            return "material_ciety"  # przekroj kwadratowy (np. PL 80 x 80 x 692) = pret, nie blacha
        if "c" in b["flags"]:
            return "blacha_c"        # cieta z plaskownika: PDF, DXF opcjonalny
        return "blacha"              # DXF + PDF
    if re.match(r"PL\.?\s*[\d,\.]+(\s*[xX]\s*[\d,\.]+)?\s*(MM)?\s*$", d):
        return "przygotowka"         # DXF, PDF opcjonalny
    if "c" in b["flags"]:
        return "material_ciety"
    return "inne"                    # PDF oczekiwany


def czy_krawedz_plaskownik(b):
    """FRONT/CUTTING EDGE z waskiego paska (szer<=400): brak DXF czesty w archiwum (grading)."""
    t = (b["title"] or "").upper()
    if not re.search(r"FRONT\s*EDGE|CUTTING\s*EDGE|KNIFE", t):
        return False
    dims = sorted(b["dims"], reverse=True)
    return len(dims) == 3 and dims[1] <= 400


# ---------------------------------------------------------------- nazwy plikow
def parsuj_nazwe(fname, bom_parts, bom_asm):
    base, ext = os.path.splitext(fname)
    ext = ext.upper()
    toks = base.split("__")
    while toks and toks[0] == "":
        toks.pop(0)
    if not toks:
        return None

    def meta_like(t):
        return (t == "" or t.upper() == "UNKNOWN" or t.isdigit()
                or re.fullmatch(r"[\d,\.]+\s*mm", t)
                or (len(t) == 1 and t.lower() in FLAG_LETTERS))

    pi = 0
    if len(toks) > 1 and not meta_like(toks[1]):
        t1u = toks[1].upper()
        if (t1u in bom_parts
                or (t1u.endswith("PRZ") and t1u[:-3] in bom_parts)
                or toks[0].upper() in bom_asm):
            pi = 1
    part, rest = toks[pi], toks[pi + 1:]
    prefiks = pi == 1
    if rest and rest[0] == "UNKNOWN":
        return dict(part=part, thk=None, mat=None, qty=None,
                    flags={t.lower() for t in rest[1:] if t}, asm=True, prefiks=prefiks, ext=ext)
    thk = mat = qty = None
    flags = set()
    for t in rest:
        if not t:
            continue
        m = re.fullmatch(r"([\d,\.]+)\s*mm", t)
        if m and thk is None:
            thk = float(m.group(1).replace(",", "."))
        elif len(t) == 1 and t.lower() in FLAG_LETTERS:
            flags.add(t.lower())
        elif t.isdigit():
            qty = int(t)      # ostatni czysto-liczbowy token = ilosc
        elif thk is None and re.fullmatch(r"\d+[,\.]\d+", t):
            thk = float(t.replace(",", "."))  # grubosc zapisana bez "mm" (starsze wydania)
        elif mat is None:
            mat = t
    return dict(part=part, thk=thk, mat=mat, qty=qty, flags=flags,
                asm=False, prefiks=prefiks, ext=ext)


def mat_zgodny(a, b):
    if not a or not b:
        return True
    def norm(x):
        x = re.sub(r"[\s\.]", "", x).upper()
        return x.replace("HARDOX", "HB").replace("HX", "HB")  # Hardox 500 == HB500 == HX500
    na, nb = norm(a), norm(b)
    return na.startswith(nb) or nb.startswith(na)


# ---------------------------------------------------------------- PDF
def czytaj_pdf(path):
    doc = fitz.open(path)
    text = "\n".join(doc[i].get_text() for i in range(len(doc)))
    doc.close()
    out = {}
    m = re.search(r"Ilość\s*:\s*(\d+)", text)
    if m:
        out["stamp"] = int(m.group(1))
    if re.search(r"ODBICI|LUSTRZ|MIRROR", text, re.I):
        out["lustro"] = True
    m = re.search(r"PL\.?\s*([\d,\.]+\s*x\s*[\d,\.]+\s*x\s*[\d,\.]+)", text)
    if m:
        dims = []
        for x in re.split(r"\s*x\s*", m.group(1)):
            try:
                dims.append(float(x.replace(",", ".").strip(".")))
            except ValueError:
                pass
        if len(dims) == 3:
            out["dims"] = sorted(dims)
    m = re.search(r"(S355(?:J2G3|J2|MC)?|HARDOX\s*\d{3}|HB\s*\d{3})", text, re.I)
    if m:
        out["mat"] = m.group(1)
    out.update(wymiary_rysunku(text))
    return out


LICZBA = r"(\d+(?:[.,]\d+)?)"


def _liczba(s):
    return float(s.replace(",", "."))


def wymiary_rysunku(text):
    """Wymiary z tekstu rysunku PDF potrzebne do porównania z DXF:
    - fazy: notki Inventora "5 x 45°", "2x 5 x 45°" (ilość z przodu), "FAZA 5x5";
      nogi fazy = (d, d*tg(kąt)), posortowane;
    - fi: średnice otworów "Ø10,5" — Inventor pisze Ø czcionką AIGDT, która w tekście PDF
      daje literę "n" ("n10,5" = Ø10,5);
    - gwinty: "M12" (w DXF jest wtedy otwór pod gwint, ~0,75-1,0 x M);
    - promienie: "R15" (duże wycięcia okrągłe bywają wymiarowane promieniem);
    - liczby: wszystkie liczby z rysunku (do sprawdzenia, czy faza z DXF jest zwymiarowana)."""
    fazy = []
    for m in re.finditer(r"(?<![\d.,])(?:(\d+)\s*[xX×]\s*)?" + LICZBA + r"\s*(?:mm)?\s*[xX×]\s*"
                         + LICZBA + r"\s*(?:deg|[°º˚])", text):
        d, kat = _liczba(m.group(2)), _liczba(m.group(3))
        if d <= 0 or not 0 < kat < 90:
            continue
        nogi = tuple(sorted((round(d, 2), round(d * math.tan(math.radians(kat)), 2))))
        fazy.append(dict(nogi=nogi, ile=int(m.group(1)) if m.group(1) else None,
                         tekst=re.sub(r"\s+", " ", m.group(0)).strip()))
    for m in re.finditer(r"FAZ[AY]?\s+" + LICZBA + r"\s*[xX×]\s*" + LICZBA
                         + r"(?![\d.,])(?!\s*(?:deg|[°º˚]))", text, re.I):
        a, b = _liczba(m.group(1)), _liczba(m.group(2))
        nogi = (a, a) if b == 45 else tuple(sorted((a, b)))
        fazy.append(dict(nogi=nogi, ile=None, tekst=re.sub(r"\s+", " ", m.group(0)).strip()))
    fi = [_liczba(x) for x in re.findall(r"(?:[Ø⌀ø∅]|(?:(?<=[\s\dxX×(])|^)n)\s?" + LICZBA, text, re.M)]
    gwinty = [_liczba(x) for x in re.findall(r"(?<![A-Za-z])M" + LICZBA, text)]
    promienie = [_liczba(x) for x in re.findall(r"(?<![A-Za-z])R\s?" + LICZBA, text)]
    liczby = {_liczba(x) for x in re.findall(r"\d+(?:[.,]\d+)?", text)}
    return dict(fazy=fazy, fi=fi, gwinty=gwinty, promienie=promienie, liczby=liczby)


# ---------------------------------------------------------------- geometria DXF
def dxf_gabaryt(path):
    """Kandydackie obrysy konturu DXF: bbox osiowy + prostokąty we wszystkich
    orientacjach krawędzi otoczki wypukłej. DESCRIPTION Inventora to gabaryt
    w osiach części — przy obróconym eksporcie (SIDE KNIFE) albo części
    trójkątnej (SIDE MEMBER) właściwa orientacja to któraś z krawędzi konturu."""
    import math
    from ezdxf import path as ezpath
    doc = ezdxf.readfile(path)
    pts = []
    for e in doc.modelspace():
        if e.dxftype() in ("TEXT", "MTEXT", "DIMENSION", "POINT", "ATTRIB"):
            continue
        try:
            p = ezpath.make_path(e)
            pts.extend((round(v.x, 2), round(v.y, 2)) for v in p.flattening(0.3))
        except Exception:
            pass
    if not pts:
        return []
    xs = [p[0] for p in pts]; ys = [p[1] for p in pts]
    kand = [(max(xs) - min(xs), max(ys) - min(ys))]
    spts = sorted(set(pts))
    if len(spts) <= 2:
        return kand
    def cr(o, a, b):
        return (a[0]-o[0])*(b[1]-o[1]) - (a[1]-o[1])*(b[0]-o[0])
    lo = []
    for p in spts:
        while len(lo) >= 2 and cr(lo[-2], lo[-1], p) <= 0:
            lo.pop()
        lo.append(p)
    up = []
    for p in reversed(spts):
        while len(up) >= 2 and cr(up[-2], up[-1], p) <= 0:
            up.pop()
        up.append(p)
    hull = lo[:-1] + up[:-1]
    n = len(hull)
    for i in range(n):
        x1, y1 = hull[i]; x2, y2 = hull[(i + 1) % n]
        dx, dy = x2 - x1, y2 - y1
        L = math.hypot(dx, dy)
        if L < 1e-9:
            continue
        ux, uy = dx / L, dy / L
        us = [p[0]*ux + p[1]*uy for p in hull]
        vs = [-p[0]*uy + p[1]*ux for p in hull]
        kand.append((max(us) - min(us), max(vs) - min(vs)))
    return kand


def geo_porownaj(b, dxf_info, kand):
    """Najlepsze dopasowanie kandydatów obrysu DXF do wymiarów płaskich z DESCRIPTION.
    Zwraca (odchylka, plan_dims, obrys)."""
    dims = list(b["dims"])
    warianty = []
    if dxf_info["thk"] is not None and dxf_info["thk"] in dims:
        d2 = list(dims)
        d2.remove(dxf_info["thk"])             # grubość znana z nazwy pliku
        warianty.append(sorted(d2))
    else:
        for i in range(3):
            warianty.append(sorted(dims[:i] + dims[i + 1:]))
    best = None
    for plan in warianty:
        for (w, h) in kand:
            bbs = sorted((w, h))
            dev = max(abs(plan[0] - bbs[0]), abs(plan[1] - bbs[1]))
            if best is None or dev < best[0]:
                best = (dev, plan, bbs)
    return best


def geo_porownaj_3d(b_dims, dims3):
    """Jak geo_porownaj, ale dla bryły 3D z STP: mamy already 3 prawdziwe wymiary lokalnego
    pudełka (nie trzeba zgadywać orientacji jak przy płaskim DXF) - proste porównanie
    posortowanych trójek. Zwraca (odchylka, plan_posortowany, znaleziony_posortowany)
    albo None jeśli brak 3 wymiarów po którejkolwiek stronie."""
    plan = sorted(b_dims)
    znaleziony = sorted(dims3)
    if len(plan) != 3 or len(znaleziony) != 3:
        return None
    dev = max(abs(p - f) for p, f in zip(plan, znaleziony))
    return dev, plan, znaleziony


# ---------------------------------------------------------------- kontur DXF vs rysunek (fazy, otwory)
# 2026-10-08 — przełożony zgłosił DXF niezgodny z rysunkiem i "zepsute fazy", których program nie
# wyłapał: wcześniejsza kontrola porównywała tylko GABARYT konturu z BOM, a faza 5 mm prawie nie
# zmienia gabarytu. Ta część czyta sam kontur, tak jak zobaczy go wypalarka:
#   - przerwy / wolne końce linii (kontur niedomknięty, linia wystaje za narożnik),
#   - rozgałęzienia (linia kończy się na środku innej albo 3+ linie w jednym punkcie) — typowy
#     ślad fazy dorysowanej bez przycięcia starego narożnika, albo linii fazy/ukosu krawędzi
#     wyeksportowanej z widoku, którą laser by wyciął,
#   - podwójne linie, elementy dorysowane ręcznie na warstwie "0" (opisane w PORADNIKU od
#     03.07.2026, ale w wersji v4 nie było ich w kodzie),
#   - fazy narożników i otwory z DXF porównane z notkami/wymiarami na rysunku PDF,
#   - współśrodkowe okręgi (faza/pogłębienie otworu wyeksportowane jako drugi okrąg).
STYK_TOL = 0.05      # mm — końce bliżej siebie = połączone (CAM i tak je skleja)
PRZERWA_MAX = 2.0    # mm — dwa wolne końce bliżej siebie = "przerwa", dalej = "wolny koniec"
FAZA_MAX = 60.0      # mm — dłuższy odcinek nie jest traktowany jako faza narożnika
FAZA_TOL = 0.5       # mm — tolerancja nogi fazy DXF vs rysunek
FAZA_MIN_RYS = 2.0   # mm — fazy z rysunku do tej wielkości to łamanie krawędzi, nie kontur
OTWOR_TOL = 0.15     # mm — tolerancja średnicy otworu DXF vs Ø z rysunku
WYSTAJE_MAX = 15.0   # mm — dorysowany element dalej poza obrysem = zbłąkana geometria

# warstwy, które nie są cięte (gięcie, linie ukryte/osie/styczne, wymiary, ramki, trasowanie);
# warstwa "Widoczne wąskie" to m.in. 3/4 okręgu gwintu — otwarty łuk, nie kontur
WARSTWY_NIE_TNACE = re.compile(
    r"BEND|TANGENT|TOOL_CENTER|ARC_CENTER|ALTREP|UNCONSUMED|ROLL|FEATURE|GI[EĘ]CI|"
    r"UKRYT|HIDDEN|OSI|O[SŚ]\b|CENTER|CENTRE|SYMETR|STYCZN|W[AĄ]SK|"
    r"WYMIAR|DIM|TEXT|TEKST|ANNOT|TYTU|TITLE|RAMK|BORDER|GRANIC|KRESK|HATCH|DEFPOINTS|"
    r"TRAS|GRAW|GRAV|NAPIS|MARK", re.I)


def _klastry(punkty, tol):
    """Grupuje punkty leżące bliżej niż tol (siatka + union-find). Zwraca listę list indeksów."""
    rodzic = list(range(len(punkty)))

    def korzen(i):
        while rodzic[i] != i:
            rodzic[i] = rodzic[rodzic[i]]
            i = rodzic[i]
        return i

    siatka = {}
    for i, (x, y) in enumerate(punkty):
        siatka.setdefault((math.floor(x / tol), math.floor(y / tol)), []).append(i)
    for (cx, cy), idx in siatka.items():
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for j in siatka.get((cx + dx, cy + dy), ()):
                    for i in idx:
                        if i < j and math.dist(punkty[i], punkty[j]) <= tol:
                            rodzic[korzen(i)] = korzen(j)
    grupy = {}
    for i in range(len(punkty)):
        grupy.setdefault(korzen(i), []).append(i)
    return list(grupy.values())


def _prymitywy_dxf(doc):
    """Elementy tnące z modelspace rozbite na LINE/ARC/CIRCLE/SPLINE/ELLIPSE (polilinie
    rozbijane na odcinki i łuki), każdy jako łamana pts (krok 0,01 mm na łukach)."""
    out = []

    def dodaj(e, warstwa):
        t = e.dxftype()
        if t in ("LWPOLYLINE", "POLYLINE"):
            try:
                for v in e.virtual_entities():
                    dodaj(v, warstwa)
            except Exception:
                pass
            return
        if t not in ("LINE", "ARC", "CIRCLE", "SPLINE", "ELLIPSE"):
            return
        try:
            pts = np.array([(v.x, v.y) for v in ezpath.make_path(e).flattening(0.01)])
        except Exception:
            return
        if len(pts) < 2:
            return
        dl = float(np.hypot(*np.diff(pts, axis=0).T).sum())
        if dl < STYK_TOL:
            return  # "linia" długości 0 — artefakt eksportu
        p = dict(typ=t, warstwa=warstwa, pts=pts, dl=dl, reczny=warstwa == "0",
                 zamk=t == "CIRCLE" or math.dist(pts[0], pts[-1]) < STYK_TOL)
        if t == "CIRCLE":
            p["srodek"] = (e.dxf.center.x, e.dxf.center.y)
            p["fi"] = 2 * e.dxf.radius
        out.append(p)

    for e in doc.modelspace():
        warstwa = e.dxf.get("layer", "0")
        if not WARSTWY_NIE_TNACE.search(warstwa):
            dodaj(e, warstwa)
    return out


def _sygnatura(p):
    """Ten sam element narysowany 2x daje tę samą sygnaturę (niezależnie od kierunku)."""
    if p["typ"] == "CIRCLE":
        return ("C", round(p["srodek"][0], 2), round(p["srodek"][1], 2), round(p["fi"], 2))
    konce = sorted((tuple(np.round(p["pts"][0], 2)), tuple(np.round(p["pts"][-1], 2))))
    return (p["typ"], tuple(konce), round(p["dl"], 1), tuple(np.round(p["pts"].mean(axis=0), 1)))


def _faza(S, sasiad_a, sasiad_b):
    """Odcinek S między dwoma odcinkami = faza narożnika, jeśli przedłużenia sąsiadów
    przecinają się za S (wirtualny narożnik V) pod kątem 45-135°. Zwraca (noga1, noga2, V)."""
    a, b = S["pts"][0], S["pts"][-1]
    (E1, f1), (E2, f2) = sasiad_a, sasiad_b
    u1 = (f1 - a) / np.linalg.norm(f1 - a)
    u2 = (f2 - b) / np.linalg.norm(f2 - b)
    if abs(float(np.dot(u1, u2))) > 0.7072:   # kąt narożnika poza 45-135°
        return None
    try:
        t1, t2 = np.linalg.solve(np.array([-u1, u2]).T, b - a)  # a - t1*u1 = b - t2*u2 = V
    except np.linalg.LinAlgError:
        return None
    if not (0.1 < t1 <= FAZA_MAX and 0.1 < t2 <= FAZA_MAX):
        return None
    if S["dl"] >= min(E1["dl"], E2["dl"]):
        return None
    nogi = sorted((round(float(t1), 1), round(float(t2), 1)))
    return nogi[0], nogi[1], tuple(a - t1 * u1)


def dxf_analiza(path):
    """Kontur DXF tak, jak zobaczy go wypalarka. Zwraca słownik z listami problemów
    (punkty w układzie współrzędnych DXF, żeby dało się je znaleźć w AutoCADzie)."""
    prym = _prymitywy_dxf(ezdxf.readfile(path))
    wynik = dict(duble=0, przerwy=[], wolne=[], rozgal=[], fazy=[], okregi=[],
                 wspolsrodkowe=[], reczne=0, wystaje=[])
    if not prym:
        return wynik

    # --- podwójne linie (zdublowane elementy usuwamy przed analizą konturu)
    widziane, unikalne = set(), []
    for p in prym:
        s = _sygnatura(p)
        if s in widziane:
            wynik["duble"] += 1
        else:
            widziane.add(s)
            unikalne.append(p)

    # --- węzły konturu: każdy koniec otwartego elementu powinien stykać się z dokładnie jednym innym
    otwarte = [p for p in unikalne if not p["zamk"]]
    konce = []  # konce[2*i] = początek otwarte[i], konce[2*i+1] = koniec
    for p in otwarte:
        konce.append(tuple(p["pts"][0]))
        konce.append(tuple(p["pts"][-1]))
    grupy = _klastry(konce, STYK_TOL)
    wezel = {}
    for g in grupy:
        for k in g:
            wezel[k] = g

    seg_a = np.vstack([p["pts"][:-1] for p in unikalne])
    seg_b = np.vstack([p["pts"][1:] for p in unikalne])
    seg_wl = np.concatenate([np.full(len(p["pts"]) - 1, i) for i, p in enumerate(unikalne)])
    idx_w_unikalnych = {id(p): i for i, p in enumerate(unikalne)}

    def na_innym_elemencie(pkt, wlasny):
        d = seg_b - seg_a
        dd = (d * d).sum(axis=1)
        dd[dd == 0] = 1e-12
        t = np.clip(((pkt - seg_a) * d).sum(axis=1) / dd, 0, 1)
        odl = np.hypot(*(seg_a + d * t[:, None] - pkt).T)
        return bool(((odl < STYK_TOL) & (seg_wl != wlasny)).any())

    wolne = []
    for g in grupy:
        pkt = np.mean([konce[k] for k in g], axis=0)
        if len(g) >= 3:
            wynik["rozgal"].append(tuple(pkt))
        elif len(g) == 1:
            if na_innym_elemencie(pkt, idx_w_unikalnych[id(otwarte[g[0] // 2])]):
                wynik["rozgal"].append(tuple(pkt))   # koniec linii na środku innej linii (T)
            else:
                wolne.append(tuple(pkt))
    # wolne końce blisko siebie = przerwa w konturze; reszta = linia urwana / wystająca
    uzyte = set()
    for i, p in enumerate(wolne):
        if i in uzyte:
            continue
        najbl = min(((math.dist(p, q), j) for j, q in enumerate(wolne) if j != i and j not in uzyte),
                    default=None)
        if najbl and najbl[0] <= PRZERWA_MAX:
            uzyte.update((i, najbl[1]))
            wynik["przerwy"].append((najbl[0], p))
        else:
            uzyte.add(i)
            wynik["wolne"].append(p)

    # --- fazy narożników: odcinek, którego oba końce łączą się z dokładnie jednym innym odcinkiem
    for i, S in enumerate(otwarte):
        if S["typ"] != "LINE" or S["dl"] > FAZA_MAX:
            continue
        sasiedzi = []
        for e in (0, 1):
            g = wezel[2 * i + e]
            if len(g) != 2:
                break
            inny = g[0] if g[1] == 2 * i + e else g[1]
            E = otwarte[inny // 2]
            if E["typ"] != "LINE" or E is S:
                break
            sasiedzi.append((E, E["pts"][-1] if inny % 2 == 0 else E["pts"][0]))
        else:
            f = _faza(S, sasiedzi[0], sasiedzi[1])
            if f:
                wynik["fazy"].append(f)

    # --- otwory i współśrodkowe okręgi
    okregi = [p for p in unikalne if p["typ"] == "CIRCLE"]
    wspol = set()
    for g in _klastry([p["srodek"] for p in okregi], STYK_TOL):
        fis = sorted({round(okregi[k]["fi"], 2) for k in g})
        if len(fis) >= 2:
            wynik["wspolsrodkowe"].append((fis, okregi[g[0]]["srodek"]))
            wspol.update(g)
    wynik["okregi"] = [(p["fi"], p["srodek"], p["reczny"], k in wspol) for k, p in enumerate(okregi)]

    # --- elementy dorysowane ręcznie (warstwa "0"), gdy reszta pochodzi z eksportu Inventora
    z_eksportu = [p for p in unikalne if not p["reczny"]]
    reczne = [p for p in unikalne if p["reczny"]]
    if z_eksportu and reczne:
        wynik["reczne"] = len(reczne)
        obrys = np.vstack([p["pts"] for p in z_eksportu])
        lo, hi = obrys.min(axis=0), obrys.max(axis=0)
        for p in reczne:
            wyst = float(max((lo - p["pts"].min(axis=0)).max(), (p["pts"].max(axis=0) - hi).max(), 0))
            if wyst > WYSTAJE_MAX:
                wynik["wystaje"].append((wyst, tuple(p["pts"][0])))
    return wynik


def _mm(x):
    return f"{x:.2f}".rstrip("0").rstrip(".").replace(".", ",")


def _gdzie(punkty, ile=3):
    s = ", ".join(f"({x:.1f}; {y:.1f})" for x, y in punkty[:ile])
    return s + (" …" if len(punkty) > ile else "")


def _lista_faz(fazy):
    licz = {}
    for n1, n2, _ in fazy:
        licz[(n1, n2)] = licz.get((n1, n2), 0) + 1
    return ", ".join(f"{_mm(a)}x{_mm(b)}" + (f" ({n} szt.)" if n > 1 else "") for (a, b), n in licz.items())


def ocen_kontur(b, kat, ana, tresc):
    """Zamienia wynik dxf_analiza (+ wymiary z rysunku PDF, jeśli jest) na listy
    (błędy, uwagi). Każdy komunikat zaczyna się od ".DXF:", żeby trafił do kolumny Plik."""
    bl, uw = [], []
    ukos = "u" in b["flags"]
    obrobka = "o" in b["flags"]
    gieta = "p" in b["flags"]
    przymiar = kat == "przygotowka" or re.search(r"PRZYMIAR|PRZYG", f"{b['title']} {b['desc']}", re.I)

    if ana["duble"] >= 3:
        bl.append(f".DXF: podwójne linie — {ana['duble']} elementów narysowanych 2x w tym samym miejscu "
                  f"(laser tnie dwa razy po tym samym śladzie)")
    elif ana["duble"]:
        uw.append(f".DXF: podwójne linie — {ana['duble']} zdublowany element (bywa z eksportu)")

    if ana["przerwy"]:
        d_max = max(d for d, _ in ana["przerwy"])
        bl.append(f".DXF: kontur — {len(ana['przerwy'])} przerw(a) w konturze (największa {_mm(d_max)} mm), "
                  f"przy {_gdzie([p for _, p in ana['przerwy']])}")
    if ana["wolne"]:
        bl.append(f".DXF: kontur — {len(ana['wolne'])} wolnych końców linii (kontur niedomknięty albo linia "
                  f"wystaje za narożnik), przy {_gdzie(ana['wolne'])}")
    if ana["rozgal"]:
        msg = (f".DXF: kontur — {len(ana['rozgal'])} rozgałęzień (linia kończy się na innej linii albo 3+ linie "
               f"w jednym punkcie), przy {_gdzie(ana['rozgal'])} — nieprzycięty narożnik po dorysowaniu fazy "
               f"albo linie fazy/ukosu krawędzi, które laser wytnie")
        if ukos or gieta:
            uw.append(msg + " (część " + ("z ukosowaniem" if ukos else "gięta")
                      + " — sprawdź, czy to linia " + ("ukosu" if ukos else "gięcia") + " na warstwie cięcia)")
        else:
            bl.append(msg)
    for fis, srodek in ana["wspolsrodkowe"]:
        msg = (f".DXF: otwory — współśrodkowe okręgi Ø{' / Ø'.join(_mm(f) for f in fis)} przy "
               f"{_gdzie([srodek])} — faza/pogłębienie otworu wyeksportowane do DXF, laser wytnie większy okrąg")
        (uw if obrobka else bl).append(msg)
    if ana["wystaje"] and not przymiar:
        w = max(ana["wystaje"])
        bl.append(f".DXF: dorysowane — element z warstwy 0 wystaje {w[0]:.0f} mm poza obrys części przy "
                  f"{_gdzie([w[1]])} (zbłąkana geometria, cięcie w powietrzu)")
    if ana["reczne"]:
        uw.append(f".DXF: dorysowane — {ana['reczne']} element(y) dorysowane ręcznie na warstwie 0 (do wglądu)")

    if not tresc:
        return bl, uw

    # --- otwory z DXF vs Ø na rysunku
    fi_rys = tresc.get("fi", []) + [2 * r for r in tresc.get("promienie", [])]

    def zwymiarowany(fi):
        return (any(abs(fi - f) <= OTWOR_TOL for f in fi_rys)
                or any(0.75 * d <= fi <= d + 0.5 for d in tresc.get("gwinty", [])))

    if not obrobka and not przymiar:
        reczne_bez = sorted({round(fi, 2) for fi, _, reczny, wsp in ana["okregi"]
                             if reczny and not wsp and not zwymiarowany(fi)})
        eksp_bez = sorted({round(fi, 2) for fi, _, reczny, wsp in ana["okregi"]
                           if not reczny and not wsp and not zwymiarowany(fi)})
        if reczne_bez:
            bl.append(f".DXF: otwory — dorysowany otwór Ø{', Ø'.join(_mm(f) for f in reczne_bez)} (warstwa 0) "
                      f"bez wymiaru Ø na rysunku PDF — produkcja nie wie o otworze")
        if eksp_bez and tresc.get("fi"):   # stare rysunki bez wymiarów Ø — nie sprawdzamy
            uw.append(f".DXF: otwory — otwór Ø{', Ø'.join(_mm(f) for f in eksp_bez)} z DXF nie występuje "
                      f"w wymiarach rysunku (Ø na rysunku: {', '.join(_mm(f) for f in sorted(set(tresc['fi'])))})")

    # --- fazy narożników z DXF vs notki faz na rysunku
    fazy_dxf = ana["fazy"]
    lagodnie = ukos or obrobka

    def pasuje(n, m):
        return abs(n[0] - m[0]) <= FAZA_TOL and abs(n[1] - m[1]) <= FAZA_TOL

    rys = {}
    for f in tresc.get("fazy", []):
        if max(f["nogi"]) <= FAZA_MIN_RYS:
            continue  # łamanie krawędzi (np. 1x45° na otworze) — nie ma go w konturze
        r = rys.setdefault(f["nogi"], dict(teksty=[], ile=0, z_iloscia=False))
        r["teksty"].append(f["tekst"])
        r["ile"] += f["ile"] or 1
        r["z_iloscia"] |= f["ile"] is not None
    dopisek = " (część z ukosowaniem/obróbką — faza może być robiona później)" if lagodnie else ""
    if rys:
        for nogi, r in rys.items():
            n = sum(1 for f in fazy_dxf if pasuje(f, nogi))
            if n == 0:
                (uw if lagodnie else bl).append(
                    f".DXF: fazy — na rysunku faza {r['teksty'][0]}, w DXF brak takiej fazy ("
                    + (f"fazy w DXF: {_lista_faz(fazy_dxf)}" if fazy_dxf else "w DXF nie ma żadnej fazy narożnika")
                    + ")" + dopisek)
            elif r["z_iloscia"] and n < r["ile"]:
                (uw if lagodnie else bl).append(
                    f".DXF: fazy — na rysunku {r['ile']}x faza {_mm(nogi[0])}x{_mm(nogi[1])}, w DXF tylko {n}" + dopisek)
        obce = [f for f in fazy_dxf if not any(pasuje(f, nogi) for nogi in rys)]
        if obce:
            uw.append(f".DXF: fazy — w DXF faza {_lista_faz(obce)} przy {_gdzie([f[2] for f in obce])}, "
                      f"której nie ma na rysunku (rysunek: {', '.join(r['teksty'][0] for r in rys.values())})")
    elif fazy_dxf and tresc.get("liczby"):
        nie = [f for f in fazy_dxf
               if not any(abs(f[0] - x) <= FAZA_TOL or abs(f[1] - x) <= FAZA_TOL for x in tresc["liczby"])]
        if nie:
            uw.append(f".DXF: fazy — faza {_lista_faz(nie)} z DXF nie jest zwymiarowana na rysunku PDF, "
                      f"przy {_gdzie([f[2] for f in nie])}")
    return bl, uw


# ---------------------------------------------------------------- geometria STP (3D)
_INV_APP = None
_INV_STARTED_BY_US = False
_STP_OSTRZEZONO = False  # ostrzeżenie o braku Inventora pokazujemy raz na całe uruchomienie


def get_inventor():
    """Zwraca instancję Inventor.Application przez COM — używa już uruchomionej (jeśli jest),
    inaczej odpala nową w tle (niewidoczną). Nie zamyka cudzej, już otwartej instancji na końcu."""
    global _INV_APP, _INV_STARTED_BY_US
    if _INV_APP is not None:
        return _INV_APP
    try:
        _INV_APP = win32com.client.GetActiveObject("Inventor.Application")  # dolacz do uruchomionej
    except Exception:
        _INV_APP = win32com.client.Dispatch("Inventor.Application")  # odpal nowa (w tle)
        try:
            _INV_APP.Visible = False
        except Exception:
            pass
        _INV_STARTED_BY_US = True
    return _INV_APP


@atexit.register
def _zamknij_inventor():
    if _INV_STARTED_BY_US and _INV_APP is not None:
        try:
            _INV_APP.Quit()
        except Exception:
            pass


def znajdz_stp(folder):
    """Szuka pliku STP dla projektu: najpierw '<nazwa_folderu_bez_dopiskow>.stp/.step'
    (wzorzec z archiwum: GE6725\\GE6725.stp), inaczej dowolny *.stp/*.step w folderze."""
    baza = re.sub(r"[\s\(\)].*$", "", os.path.basename(folder.rstrip("\\/")))
    for ext in (".stp", ".step", ".STP", ".STEP"):
        kand = os.path.join(folder, baza + ext)
        if os.path.isfile(kand):
            return kand
    try:
        for f in os.listdir(folder):
            if f.lower().endswith((".stp", ".step")):
                return os.path.join(folder, f)
    except OSError:
        pass
    return None


def stp_dims_mapa(stp_path):
    """Otwiera plik STP przez COM Inventora, zwraca (mapa, blad) gdzie mapa to
    {NAZWA_CZESCI_UPPER: (dx, dy, dz) w mm} z LOKALNEGO (nieobróconego) RangeBox definicji
    każdej liściastej części — odpornego na obrót/pozycję części w zespole.
    Jednostki bazowe Inventora to zawsze centymetry, stąd mnożenie *10.0 -> mm (nie wołamy
    UnitsOfMeasure.ConvertUnits przez surowe COM, żeby nie zgadywać numeru wyliczenia)."""
    global _STP_OSTRZEZONO
    if not STP_MODULE_OK:
        return {}, "brak modułu pywin32 (pip install pywin32) — kontrola STP wyłączona"
    try:
        oApp = get_inventor()
    except Exception as e:
        if not _STP_OSTRZEZONO:
            _STP_OSTRZEZONO = True
        return {}, f"nie udało się uruchomić/połączyć z Inventorem ({type(e).__name__}: {e})"

    mapa = {}
    oDoc = None
    try:
        oDoc = oApp.Documents.Open(stp_path, False)
    except TypeError:
        try:
            oDoc = oApp.Documents.Open(stp_path)
        except Exception as e:
            return {}, f"nie udało się otworzyć pliku STP ({type(e).__name__}: {e})"
    except Exception as e:
        return {}, f"nie udało się otworzyć pliku STP ({type(e).__name__}: {e})"

    try:
        try:
            # oDoc.ComponentDefinition rzuca AttributeError na "gołym" obiekcie Document
            # zwróconym przez Documents.Open (early-bound gen_py widzi tylko bazowy interfejs
            # Document, nie AssemblyDocument) - CastTo jest oficjalnym lekarstwem pywin32 na
            # dokładnie ten przypadek. Potwierdzone empirycznie 2026-07-18 na GE75206-S1.stp
            # (64 wystąpienia, iProperty "Part Number" poprawnie = numer części z BOM).
            oAsmDoc = win32com.client.CastTo(oDoc, "AssemblyDocument")
            oCompDef = oAsmDoc.ComponentDefinition
            oLeafOcc = oCompDef.Occurrences.AllLeafOccurrences  # wymusza blad tu, nie w petli, jesli to nie zespol
        except Exception:
            return {}, ("plik STP zaimportowany jako pojedyncza część (nie zespół) — Inventor "
                        "nie rozbił go na komponenty; sprawdzenie geometrii per-część pominięte "
                        "(sam plik nie jest błędem)")
        for oOcc in oLeafOcc:
            nazwa = None
            try:
                nazwa = oOcc.Definition.Document.PropertySets.Item(
                    "Design Tracking Properties").Item("Part Number").Value
            except Exception:
                pass
            if not nazwa:
                try:
                    nazwa = oOcc.Definition.Document.DisplayName
                except Exception:
                    pass
            if not nazwa:
                try:
                    nazwa = oOcc.Name.split(":")[0]
                except Exception:
                    continue
            try:
                box = oOcc.Definition.RangeBox
                dx = abs(box.MaxPoint.X - box.MinPoint.X) * 10.0
                dy = abs(box.MaxPoint.Y - box.MinPoint.Y) * 10.0
                dz = abs(box.MaxPoint.Z - box.MinPoint.Z) * 10.0
            except Exception:
                continue
            mapa[str(nazwa).strip().upper()] = (dx, dy, dz)
    except Exception as e:
        return mapa, f"błąd podczas czytania zespołu STP ({type(e).__name__}: {e})"
    finally:
        try:
            if oDoc:
                oDoc.Close(True)
        except Exception:
            pass
    return mapa, None


# ---------------------------------------------------------------- pomocnicze do raportu xlsx
def etykieta_split(etykieta):
    """'[7.3.2] PG4104413 (WZM BOKU)' -> ('7.3.2', 'PG4104413', 'WZM BOKU')."""
    m = re.match(r"^\[(.*?)\]\s+(\S+)\s+\((.*)\)\s*$", etykieta)
    if m:
        return m.group(1), m.group(2), m.group(3)
    return "", etykieta, ""


def plik_i_reszta(opis):
    """'.STP: grubosc 12 vs BOM (...)' -> ('STP', 'grubosc 12 vs BOM (...)')."""
    m = re.match(r"^\.(DXF|PDF|STP)\s*:\s*(.*)$", opis, re.I)
    if m:
        return m.group(1).upper(), m.group(2)
    return "", opis


RODZAJ_WZORCE = [
    (re.compile(r"^kontur —"), "kontur DXF (przerwy/rozgałęzienia)"),
    (re.compile(r"^fazy —"), "fazy DXF vs rysunek"),
    (re.compile(r"^otwory —"), "otwory DXF vs rysunek"),
    (re.compile(r"^podwójne linie —"), "podwójne linie DXF"),
    (re.compile(r"^dorysowane —"), "dorysowane ręcznie (warstwa 0)"),
    (re.compile(r"^BRAK (DXF|PDF)"), "plik brakujący"),
    (re.compile(r"geometria DXF|DXF zawiera więcej"), "geometria DXF"),
    (re.compile(r"bryła", re.I), "geometria STP (3D)"),  # prefiks ".STP:" juz wyciety do kolumny Plik
    (re.compile(r"grubo(ś|s)ć"), "grubość vs BOM"),
    (re.compile(r"materia(ł|l)"), "materiał vs BOM"),
    (re.compile(r"ilo(ś|s)ć|stempel"), "ilość/stempel"),
    (re.compile(r"flagi"), "flagi MOPUC"),
    (re.compile(r"tabliczka"), "tabliczka PDF"),
    (re.compile(r"prefiks"), "prefiks nazwy pliku"),
    (re.compile(r"bez pozycji w BOM"), "plik bez pozycji w BOM"),
]


def rodzaj_z_opisu(opis):
    for pat, nazwa in RODZAJ_WZORCE:
        if pat.search(opis):
            return nazwa
    return "inne"


def pos_key(pos):
    """Sortowanie POS naturalnie/numerycznie: '7.3.2' < '10' (nie leksykograficznie)."""
    klucz = []
    for p in re.split(r"[.\-]", pos or ""):
        try:
            klucz.append((0, float(p.replace(",", "."))))
        except ValueError:
            klucz.append((1, p))
    return klucz


def dodaj_wiersze(folder, lista, poziom):
    for etykieta_pelna in lista:
        if ": " in etykieta_pelna and etykieta_pelna.count("[") >= 1:
            etykieta, opis = etykieta_pelna.split(": ", 1)
        else:
            etykieta, opis = "", etykieta_pelna
        pos, part, tytul = etykieta_split(etykieta) if etykieta else ("", etykieta_pelna.split(":")[0], "")
        plik, reszta = plik_i_reszta(opis)
        rodzaj = "model 3D .stp (informacyjnie)" if poziom == "INFO STP" else rodzaj_z_opisu(reszta)
        WSZYSTKIE_WIERSZE.append(dict(
            folder=os.path.basename(folder.rstrip("\\/")),
            pos=pos, part=part, tytul=tytul, poziom=poziom,
            plik=plik, rodzaj=rodzaj, opis=reszta,
        ))


# ---------------------------------------------------------------- kontrola
def sprawdz(folder):
    folder = folder.rstrip("\\/")
    print("\n" + "=" * 92)
    print(f"KONTROLA (GEO+STP): {folder}")
    if not GEO:
        print("  UWAGA: brak ezdxf (pip install ezdxf) — kontrola geometrii DXF wyłączona")
    if not STP_MODULE_OK:
        print("  UWAGA: brak pywin32 (pip install pywin32) — kontrola geometrii STP wyłączona")
    try:
        top = os.listdir(folder)
    except OSError as e:
        print("  BŁĄD dostępu:", e)
        return 1
    xlsx = [f for f in top if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$")]
    if not xlsx:
        print("  BŁĄD: brak pliku BOM (.xlsx) w folderze.")
        return 1
    if len(xlsx) > 1:
        print(f"  UWAGA: kilka xlsx, biorę {xlsx[0]}")
    bom, err = czytaj_bom(os.path.join(folder, xlsx[0]))
    if err:
        print("  BŁĄD BOM:", err)
        return 1

    bom_parts = {b["part"].upper() for b in bom}
    bom_asm = {b["part"].upper() for b in bom if "m" in b["flags"]}
    proj_root = re.sub(r"[\s\(\)].*$", "", os.path.basename(folder).upper())
    bom_asm.add(proj_root)

    pliki = {}
    pominiete = []
    for f in sorted(top):
        ext = os.path.splitext(f)[1].upper()
        if ext not in (".DXF", ".PDF"):
            continue
        if JUNK.search(f):
            pominiete.append(f)
            continue
        info = parsuj_nazwe(f, bom_parts, bom_asm)
        if info:
            info["file"] = f
            pliki.setdefault(info["part"].upper(), {})[ext] = info

    # --- geometria STP (raz na cały folder, nie per-część)
    stp_mapa, stp_blad = {}, None
    sciezka_stp = znajdz_stp(folder)
    uwagi_wstepne = []
    if sciezka_stp:
        print(f"  Otwieranie STP przez Inventor (może to potrwać kilkadziesiąt sekund): "
              f"{os.path.basename(sciezka_stp)}")
        stp_mapa, stp_blad = stp_dims_mapa(sciezka_stp)
        if stp_blad:
            uwagi_wstepne.append(f"STP ({os.path.basename(sciezka_stp)}): {stp_blad}")
        elif stp_mapa:
            print(f"  STP: odczytano geometrię {len(stp_mapa)} części.")
    else:
        uwagi_wstepne.append("brak pliku .stp/.step w folderze — kontrola geometrii 3D pominięta "
                              "(sam brak pliku nie jest błędem, nie wszystkie projekty go mają)")

    bledy, uwagi, bez_stempla = [], [], []
    stp_info = []  # różnice z modelu .stp — tylko informacyjnie, nie są błędami
    uwagi.extend(uwagi_wstepne)
    ok = 0
    dxf_sprawdzone, dxf_bledne, dxf_uwagi = 0, [], 0

    for b in bom:
        uwagi_przed = len(uwagi)
        kat = kategoria(b, proj_root)
        f = pliki.pop(b["part"].upper(), {})
        dxf, pdf = f.get(".DXF"), f.get(".PDF")
        problemy = []
        etykieta = f"[{b['pos']}] {b['part']} ({b['title'] or b['desc'][:25] or kat})"

        # --- wymagane pliki wg kategorii
        if kat == "purchased":
            if not dxf and not pdf:
                continue  # norma (93% w archiwum) — bez szumu
        elif kat == "ramka":
            if not pdf:
                uwagi.append(f"{etykieta}: ramka tabliczki bez PDF — w archiwum 61% ma PDF, 20% nie; do decyzji")
                continue
        elif kat == "zlozenie":
            if not pdf:
                problemy.append("BRAK PDF złożenia")
            if dxf:
                uwagi.append(f"{etykieta}: złożenie ma DXF — rzadkie (szablony G BOLTON), sprawdź czy celowe")
        elif kat == "blacha":
            if not dxf:
                if czy_krawedz_plaskownik(b):
                    uwagi.append(f"{etykieta}: krawędź z paska {b['desc']} bez DXF — typowe dla GRADING "
                                 f"(płaskownik, cięcie+ukos), ale flaga C w BOM by to potwierdziła (jest: "
                                 f"'{''.join(sorted(b['flags'])) or '-'}')")
                else:
                    problemy.append("BRAK DXF (blacha wypalana — 99% archiwum ma DXF)")
            if not pdf:
                problemy.append("BRAK PDF")
        elif kat == "blacha_c":
            if not pdf:
                problemy.append("BRAK PDF")
        elif kat == "przygotowka":
            if not dxf:
                problemy.append("BRAK DXF (przygotówka/PL-t)")
        else:  # material_ciety, inne
            if not pdf:
                if kat == "inne" and not b["desc"] and not dxf:
                    uwagi.append(f"{etykieta}: brak plików, puste DESCRIPTION — wygląda na część kupną "
                                 f"oznaczoną Normal; sprawdź BOM STRUCTURE")
                    continue
                problemy.append("BRAK PDF")

        # --- nazwa pliku vs BOM
        for info in (dxf, pdf):
            if not info:
                continue
            e = info["ext"]
            if info["asm"]:
                if info["flags"] != b["flags"]:
                    problemy.append(f"{e}: flagi '{''.join(sorted(info['flags'])) or '-'}' vs BOM "
                                    f"'{''.join(sorted(b['flags'])) or '-'}'")
                continue
            if info["thk"] is not None and b["dims"]:
                if info["thk"] not in b["dims"]:
                    problemy.append(f"{e}: grubość {info['thk']:g} nie występuje w wymiarach BOM ({b['desc']})")
            if kat == "blacha" and info["thk"] is None:
                problemy.append(f"{e}: brak grubości w nazwie ({b['desc']})")
            if info["mat"] and not mat_zgodny(info["mat"], b["mat"]):
                problemy.append(f"{e}: materiał '{info['mat']}' vs BOM '{b['mat']}'")
            if info["qty"] is not None and b["qty"] is not None:
                if info["qty"] not in (b["qty"], b["qty_total"]):
                    problemy.append(f"{e}: ilość {info['qty']} vs BOM {b['qty']}"
                                    + (f" (całk. {b['qty_total']:g})" if b["qty_total"] != b["qty"] else ""))
            if info["flags"] != b["flags"]:
                problemy.append(f"{e}: flagi '{''.join(sorted(info['flags'])) or '-'}' vs BOM "
                                f"'{''.join(sorted(b['flags'])) or '-'}'")

        # --- prefiksy
        if dxf and not dxf["prefiks"]:
            uwagi.append(f"{etykieta}: DXF bez prefiksu złożenia w nazwie")
        if pdf and pdf["prefiks"] and not pdf["asm"]:
            uwagi.append(f"{etykieta}: PDF z prefiksem złożenia (wg wzorca PDF-y części są bez)")

        # --- zawartosc PDF
        tresc = None
        if pdf:
            tresc = czytaj_pdf(os.path.join(folder, pdf["file"]))
            if "stamp" not in tresc:
                bez_stempla.append(etykieta)
            elif b["qty"] is not None and tresc["stamp"] not in (b["qty"], b["qty_total"]):
                if tresc.get("lustro"):
                    uwagi.append(f"{etykieta}: stempel 'Ilość: {tresc['stamp']}' vs BOM {b['qty']}, ale jest "
                                 f"adnotacja o odbiciu lustrzanym — prawdopodobnie celowe")
                else:
                    problemy.append(f"stempel 'Ilość: {tresc['stamp']}' vs BOM {b['qty']}")
            if kat in ("blacha", "blacha_c") and "dims" in tresc and b["dims"]:
                if tresc["dims"] != sorted(b["dims"]):
                    problemy.append(f"tabliczka PDF: wymiary {tresc['dims']} vs BOM {sorted(b['dims'])}")
            if kat in ("blacha", "blacha_c") and "mat" in tresc and not mat_zgodny(tresc["mat"], b["mat"]):
                problemy.append(f"tabliczka PDF: materiał '{tresc['mat']}' vs BOM '{b['mat']}'")

        # --- geometria DXF vs DESCRIPTION (tylko wersja GEO)
        if GEO and dxf and not dxf["asm"] and len(b["dims"]) == 3:
            bb = None
            try:
                bb = dxf_gabaryt(os.path.join(folder, dxf["file"]))
            except Exception as e:
                uwagi.append(f"{etykieta}: nie udało się odczytać geometrii DXF ({type(e).__name__})")
            if bb:
                dev, plan, bbs = geo_porownaj(b, dxf, bb)
                opis_geo = (f"geometria DXF {bbs[0]:.0f}x{bbs[1]:.0f} vs BOM {plan[0]:g}x{plan[1]:g} "
                            f"(odchyłka {dev:.1f} mm)")
                gieta = "p" in b["flags"]
                obrabiana = "o" in b["flags"] or "u" in b["flags"]
                if max(plan) > 0 and max(bbs) > 1.5 * max(plan):
                    uwagi.append(f"{etykieta}: DXF zawiera więcej niż kontur części "
                                 f"(ramka/arkusz? {opis_geo})")
                elif gieta:
                    if dev > GEO_TOL_BLAD:
                        uwagi.append(f"{etykieta}: {opis_geo} — część gięta/walcowana; jeśli DESCRIPTION "
                                     f"to wymiar po gięciu, różnica jest naturalna")
                elif obrabiana:
                    if dev > GEO_TOL_BLAD:
                        uwagi.append(f"{etykieta}: {opis_geo} — część z obróbką/ukosem (możliwy naddatek); "
                                     f"zweryfikuj")
                else:
                    if dev > GEO_TOL_BLAD:
                        problemy.append(opis_geo)
                    elif dev > GEO_TOL_INFO:
                        uwagi.append(f"{etykieta}: {opis_geo} — strefa szara")

        # --- gabaryt DXF vs tabliczka PDF, gdy BOM nie ma 3 wymiarów (np. "PL 10" przygotówki)
        if GEO and dxf and not dxf["asm"] and len(b["dims"]) != 3 and tresc and "dims" in tresc:
            try:
                bb = dxf_gabaryt(os.path.join(folder, dxf["file"]))
            except Exception:
                bb = None
            if bb:
                dev, plan, bbs = geo_porownaj(dict(b, dims=tresc["dims"]), dxf, bb)
                opis_geo = (f"geometria DXF {bbs[0]:.0f}x{bbs[1]:.0f} vs rysunek PDF {plan[0]:g}x{plan[1]:g} "
                            f"(odchyłka {dev:.1f} mm)")
                if dev > GEO_TOL_BLAD and not ("p" in b["flags"] or "o" in b["flags"] or "u" in b["flags"]):
                    problemy.append(opis_geo)
                elif dev > GEO_TOL_INFO:
                    uwagi.append(f"{etykieta}: {opis_geo} — strefa szara / część gięta lub obrabiana")

        # --- kontur DXF tak, jak zobaczy go wypalarka + fazy/otwory vs rysunek PDF
        if GEO and dxf and not dxf["asm"]:
            ana = None
            try:
                ana = dxf_analiza(os.path.join(folder, dxf["file"]))
            except Exception as e:
                uwagi.append(f"{etykieta}: nie udało się przeanalizować konturu DXF ({type(e).__name__})")
            if ana:
                bl_k, uw_k = ocen_kontur(b, kat, ana, tresc)
                problemy.extend(bl_k)
                uwagi.extend(f"{etykieta}: {u}" for u in uw_k)

        # --- geometria STP (bryła 3D) vs DESCRIPTION — niezależne od DXF, tylko jeśli mamy mapę.
        # Tylko INFORMACYJNIE (poziom "INFO STP", nie liczy się do błędów): gabaryt bryły z STP
        # zależy od ułożenia części w złożeniu (obrót, pochylenie, gięcie), więc różnica przy
        # zgodnym DXF/PDF nie oznacza błędu dokumentacji (GE98022-S2, 2026-09-30).
        if stp_mapa and not (dxf and dxf.get("asm")) and len(b["dims"]) == 3:
            dims3 = stp_mapa.get(b["part"].upper())
            if dims3:
                wynik = geo_porownaj_3d(b["dims"], dims3)
                if wynik:
                    dev, plan, znaleziony = wynik
                    opis_geo = (f".STP: bryła {znaleziony[0]:.0f}x{znaleziony[1]:.0f}x{znaleziony[2]:.0f} "
                                f"vs BOM {plan[0]:g}x{plan[1]:g}x{plan[2]:g} (odchyłka {dev:.1f} mm)")
                    gieta = "p" in b["flags"]
                    obrabiana = "o" in b["flags"] or "u" in b["flags"]
                    if dev > GEO_TOL_BLAD:
                        if gieta:
                            powod = "część gięta — bryła w STP jest już po gięciu"
                        elif obrabiana:
                            powod = "część z obróbką — możliwy naddatek"
                        else:
                            powod = "najczęściej część obrócona/pochylona w złożeniu"
                        stp_info.append(f"{etykieta}: {opis_geo} — {powod}")
            elif b["part"].upper() != proj_root and "m" not in b["flags"] and kat not in ("purchased", "ramka"):
                stp_info.append(f"{etykieta}: .STP: część nie znaleziona w bryle zespołu STP po nazwie "
                                f"(inna nazwa w modelu albo część spoza tego pliku STP)")

        # --- zgodność DXF tej pozycji (do podsumowania na końcu raportu)
        if dxf and not dxf["asm"]:
            dxf_sprawdzone += 1
        if any("DXF" in p for p in problemy):
            dxf_bledne.append(b["part"])
        dxf_uwagi += sum(1 for u in uwagi[uwagi_przed:] if "DXF" in u)

        if problemy:
            bledy.append(f"{etykieta}: " + "; ".join(problemy))
        else:
            ok += 1

    if bez_stempla:
        if len(bez_stempla) > 5:
            uwagi.append(f"{len(bez_stempla)} PDF-ów bez stempla 'Ilość:' (folder nie przeszedł przez "
                         f"PdfRenamer / starsze wydanie?), np. {', '.join(bez_stempla[:3])}")
        else:
            for e in bez_stempla:
                uwagi.append(f"{e}: PDF bez stempla 'Ilość:' (nie przeszedł przez PdfRenamer?)")

    for key, f in pliki.items():
        if key == proj_root or key in bom_asm:
            continue  # glowne zlozenie bez wlasnego wiersza w BOM — norma w starszych wydaniach
        uwagi.append(f"{key}: plik(i) {', '.join(v['file'] for v in f.values())} bez pozycji w BOM")
        if ".DXF" in uwagi[-1].upper():
            dxf_uwagi += 1

    print(f"  BOM: {len(bom)} pozycji | zgodnych: {ok} | błędów: {len(bledy)} | uwag: {len(uwagi)}"
          + (f" | info STP: {len(stp_info)}" if stp_info else "")
          + (f" | pominięte wytyczne: {len(pominiete)}" if pominiete else ""))
    if bledy:
        print("  --- BŁĘDY " + "-" * 62)
        for x in bledy:
            print("   !!", x)
    if uwagi:
        print("  --- UWAGI " + "-" * 62)
        for x in uwagi:
            print("    -", x)
    if stp_info:
        print("  --- INFO STP (porównanie z modelem 3D .stp — informacyjnie, to nie są błędy) " + "-" * 12)
        for x in stp_info:
            print("    i", x)
    if dxf_sprawdzone and not dxf_bledne:
        print(f"  >>> WSZYSTKIE PLIKI DXF ZGODNE Z BOM ({dxf_sprawdzone}) <<<")
    if not bledy:
        print("  >>> FOLDER ZGODNY Z WZORCEM <<<")

    dodaj_wiersze(folder, bledy, "BŁĄD")
    dodaj_wiersze(folder, uwagi, "UWAGA")
    dodaj_wiersze(folder, stp_info, "INFO STP")
    PODSUMOWANIE_FOLDEROW.append(dict(
        folder=os.path.basename(folder), sciezka=folder,
        ok=ok, bledy=len(bledy), uwagi=len(uwagi), pozycji=len(bom),
        stp="TAK" if stp_mapa else ("BŁĄD" if stp_blad else "brak pliku"),
        stp_info=len(stp_info),
        dxf_sprawdzone=dxf_sprawdzone, dxf_bledne=dxf_bledne, dxf_uwagi=dxf_uwagi,
    ))
    return 1 if bledy else 0


# ---------------------------------------------------------------- raport xlsx
def zapisz_raport_xlsx(sciezka):
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "Błędy i uwagi"
    naglowki = ["Folder", "POS (miejsce w BOM)", "Część", "Tytuł/opis", "Poziom",
                "Plik", "Rodzaj sprawdzenia", "Opis problemu"]
    ws.append(naglowki)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="404040")
        c.alignment = Alignment(vertical="center")
    ws.freeze_panes = "A2"

    czerwony = PatternFill("solid", fgColor="F8CBCB")
    zolty = PatternFill("solid", fgColor="FFF3B0")
    niebieski = PatternFill("solid", fgColor="DDEBF7")
    zielony = PatternFill("solid", fgColor="C6EFCE")
    szary = PatternFill("solid", fgColor="EDEDED")
    kolor_poziomu = {"BŁĄD": czerwony, "UWAGA": zolty, "INFO STP": niebieski}
    kolejnosc_poziomu = {"BŁĄD": 0, "UWAGA": 1, "INFO STP": 2}

    wiersze = sorted(WSZYSTKIE_WIERSZE,
                      key=lambda w: (w["folder"], kolejnosc_poziomu.get(w["poziom"], 1), pos_key(w["pos"])))
    for w in wiersze:
        ws.append([w["folder"], w["pos"], w["part"], w["tytul"], w["poziom"],
                   w["plik"], w["rodzaj"], w["opis"]])
        fill = kolor_poziomu.get(w["poziom"], zolty)
        for c in ws[ws.max_row]:
            c.fill = fill
            c.alignment = Alignment(vertical="top", wrap_text=(c.column_letter == "H"))

    szerokosci = [24, 14, 16, 26, 10, 6, 20, 70]
    for i, sz in enumerate(szerokosci, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = sz
    if ws.max_row > 1:
        ws.auto_filter.ref = f"A1:H{ws.max_row}"

    # --- podsumowanie na końcu arkusza (poza autofiltrem): DXF, model .stp, werdykt
    ws.append([])
    ws.append(["PODSUMOWANIE"])
    ws[ws.max_row][0].font = Font(bold=True)

    def wiersz_podsumowania(folder, etykieta, tekst, fill):
        ws.append([folder, etykieta, tekst])
        r = ws.max_row
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
        for c in ws[r]:
            c.fill = fill
            c.alignment = Alignment(vertical="top", wrap_text=True)
        ws[r][1].font = Font(bold=True)
        ws.row_dimensions[r].height = 30

    for p in PODSUMOWANIE_FOLDEROW:
        n, zle, uw = p["dxf_sprawdzone"], p["dxf_bledne"], p["dxf_uwagi"]
        if zle:
            wiersz_podsumowania(p["folder"], "DXF", f"DXF NIEZGODNE Z BOM: {len(zle)} poz. "
                                f"({', '.join(zle)}) — czerwone wiersze wyżej.", czerwony)
        elif n:
            wiersz_podsumowania(p["folder"], "DXF", f"Wszystkie pliki DXF zgodne z BOM ({n} szt.: grubość, "
                                f"materiał, ilość, flagi, geometria)."
                                + (f" Uwagi do DXF wyżej ({uw}) są do przejrzenia, to nie są błędy." if uw else ""),
                                zielony)
        else:
            wiersz_podsumowania(p["folder"], "DXF", "Brak plików DXF do sprawdzenia.", szary)

        if p["stp_info"]:
            wiersz_podsumowania(p["folder"], "Model .stp",
                                f"Niebieskie wiersze INFO STP ({p['stp_info']}) to porównanie z modelem 3D (.stp). "
                                f"Gabaryt bryły w .stp zależy od ułożenia części w złożeniu (obrót, pochylenie, "
                                f"gięcie), więc taka różnica NIE oznacza błędu dokumentacji, jeśli DXF i PDF są "
                                f"zgodne z BOM.", niebieski)
        elif p["stp"] == "TAK":
            wiersz_podsumowania(p["folder"], "Model .stp", "Model 3D (.stp) zgodny z BOM.", zielony)

        if p["bledy"]:
            wiersz_podsumowania(p["folder"], "Wynik", f"Błędów w dokumentacji: {p['bledy']} — czerwone wiersze wyżej.",
                                czerwony)
        else:
            wiersz_podsumowania(p["folder"], "Wynik", "Dokumentacja zgodna z BOM — brak błędów."
                                + (f" Uwagi ({p['uwagi']}) są do przejrzenia." if p["uwagi"] else ""), zielony)

    ws2 = wb.create_sheet("Podsumowanie folderów")
    ws2.append(["Folder", "Pozycji BOM", "Zgodnych", "Błędów", "Uwag", "DXF", "Plik STP", "Info STP"])
    for c in ws2[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="404040")
    for p in PODSUMOWANIE_FOLDEROW:
        if p["dxf_bledne"]:
            dxf_stan = f"NIEZGODNE ({len(p['dxf_bledne'])})"
        elif p["dxf_sprawdzone"]:
            dxf_stan = f"ZGODNE ({p['dxf_sprawdzone']})"
        else:
            dxf_stan = "brak"
        ws2.append([p["folder"], p["pozycji"], p["ok"], p["bledy"], p["uwagi"], dxf_stan, p["stp"], p["stp_info"]])
        if p["bledy"] > 0:
            for c in ws2[ws2.max_row]:
                c.fill = czerwony
        else:
            ws2[ws2.max_row][5].fill = zielony if p["dxf_sprawdzone"] else szary
        if p["stp_info"]:
            ws2[ws2.max_row][7].fill = niebieski
    for i, sz in enumerate([28, 12, 10, 8, 8, 14, 10, 9], start=1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = sz
    if ws2.max_row > 1:
        ws2.auto_filter.ref = f"A1:H{ws2.max_row}"

    wb.save(sciezka)
    return sciezka


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)
    kod = 0
    folder_dla_raportu = None
    for arg in sys.argv[1:]:
        if os.path.isdir(arg):
            if folder_dla_raportu is None:
                folder_dla_raportu = os.path.dirname(arg.rstrip("\\/")) or arg
            kod |= sprawdz(arg)
        else:
            print("Pominięto (to nie folder):", arg)

    if WSZYSTKIE_WIERSZE or PODSUMOWANIE_FOLDEROW:
        znacznik = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        sciezka_raport = os.path.join(folder_dla_raportu or ".", f"RAPORT_SPRAWDZENIA_{znacznik}.xlsx")
        try:
            zapisz_raport_xlsx(sciezka_raport)
            print(f"\nZapisano raport zbiorczy: {sciezka_raport}")
            if not os.environ.get("WYDAJ_FOLDER_PIPELINE"):
                try:
                    os.startfile(sciezka_raport)
                except Exception:
                    pass
        except Exception as e:
            print(f"\nNie udało się zapisać raportu xlsx: {e}")

    sys.exit(kod)
