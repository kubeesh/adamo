# Logo FB-JELCZ – pliki CAD

| Plik | Do czego |
|---|---|
| `fb-jelcz.step` | Bryła 3D (199,7 × 36 × 3 mm), dwie bryły: niebieska `FB-` i czerwona `JELCZ`. Otwierasz w Inventorze i zapisujesz jako `.ipt`. |
| `fb-jelcz.dxf` | Kontury 2D (warstwy `BLUE` i `RED`) do wstawienia w szkic Inventora. |
| `fb-jelcz.svg` | Podgląd wektora. |
| `make_logo.py` | Skrypt, który generuje te pliki z `fb-jelcz-source.png`. |

## Inventor – wariant 1 (STEP → IPT)
1. **Plik → Otwórz**, typ pliku *STEP (*.stp, *.ste, *.step)*, wybierz `fb-jelcz.step`.
2. W opcjach importu ustaw **Typ obiektu: Część (Part)**, żeby powstał jeden plik części z wieloma bryłami (multi-body).
3. **Plik → Zapisz jako** → `fb-jelcz.ipt`.

## Inventor – wariant 2 (DXF → szkic → wyciągnięcie)
1. Nowa część (`Standard (mm).ipt`), **Rozpocznij szkic 2D** na płaszczyźnie XY.
2. **Wstaw → ACAD** (Insert AutoCAD file), wybierz `fb-jelcz.dxf`, jednostki **mm**.
3. Zakończ szkic → **Wyciągnięcie** (Extrude) zamkniętych profili na wybraną grubość.

## Inny rozmiar lub grubość
```
pip install opencv-python-headless ezdxf cadquery
python3 make_logo.py 50 5   # wysokość napisu 50 mm, grubość 5 mm
```
