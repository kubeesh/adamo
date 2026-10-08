@echo off
rem Buduje sprawdz_hybrydowo_GEO.exe z pliku sprawdz_hybrydowo_GEO.py lezacego w tym samym folderze
rem i kopiuje go obok tego pliku .bat - tam, gdzie szukaja go pliki "Sprawdz ... .bat".
rem Jeden plik .exe, bez bibliotek, ktorych program nie uzywa (mniejszy rozmiar).
setlocal
set "TU=%~dp0"
cd /d "%TU%" || goto blad
if not exist "%TU%sprawdz_hybrydowo_GEO.py" (
    echo BLAD: brak pliku sprawdz_hybrydowo_GEO.py w folderze:
    echo    %TU%
    goto blad
)
echo [1/3] Instaluje/aktualizuje biblioteki...
python -m pip install --upgrade pyinstaller openpyxl pymupdf ezdxf pywin32 || goto blad
echo.
echo [2/3] Buduje .exe - to trwa ok. 1 minute...
python -m PyInstaller --noconfirm --onefile --console --name sprawdz_hybrydowo_GEO ^
  --distpath "%TU%dist" --workpath "%TU%build" --specpath "%TU%build" ^
  --exclude-module matplotlib --exclude-module tkinter --exclude-module PyQt5 --exclude-module PySide6 ^
  --exclude-module IPython --exclude-module pandas --exclude-module scipy --exclude-module pytest ^
  "%TU%sprawdz_hybrydowo_GEO.py" || goto blad
echo.
echo [3/3] Kopiuje nowy .exe obok plikow .bat...
if not exist "%TU%dist\sprawdz_hybrydowo_GEO.exe" (
    echo BLAD: po budowaniu nie ma pliku:
    echo    %TU%dist\sprawdz_hybrydowo_GEO.exe
    echo Najczestsza przyczyna: antywirus usunal nowy plik .exe - sprawdz
    echo Zabezpieczenia Windows, Ochrona przed wirusami, Historia ochrony.
    goto blad
)
copy /y "%TU%dist\sprawdz_hybrydowo_GEO.exe" "%TU%sprawdz_hybrydowo_GEO.exe"
if errorlevel 1 (
    echo BLAD kopiowania. Zamknij okna, w ktorych dziala stary sprawdz_hybrydowo_GEO.exe, i uruchom ten plik
    echo jeszcze raz - albo skopiuj recznie plik:
    echo    %TU%dist\sprawdz_hybrydowo_GEO.exe
    echo do folderu:
    echo    %TU%
    goto blad
)
echo.
echo Gotowe: %TU%sprawdz_hybrydowo_GEO.exe
echo Teraz przeciagnij folder wydania na "Sprawdz hybrydowo GEO - przeciagnij folder.bat".
pause
exit /b 0
:blad
echo.
echo BLAD budowania - zobacz komunikaty wyzej.
pause
exit /b 1
