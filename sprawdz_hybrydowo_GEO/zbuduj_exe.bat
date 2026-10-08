@echo off
rem Buduje sprawdz_hybrydowo_GEO.exe z pliku .py lezacego w tym samym folderze.
rem Jeden plik .exe, bez bibliotek, ktorych program nie uzywa (mniejszy rozmiar).
cd /d "%~dp0"
python -m pip install --upgrade pyinstaller openpyxl pymupdf ezdxf pywin32 || goto blad
python -m PyInstaller --noconfirm --onefile --console --name sprawdz_hybrydowo_GEO ^
  --exclude-module matplotlib --exclude-module tkinter --exclude-module PyQt5 --exclude-module PySide6 ^
  --exclude-module IPython --exclude-module pandas --exclude-module scipy --exclude-module pytest ^
  sprawdz_hybrydowo_GEO.py || goto blad
copy /y dist\sprawdz_hybrydowo_GEO.exe . >/dev/null || goto blad
echo.
echo Gotowe: sprawdz_hybrydowo_GEO.exe (obok tego pliku .bat)
pause
exit /b 0
:blad
echo.
echo BLAD budowania - zobacz komunikaty wyzej.
pause
exit /b 1
