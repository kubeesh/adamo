@echo off
rem Kontrola folderu wydania hybrydowego + GEOMETRIA DXF (wersja EXE - bez Pythona).
rem Przeciagnij jeden lub wiecej folderow na ten plik.
cd /d "%~dp0"
if "%~1"=="" (
    echo Przeciagnij folder projektu na ten plik .bat
    pause
    exit /b
)

if not exist "%~dp0sprawdz_hybrydowo_GEO.exe" (
    echo BLAD: nie znaleziono sprawdz_hybrydowo_GEO.exe w tym folderze.
    echo Upewnij sie, ze plik .exe jest w tym samym folderze co ten .bat.
    pause
    exit /b 1
)

:petla
if "%~1"=="" goto koniec
if exist "%~1\" (
    "%~dp0sprawdz_hybrydowo_GEO.exe" "%~1"
) else (
    echo Pominieto - nie istnieje: %~1
)
shift
goto petla
:koniec
pause
