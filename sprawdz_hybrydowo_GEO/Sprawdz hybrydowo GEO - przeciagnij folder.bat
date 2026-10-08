@echo off
rem Kontrola folderu wydania hybrydowego + GEOMETRIA DXF + nakladka DXF na rysunek (wersja EXE - bez Pythona).
rem Przeciagnij jeden lub wiecej folderow wydan na ten plik - powstanie JEDEN wspolny raport.
cd /d "%~dp0"
echo ============================================================
echo  SPRAWDZ HYBRYDOWO GEO
echo ============================================================
if "%~1"=="" (
    echo Przeciagnij folder wydania ^(np. GE21755-S1^) na ten plik .bat
    echo - nie uruchamiaj go dwuklikiem.
    pause
    exit /b
)
if not exist "%~dp0sprawdz_hybrydowo_GEO.exe" (
    echo BLAD: nie znaleziono sprawdz_hybrydowo_GEO.exe w folderze:
    echo    %~dp0
    echo Zbuduj go plikiem zbuduj_exe.bat albo skopiuj .exe obok tego pliku .bat.
    pause
    exit /b 1
)
set "LISTA="
set "ILE=0"
:petla
if "%~1"=="" goto uruchom
if exist "%~1\" (
    echo   do sprawdzenia: %~1
    set LISTA=%LISTA% "%~1"
    set /a ILE+=1
) else (
    echo   POMINIETO - to nie jest folder: %~1
)
shift
goto petla
:uruchom
if "%ILE%"=="0" (
    echo Nie ma zadnego folderu do sprawdzenia.
    pause
    exit /b 1
)
echo.
echo Uruchamiam program - pierwsze uruchomienie .exe trwa do ok. 30 s, prosze czekac...
echo.
"%~dp0sprawdz_hybrydowo_GEO.exe"%LISTA%
echo.
echo Zakonczono. Kod wyjscia: %ERRORLEVEL%  (0 = bez bledow, 1 = sa bledy albo cos nie zadzialalo)
pause
