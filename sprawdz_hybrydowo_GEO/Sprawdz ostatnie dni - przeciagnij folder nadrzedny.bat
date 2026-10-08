@echo off
rem Sprawdza WSZYSTKIE foldery wydan zmienione/utworzone w ostatnich N dniach (domyslnie 14)
rem w przeciagnietym folderze NADRZEDNYM, np. S:\14_Gotowe_Projekty\GJERSTAD.
rem Jesli przeciagniesz pojedynczy folder wydania (z plikiem BOM .xlsx), sprawdzi po prostu ten folder.
rem Wynik: jeden zbiorczy raport RAPORT_SPRAWDZENIA_<data>.xlsx + podglad NAKLADKI_<data>.pdf.
rem Inna liczba dni (np. 30) - w oknie cmd:
rem   "Sprawdz ostatnie dni - przeciagnij folder nadrzedny.bat" "S:\14_Gotowe_Projekty\GJERSTAD" 30
cd /d "%~dp0"
echo ============================================================
echo  SPRAWDZ HYBRYDOWO GEO - foldery z ostatnich dni
echo ============================================================
set "DNI=14"
if not "%~2"=="" set "DNI=%~2"
if "%~1"=="" (
    echo Przeciagnij folder nadrzedny z wydaniami, np. S:\14_Gotowe_Projekty\GJERSTAD, na ten plik .bat
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
set "NADRZEDNY=%~1"
set "EXE=%~dp0sprawdz_hybrydowo_GEO.exe"
echo Folder: %~1
echo Szukam folderow wydan z ostatnich %DNI% dni...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$od = (Get-Date).AddDays(-[int]$env:DNI);" ^
  "if (Get-ChildItem -LiteralPath $env:NADRZEDNY -File -Filter *.xlsx -ErrorAction SilentlyContinue) { Write-Host 'To jest pojedynczy folder wydania (ma BOM .xlsx) - sprawdzam ten folder.'; $f = @($env:NADRZEDNY) }" ^
  "else { $f = @(Get-ChildItem -LiteralPath $env:NADRZEDNY -Directory | Where-Object { $_.LastWriteTime -gt $od -or $_.CreationTime -gt $od } | Sort-Object Name | ForEach-Object { $_.FullName }) };" ^
  "if ($f.Count -eq 0) { Write-Host ('Brak folderow zmienionych w ostatnich ' + $env:DNI + ' dniach w: ' + $env:NADRZEDNY); exit 0 };" ^
  "Write-Host ('Folderow do sprawdzenia: ' + $f.Count); $f | ForEach-Object { Write-Host ('   ' + $_) };" ^
  "Write-Host ''; Write-Host 'Uruchamiam program - pierwsze uruchomienie .exe trwa do ok. 30 s, prosze czekac...'; Write-Host '';" ^
  "& $env:EXE @f; exit $LASTEXITCODE"
if errorlevel 9009 echo BLAD: nie mozna uruchomic PowerShella - uzyj "Sprawdz hybrydowo GEO - przeciagnij folder.bat" i przeciagnij foldery recznie.
echo.
echo Zakonczono.
pause
