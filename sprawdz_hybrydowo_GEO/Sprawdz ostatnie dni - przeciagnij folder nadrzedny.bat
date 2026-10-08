@echo off
rem Sprawdza WSZYSTKIE foldery wydan zmienione/utworzone w ostatnich N dniach (domyslnie 14)
rem w przeciagnietym folderze nadrzednym, np. S:\14_Gotowe_Projekty\GJERSTAD.
rem Wynik: jeden zbiorczy raport RAPORT_SPRAWDZENIA_<data>.xlsx + podglad NAKLADKI_<data>.pdf.
rem Inna liczba dni (np. 30) - w oknie cmd:
rem   "Sprawdz ostatnie dni - przeciagnij folder nadrzedny.bat" "S:\14_Gotowe_Projekty\GJERSTAD" 30
cd /d "%~dp0"
set "DNI=14"
if not "%~2"=="" set "DNI=%~2"
if "%~1"=="" (
    echo Przeciagnij folder nadrzedny z wydaniami, np. S:\14_Gotowe_Projekty\GJERSTAD, na ten plik .bat
    pause
    exit /b
)
if not exist "%~dp0sprawdz_hybrydowo_GEO.exe" (
    echo BLAD: nie znaleziono sprawdz_hybrydowo_GEO.exe obok tego pliku .bat
    pause
    exit /b 1
)
set "NADRZEDNY=%~1"
set "EXE=%~dp0sprawdz_hybrydowo_GEO.exe"
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$od = (Get-Date).AddDays(-[int]$env:DNI);" ^
  "$f = @(Get-ChildItem -LiteralPath $env:NADRZEDNY -Directory | Where-Object { $_.LastWriteTime -gt $od -or $_.CreationTime -gt $od } | Sort-Object Name | ForEach-Object { $_.FullName });" ^
  "if ($f.Count -eq 0) { Write-Host ('Brak folderow zmienionych w ostatnich ' + $env:DNI + ' dniach.'); exit 0 };" ^
  "Write-Host ('Folderow do sprawdzenia (ostatnie ' + $env:DNI + ' dni): ' + $f.Count); $f | ForEach-Object { Write-Host ('   ' + $_) };" ^
  "& $env:EXE @f"
pause
