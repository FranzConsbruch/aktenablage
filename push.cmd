@echo off
setlocal
cd /d "%~dp0"

where git >nul 2>&1
if errorlevel 1 (
  echo Git ist nicht installiert oder nicht im PATH.
  echo Download: https://git-scm.com/download/win
  pause
  exit /b 1
)

if not exist ".git" (
  echo Erste Einrichtung: lokales Repository wird angelegt...
  git init
  git branch -M main
  git remote add origin https://github.com/FranzConsbruch/aktenablage.git
  echo Stand von GitHub holen...
  git fetch origin
  git rev-parse --verify origin/main >nul 2>&1
  if not errorlevel 1 (
    echo Vorhandene GitHub-Historie uebernehmen, lokale Dateien behalten...
    git reset --soft origin/main
  )
)

set "MSG=%~1"
if "%MSG%"=="" set "MSG=Aktualisierung %DATE% %TIME%"

git add -A
git diff --cached --quiet
if not errorlevel 1 (
  echo Keine Aenderungen gefunden - es gibt nichts zu pushen.
  pause
  exit /b 0
)

echo.
echo Folgende Dateien werden gepusht:
git diff --cached --name-status
echo.

git commit -m "%MSG%"
git push -u origin main
if errorlevel 1 (
  echo.
  echo Push fehlgeschlagen.
  echo Beim ersten Mal fragt Git nach der GitHub-Anmeldung - im Browserfenster
  echo anmelden und push.cmd danach noch einmal ausfuehren.
  pause
  exit /b 1
)

echo.
echo Fertig. GitHub Pages aktualisiert sich in ein bis zwei Minuten.
echo Danach im Outlook das Panel schliessen und neu oeffnen.
pause
