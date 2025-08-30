@echo off
echo ========================================
echo PAS WebService Installation
echo ========================================
echo.

REM Administrator-Rechte prüfen
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo FEHLER: Dieses Script muss als Administrator ausgeführt werden!
    pause
    exit /b 1
)

set SERVICE_NAME=PAS_WebService
set SERVICE_DISPLAY=PAS Web-Service
set SERVICE_DESC=Patientenaufrufsystem Web-Service für Anzeige und API
set SERVICE_PATH=%~dp0PAS_WebService.exe

echo Service-Informationen:
echo Name: %SERVICE_NAME%
echo Pfad: %SERVICE_PATH%
echo.

REM Prüfen ob Service bereits existiert
sc query %SERVICE_NAME% >nul 2>&1
if %errorlevel% equ 0 (
    echo Service existiert bereits. Wird gestoppt und entfernt...
    sc stop %SERVICE_NAME% >nul 2>&1
    timeout /t 2 /nobreak >nul
    sc delete %SERVICE_NAME%
    timeout /t 2 /nobreak >nul
)

echo Installiere Service...
REM NSSM verwenden falls vorhanden, sonst sc
where nssm >nul 2>&1
if %errorlevel% equ 0 (
    echo Verwende NSSM für Installation...
    nssm install %SERVICE_NAME% "%SERVICE_PATH%"
    nssm set %SERVICE_NAME% DisplayName "%SERVICE_DISPLAY%"
    nssm set %SERVICE_NAME% Description "%SERVICE_DESC%"
    nssm set %SERVICE_NAME% Start SERVICE_AUTO_START
    nssm set %SERVICE_NAME% AppStdout "%~dp0Logs\service.log"
    nssm set %SERVICE_NAME% AppStderr "%~dp0Logs\error.log"
    nssm set %SERVICE_NAME% AppRotateFiles 1
    nssm set %SERVICE_NAME% AppRotateBytes 10485760
) else (
    echo NSSM nicht gefunden. Verwende SC für Installation...
    echo.
    echo HINWEIS: Für bessere Service-Verwaltung wird NSSM empfohlen!
    echo Download: https://nssm.cc/download
    echo.
    sc create %SERVICE_NAME% binPath= "%SERVICE_PATH%" DisplayName= "%SERVICE_DISPLAY%" start= auto
    sc description %SERVICE_NAME% "%SERVICE_DESC%"
)

echo.
echo Konfiguriere Firewall...
netsh advfirewall firewall delete rule name="PAS WebService HTTP" >nul 2>&1
netsh advfirewall firewall add rule name="PAS WebService HTTP" dir=in action=allow protocol=TCP localport=8080

echo.
echo Starte Service...
sc start %SERVICE_NAME%

timeout /t 2 /nobreak >nul
sc query %SERVICE_NAME%

echo.
echo ========================================
echo Installation abgeschlossen!
echo.
echo Der Service läuft auf:
echo http://localhost:8080/
echo.
echo Service-Verwaltung:
echo - Starten:  sc start %SERVICE_NAME%
echo - Stoppen:  sc stop %SERVICE_NAME%
echo - Status:   sc query %SERVICE_NAME%
echo - Löschen:  sc delete %SERVICE_NAME%
echo ========================================
pause