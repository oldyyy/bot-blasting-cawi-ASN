@echo off
REM Buka Chrome dengan port remote debugging agar bot bisa menumpang. Login SSO dilakukan manual.
REM Tiap port memakai profil sendiri:  buka_chrome.bat 9222 , buka_chrome.bat 9223 , dst.
set PORT=%1
if "%PORT%"=="" set PORT=9222
set PROFILE=%~dp0.chrome-cdp-profile
if not "%PORT%"=="9222" set PROFILE=%PROFILE%-%PORT%
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" ^
  --remote-debugging-port=%PORT% ^
  --user-data-dir="%PROFILE%" ^
  "https://fasih-sm.bps.go.id/"
