@echo off
setlocal

REM ──────────────────────────────────────────────
REM Copy the pre-built AgencyProposal.dotm into 
REM Word’s Startup folder so it auto-loads.
REM ──────────────────────────────────────────────

set "TEMPLATE=C:\Agency Proposal\AgencyProposal.dotm"
set "STARTUP=%APPDATA%\Microsoft\Word\STARTUP"

echo Installing AgencyProposal.dotm to Word Startup...
if not exist "%STARTUP%" (
  mkdir "%STARTUP%"
)

copy /Y "%TEMPLATE%" "%STARTUP%" >nul 2>&1
if %ERRORLEVEL% neq 0 (
  echo.
  echo [ERROR] Could not copy the template.
  echo  • Verify "%TEMPLATE%" exists.
  pause
  exit /b 1
)

echo.
echo ✔  Installed!
echo   Please restart Word. Your MAIN and EMAIL macros 
echo   will now be available in every document.
pause
