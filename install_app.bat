@echo off
SETLOCAL ENABLEEXTENSIONS

REM Define paths
SET "TARGET_DIR=C:\Agency Proposal"
SET "SHORTCUT_NAME=Agency Proposal.lnk"
SET "DESKTOP_PATH=%USERPROFILE%\Desktop"
SET "QUICKLAUNCH_PATH=%APPDATA%\Microsoft\Internet Explorer\Quick Launch"

REM Create the target directory if it doesn't exist
IF NOT EXIST "%TARGET_DIR%" (
    mkdir "%TARGET_DIR%"
)

REM Copy all files to the target directory
xcopy * "%TARGET_DIR%" /E /I /Y

REM Change to the target directory
cd /d "%TARGET_DIR%"

REM Install Python dependencies
pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org --trusted-host pypi.python.org -r requirements.txt

REM Copy shortcut to Desktop
IF EXIST "%SHORTCUT_NAME%" (
    echo Copying shortcut to Desktop...
    copy "%SHORTCUT_NAME%" "%DESKTOP_PATH%" /Y
) ELSE (
    echo Shortcut not found: %SHORTCUT_NAME%
)

REM Copy shortcut to Quick Launch
IF EXIST "%SHORTCUT_NAME%" (
    echo Copying shortcut to Quick Launch...
    copy "%SHORTCUT_NAME%" "%QUICKLAUNCH_PATH%" /Y
) ELSE (
    echo Shortcut not found: %SHORTCUT_NAME%
)

echo Installation complete. Press any key to exit.
pause >nul
