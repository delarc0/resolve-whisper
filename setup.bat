@echo off
setlocal

echo.
echo   ___       _   ___  ____  _____
echo  / / |     / / / _ )/__  //___  /
echo / /  ^|    / / / _  ^|  / /    / /
echo / /___^|   / / / ____/ / /    / /
echo /_____/  /_/ /_/     /_/    /_/
echo.
echo  LAB37 TOOLS: Resolve Whisper
echo  AI-powered captions for DaVinci Resolve
echo  ========================================
echo.

:: Check for Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  ERROR: Python not found.
    echo  Install Python 3.10+ from python.org and check "Add to PATH".
    echo.
    pause
    exit /b 1
)

:: Get this script's directory
set "APP_DIR=%~dp0"
set "APP_DIR=%APP_DIR:~0,-1%"

echo  [1/4] Creating virtual environment...
if not exist "%APP_DIR%\.venv" (
    python -m venv "%APP_DIR%\.venv"
) else (
    echo         Already exists, skipping.
)

echo  [2/4] Installing dependencies (this takes a few minutes)...
"%APP_DIR%\.venv\Scripts\python.exe" -m pip install --quiet --upgrade pip
"%APP_DIR%\.venv\Scripts\python.exe" -m pip install --quiet torch --index-url https://download.pytorch.org/whl/cu121
"%APP_DIR%\.venv\Scripts\python.exe" -m pip install --quiet faster-whisper numpy

:: Verify install
"%APP_DIR%\.venv\Scripts\python.exe" -c "import faster_whisper; print('  faster-whisper OK')"
if %errorlevel% neq 0 (
    echo  ERROR: Dependencies failed to install. Check your internet connection.
    pause
    exit /b 1
)

:: Resolve Scripts folder
set "RESOLVE_SCRIPTS=%APPDATA%\Blackmagic Design\DaVinci Resolve\Support\Fusion\Scripts\Utility"

echo  [3/4] Installing script to Resolve...
if not exist "%RESOLVE_SCRIPTS%" (
    mkdir "%RESOLVE_SCRIPTS%"
)
copy /y "%APP_DIR%\resolve_script.py" "%RESOLVE_SCRIPTS%\LAB37 Resolve Whisper.py" >nul

:: Write a pointer file so the Resolve script knows where we live
echo %APP_DIR%> "%RESOLVE_SCRIPTS%\resolve_whisper_path.txt"

echo  [4/4] Downloading AI model (first time only, ~3 GB)...
"%APP_DIR%\.venv\Scripts\python.exe" -c "from faster_whisper import WhisperModel; WhisperModel('large-v3', device='cpu', compute_type='int8')" 2>nul
if %errorlevel% neq 0 (
    echo         Model will download on first use instead.
) else (
    echo         Model ready.
)

echo.
echo  ========================================
echo   SETUP COMPLETE
echo  ========================================
echo.
echo  How to use:
echo    1. Open DaVinci Resolve Studio
echo    2. Select a timeline, set in/out points (I and O)
echo    3. Workspace ^> Scripts ^> LAB37 Resolve Whisper
echo    4. Wait for the progress window to finish
echo    5. File ^> Import ^> Subtitle to add the .srt
echo.
echo  Captions are saved to: %USERPROFILE%\Desktop\Captions\
echo.
echo  LAB37 TOOLS // lab37.se
echo.
pause
