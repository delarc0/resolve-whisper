@echo off
setlocal

echo.
echo  ========================================
echo   Resolve Whisper - Setup
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

echo  [2/4] Installing dependencies...
"%APP_DIR%\.venv\Scripts\pip.exe" install --quiet --upgrade pip
"%APP_DIR%\.venv\Scripts\pip.exe" install --quiet torch --index-url https://download.pytorch.org/whl/cu121
"%APP_DIR%\.venv\Scripts\pip.exe" install --quiet faster-whisper numpy

:: Resolve Scripts folder
set "RESOLVE_SCRIPTS=%APPDATA%\Blackmagic Design\DaVinci Resolve\Support\Fusion\Scripts\Utility"

echo  [3/4] Installing Resolve script...
if not exist "%RESOLVE_SCRIPTS%" (
    mkdir "%RESOLVE_SCRIPTS%"
)
copy /y "%APP_DIR%\resolve_script.py" "%RESOLVE_SCRIPTS%\Resolve Whisper.py" >nul

:: Write a pointer file so the Resolve script knows where we live
echo %APP_DIR%> "%RESOLVE_SCRIPTS%\resolve_whisper_path.txt"

echo  [4/4] Downloading AI model (first time only, ~1.5 GB)...
"%APP_DIR%\.venv\Scripts\python.exe" -c "from faster_whisper import WhisperModel; WhisperModel('deepdml/faster-whisper-large-v3-turbo-ct2', device='cpu', compute_type='int8')" 2>nul
if %errorlevel% neq 0 (
    echo         Model will download on first use instead.
) else (
    echo         Model ready.
)

echo.
echo  ========================================
echo   DONE
echo  ========================================
echo.
echo  How to use:
echo    1. Open DaVinci Resolve Studio
echo    2. Open a project and select a timeline
echo    3. Go to: Workspace ^> Scripts ^> Resolve Whisper
echo    4. Wait for transcription to finish
echo    5. SRT file opens automatically
echo.
echo  Captions are saved to: %USERPROFILE%\Desktop\Captions\
echo.
pause
