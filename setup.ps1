# LAB37 TOOLS: Resolve Whisper (Windows installer)
# AI-powered captions for DaVinci Resolve Studio.
#
# Run from PowerShell:
#   cd path\to\resolve-whisper
#   powershell -ExecutionPolicy Bypass -File .\setup.ps1
#
# Windows uses the faster-whisper backend (CUDA if an NVIDIA GPU is present,
# otherwise CPU). The Mac mlx path is not used here.

$ErrorActionPreference = "Stop"
# Under PowerShell 7+ a non-zero native exit becomes a terminating error when
# ErrorActionPreference is Stop. Our tolerant exit-code checks (the import
# probe and the model pre-download) rely on non-zero NOT throwing, so disable
# that. Harmless on Windows PowerShell 5.1 (the variable is just ignored).
$PSNativeCommandUseErrorActionPreference = $false

Write-Host ""
Write-Host "  LAB37 TOOLS: Resolve Whisper (Windows)"
Write-Host "  AI-powered captions for DaVinci Resolve"
Write-Host "  ========================================"
Write-Host ""

$AppDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# --- Python ---
$python = $null
foreach ($cand in @("python", "python3", "py")) {
    $cmd = Get-Command $cand -ErrorAction SilentlyContinue
    if ($cmd) { $python = $cmd.Source; break }
}
if (-not $python) {
    Write-Host "  ERROR: Python 3 not found. Install Python 3.10+ from python.org"
    Write-Host "  (check 'Add python.exe to PATH' during install), then re-run."
    exit 1
}
$pyVer = & $python -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')"
if (-not $pyVer) {
    # The Microsoft Store "app execution alias" stub matches Get-Command but
    # produces no output. Fail with the friendly message, not a raw exception.
    Write-Host "  ERROR: Python 3 not found (the Microsoft Store stub does not count)."
    Write-Host "  Install Python 3.10+ from python.org and tick 'Add python.exe to PATH'."
    exit 1
}
$maj, $min = $pyVer.Split(".")
if ([int]$maj -lt 3 -or ([int]$maj -eq 3 -and [int]$min -lt 10)) {
    Write-Host "  ERROR: Python 3.10+ required. Found: $pyVer"
    exit 1
}
Write-Host "  Python $pyVer OK"

# --- ffmpeg (optional but recommended) ---
# The faster-whisper backend decodes audio itself (bundled PyAV), so ffmpeg
# isn't strictly required. It IS used for the duration probe fallback, which
# keeps the progress bar accurate on non-WAV renders. Install it if easy.
if (-not (Get-Command ffmpeg -ErrorAction SilentlyContinue)) {
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Write-Host "  ffmpeg not found; installing via winget (optional)..."
        winget install --id Gyan.FFmpeg -e --accept-source-agreements --accept-package-agreements
        Write-Host "        NOTE: you may need a new terminal for ffmpeg to be on PATH."
    } else {
        Write-Host "  ffmpeg not found (optional). Get it from https://ffmpeg.org/download.html if the progress bar looks off."
    }
} else {
    Write-Host "  ffmpeg OK"
}

# --- GPU detection ---
# The Windows backend is faster-whisper (CTranslate2). For GPU it needs the
# CUDA cuBLAS + cuDNN libraries; the CUDA build of PyTorch ships those DLLs
# (and gives us the cuda-availability probe), and transcribe.py adds torch's
# lib dir to the DLL search path so CTranslate2 can load them. No NVIDIA GPU
# -> CPU-only install (works everywhere, slower on long timelines).
$hasNvidia = $false
if (Get-Command nvidia-smi -ErrorAction SilentlyContinue) {
    try { & nvidia-smi | Out-Null; $hasNvidia = $true } catch { $hasNvidia = $false }
}
if ($hasNvidia) {
    Write-Host "  NVIDIA GPU detected -> GPU (CUDA) mode"
} else {
    Write-Host "  No NVIDIA GPU detected -> CPU mode (slower on long timelines)"
}

Write-Host "  [1/4] Creating virtual environment..."
$venv = Join-Path $AppDir ".venv"
if (-not (Test-Path $venv)) {
    & $python -m venv $venv
} else {
    Write-Host "         Already exists, skipping."
}
$venvPy = Join-Path $venv "Scripts\python.exe"

Write-Host "  [2/4] Installing dependencies..."
& $venvPy -m pip install --quiet --upgrade pip
if ($hasNvidia) {
    # CUDA build of PyTorch: provides torch.cuda.is_available() (the device
    # probe in config.py) and bundles the cuBLAS/cuDNN DLLs CTranslate2 needs.
    # cu121 ships cuDNN 9, which matches faster-whisper's CTranslate2 4.x.
    Write-Host "         Installing CUDA PyTorch (provides cuDNN/cuBLAS for GPU)..."
    & $venvPy -m pip install --quiet torch --index-url https://download.pytorch.org/whl/cu121
}
& $venvPy -m pip install --quiet -r (Join-Path $AppDir "requirements.txt")

& $venvPy -c "import faster_whisper" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "  ERROR: Dependencies failed to install."
    exit 1
}

# --- Resolve Scripts folder ---
# Windows Resolve reads Edit-page menu scripts from %APPDATA%\...\Fusion\Scripts\Edit.
$ResolveScripts = Join-Path $env:APPDATA "Blackmagic Design\DaVinci Resolve\Support\Fusion\Scripts\Edit"

Write-Host "  [3/4] Installing Lua launchers to Resolve..."
New-Item -ItemType Directory -Force -Path $ResolveScripts | Out-Null
# Remove legacy installs. The old Windows installer put an in-process Tk stub
# in Scripts\Utility; it runs inside Resolve and freezes it, and Resolve 21
# shows Utility scripts in the menu again, so users can click the wrong entry.
$legacyUtility = Join-Path $env:APPDATA "Blackmagic Design\DaVinci Resolve\Support\Fusion\Scripts\Utility"
foreach ($stale in @(
    (Join-Path $legacyUtility "LAB37 Resolve Whisper.py"),
    (Join-Path $legacyUtility "resolve_whisper_path.txt"),
    (Join-Path $ResolveScripts "_launcher.lua")
)) {
    Remove-Item -Force -ErrorAction SilentlyContinue $stale
}
foreach ($p in @("LAB37 Reels.lua", "LAB37 Podcast.lua", "LAB37 Auto.lua", "LAB37 Custom.lua", "LAB37 Check.lua")) {
    Copy-Item -Force (Join-Path $AppDir "presets\$p") $ResolveScripts
}
# Pointer file: presets read this to find the install dir (and the shared
# preset_launch.lua that lives there).
Set-Content -Path (Join-Path $ResolveScripts "resolve_whisper_path.txt") -Value $AppDir -NoNewline

# Pre-create the captions output dir referenced in the finish message.
$captionsDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "Captions"
New-Item -ItemType Directory -Force -Path $captionsDir | Out-Null

Write-Host "  [4/4] Downloading AI model (first time only, ~3 GB)..."
# Pre-fetch the model the shipped presets actually use. The Swedish presets
# route to KB-Whisper, so fetching only large-v3 left a second ~3 GB download
# to happen mid-run, after the timeline audio had already been rendered.
$dl = @"
import sys
sys.path.insert(0, r'$AppDir')
from faster_whisper import WhisperModel
import config
names = {config.MODEL_SIZE}
if config.cfg.get('use_kb_whisper', True):
    names.add(config.KB_WHISPER_CT2)
for n in sorted(names):
    print('  downloading', n)
    WhisperModel(n, device='cpu', compute_type='int8')
"@
& $venvPy -c $dl 2>$null
if ($LASTEXITCODE -eq 0) { Write-Host "         Model ready." } else { Write-Host "         Model will download on first use instead." }

Write-Host ""
Write-Host "  ========================================"
Write-Host "   SETUP COMPLETE"
Write-Host "  ========================================"
Write-Host ""
Write-Host "  How to use:"
Write-Host "    1. Open DaVinci Resolve Studio"
Write-Host "    2. Select a timeline, set in/out points (I and O)"
Write-Host "    3. Workspace > Scripts > Edit > LAB37 (pick a preset):"
Write-Host "       - LAB37 Reels    -- single-line SRT, 1-3 words, no punctuation"
Write-Host "       - LAB37 Podcast  -- plain SRT, full sentences"
Write-Host "       - LAB37 Auto     -- auto-detect language, plain SRT"
Write-Host "       - LAB37 Custom   -- settings dialog (language, style, line length)"
Write-Host "       - LAB37 Check    -- pre-flight check (run after install/Resolve update)"
Write-Host "    4. Captions are placed on a subtitle track automatically"
Write-Host "    5. If auto-place was skipped, the SRT is in the Media Pool:"
Write-Host "       right-click it > Insert Selected Subtitles to Timeline"
Write-Host ""
Write-Host "  Captions are saved to: $captionsDir"
Write-Host "  Live progress log:      $env:TEMP\resolve_whisper.log"
Write-Host ""
Write-Host "  LAB37 TOOLS // lab37.tools"
Write-Host ""
