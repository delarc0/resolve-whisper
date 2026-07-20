#!/bin/bash
set -e

echo ""
echo "  LAB37 TOOLS: Resolve Whisper (Mac)"
echo "  AI-powered captions for DaVinci Resolve"
echo "  ========================================"
echo ""

# --- Python ---
if ! command -v python3 &>/dev/null; then
    echo "  ERROR: Python 3 not found."
    echo "  Install with: brew install python@3.12"
    exit 1
fi

PY_VER=$(python3 -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')")
PY_MAJ=$(echo "$PY_VER" | cut -d. -f1)
PY_MIN=$(echo "$PY_VER" | cut -d. -f2)
if [ "$PY_MAJ" -lt 3 ] || ([ "$PY_MAJ" -eq 3 ] && [ "$PY_MIN" -lt 10 ]); then
    echo "  ERROR: Python 3.10+ required. Found: $PY_VER"
    echo "  brew install python@3.12, then re-run with PATH=\"/opt/homebrew/opt/python@3.12/libexec/bin:\$PATH\""
    exit 1
fi
echo "  Python $PY_VER OK"

# --- ffmpeg (mlx-whisper needs it for non-WAV decoding) ---
if ! command -v ffmpeg &>/dev/null; then
    echo "  WARN: ffmpeg not in PATH. Installing via brew..."
    if ! command -v brew &>/dev/null; then
        echo "  ERROR: brew not found. Install ffmpeg manually: brew install ffmpeg"
        exit 1
    fi
    brew install ffmpeg
fi
echo "  ffmpeg OK"

# --- Tk bindings (needed for the floating progress window) ---
# Homebrew Python doesn't bundle tk; without this the UI silently falls back
# to log-only mode. Detect first to avoid an unnecessary brew install.
if ! python3 -c "import tkinter" 2>/dev/null; then
    PY_MM="$PY_MAJ.$PY_MIN"
    if command -v brew &>/dev/null; then
        echo "  WARN: tkinter missing. Installing python-tk@$PY_MM..."
        brew install "python-tk@$PY_MM" || \
            echo "  WARN: python-tk@$PY_MM install failed; UI will be log-only."
    else
        echo "  WARN: tkinter missing and brew unavailable; UI will be log-only."
    fi
fi
python3 -c "import tkinter" 2>/dev/null && echo "  tkinter OK" || echo "  tkinter NOT available (UI will fall back to log-only)"

APP_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "  [1/4] Creating virtual environment..."
if [ ! -d "$APP_DIR/.venv" ]; then
    python3 -m venv "$APP_DIR/.venv"
else
    echo "         Already exists, skipping."
fi

echo "  [2/4] Installing dependencies..."
"$APP_DIR/.venv/bin/pip" install --quiet --upgrade pip
# Single source of truth for deps is requirements-mac.txt
"$APP_DIR/.venv/bin/pip" install --quiet -r "$APP_DIR/requirements-mac.txt"

"$APP_DIR/.venv/bin/python3" -c "import mlx_whisper, silero_vad" || {
    echo "  ERROR: Dependencies failed to install."
    exit 1
}

# --- Resolve Scripts folder ---
# Mac Resolve 20 has a known bug where scripts in Utility/ don't appear in
# Workspace > Scripts. Page-specific subfolders (Edit/Comp/Color/Deliver) DO
# work, so we install to Edit/ since captions are an Edit-page task.
# Mac Resolve also doesn't list .py files in the menu, only .lua -- so a
# Lua launcher shells out to caption.py.
RESOLVE_SCRIPTS="$HOME/Library/Application Support/Blackmagic Design/DaVinci Resolve/Fusion/Scripts/Edit"

echo "  [3/4] Installing Lua launchers to Resolve..."
mkdir -p "$RESOLVE_SCRIPTS"
# Clean up legacy installs:
# - _launcher.lua: old shared launcher, junk menu entry (presets are
#   self-contained now)
# - Utility/LAB37 Resolve Whisper.py: the original in-process stub. It runs
#   Tkinter dialogs INSIDE Resolve via fuscript, which freezes Resolve on
#   macOS. Hidden on Resolve 20 (Utility menu bug) but Resolve 21 shows it
#   again, so it must go or users click the wrong entry.
rm -f "$RESOLVE_SCRIPTS/_launcher.lua"
rm -f "$HOME/Library/Application Support/Blackmagic Design/DaVinci Resolve/Fusion/Scripts/Utility/LAB37 Resolve Whisper.py"
# Preset menu entries: one per common workflow
cp "$APP_DIR/presets/LAB37 Reels.lua"   "$RESOLVE_SCRIPTS/"
cp "$APP_DIR/presets/LAB37 Podcast.lua" "$RESOLVE_SCRIPTS/"
cp "$APP_DIR/presets/LAB37 Auto.lua"    "$RESOLVE_SCRIPTS/"
cp "$APP_DIR/presets/LAB37 Custom.lua"  "$RESOLVE_SCRIPTS/"
cp "$APP_DIR/presets/LAB37 Check.lua"   "$RESOLVE_SCRIPTS/"
echo "$APP_DIR" > "$RESOLVE_SCRIPTS/resolve_whisper_path.txt"

# Pre-create the captions output dir referenced in the post-install message
mkdir -p "$HOME/Desktop/Captions"

echo "  [4/4] Downloading AI model (first time only, ~3 GB)..."
"$APP_DIR/.venv/bin/python3" -c "
import mlx_whisper, numpy as np
mlx_whisper.transcribe(np.zeros(16000, dtype=np.float32), path_or_hf_repo='mlx-community/whisper-large-v3-mlx')
" 2>/dev/null && echo "         Model ready." || echo "         Model will download on first use instead."

echo ""
echo "  ========================================"
echo "   SETUP COMPLETE"
echo "  ========================================"
echo ""
echo "  How to use:"
echo "    1. Open DaVinci Resolve Studio"
echo "    2. Select a timeline, set in/out points (I and O)"
echo "    3. Workspace > Scripts > Edit > LAB37 (pick a preset):"
echo "       - LAB37 Reels    -- single-line SRT, 1-3 words, no punctuation"
echo "       - LAB37 Podcast  -- plain SRT, full sentences"
echo "       - LAB37 Auto     -- auto-detect language, plain SRT"
echo "       - LAB37 Custom   -- settings dialog (language, style, line length)"
echo "       - LAB37 Check    -- pre-flight check (run after install/Resolve update)"
echo "    4. Captions are placed on a subtitle track automatically"
echo "    5. If auto-place was skipped, the SRT is in the Media Pool:"
echo "       right-click it > Insert Selected Subtitles to Timeline"
echo "       (never drag an .srt from Finder into Resolve -- it crashes Resolve 21)"
echo ""
echo "  Captions are saved to: ~/Desktop/Captions/ (auto-cleaned after 30 days)"
echo "  Live progress log:      /tmp/resolve_whisper.log"
echo ""
echo "  LAB37 TOOLS // lab37.tools"
echo ""
