#!/bin/bash
set -e

echo ""
echo "  LAB37 TOOLS: Resolve Whisper"
echo "  AI-powered captions for DaVinci Resolve"
echo "  ========================================"
echo ""

# Check for Python
if ! command -v python3 &>/dev/null; then
    echo "  ERROR: Python 3 not found."
    echo "  Install with: brew install python"
    exit 1
fi

APP_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "  [1/4] Creating virtual environment..."
if [ ! -d "$APP_DIR/.venv" ]; then
    python3 -m venv "$APP_DIR/.venv"
else
    echo "         Already exists, skipping."
fi

echo "  [2/4] Installing dependencies..."
"$APP_DIR/.venv/bin/pip" install --quiet --upgrade pip
"$APP_DIR/.venv/bin/pip" install --quiet mlx mlx-whisper soundfile resampy numpy

# Verify
"$APP_DIR/.venv/bin/python3" -c "import mlx_whisper; print('  mlx-whisper OK')" || {
    echo "  ERROR: Dependencies failed to install."
    exit 1
}

# Resolve Scripts folder
RESOLVE_SCRIPTS="$HOME/Library/Application Support/Blackmagic Design/DaVinci Resolve/Fusion/Scripts/Utility"

echo "  [3/4] Installing script to Resolve..."
mkdir -p "$RESOLVE_SCRIPTS"
cp "$APP_DIR/resolve_script.py" "$RESOLVE_SCRIPTS/LAB37 Resolve Whisper.py"

# Write pointer file
echo "$APP_DIR" > "$RESOLVE_SCRIPTS/resolve_whisper_path.txt"

echo "  [4/4] Downloading AI model (first time only, ~1.5 GB)..."
"$APP_DIR/.venv/bin/python3" -c "
import mlx_whisper, numpy as np
mlx_whisper.transcribe(np.zeros(16000, dtype=np.float32), path_or_hf_repo='mlx-community/whisper-large-v3-turbo')
" 2>/dev/null && echo "         Model ready." || echo "         Model will download on first use instead."

echo ""
echo "  ========================================"
echo "   SETUP COMPLETE"
echo "  ========================================"
echo ""
echo "  How to use:"
echo "    1. Open DaVinci Resolve Studio"
echo "    2. Select a timeline, set in/out points (I and O)"
echo "    3. Workspace > Scripts > LAB37 Resolve Whisper"
echo "    4. Wait for the progress window to finish"
echo "    5. File > Import > Subtitle to add the .srt"
echo ""
echo "  Captions are saved to: ~/Desktop/Captions/"
echo ""
echo "  LAB37 TOOLS // lab37.se"
echo ""
