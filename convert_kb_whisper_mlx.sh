#!/bin/bash
# Convert KBLab/kb-whisper-large (HF/PyTorch) to MLX for Apple Silicon.
# Runs standalone in a scratch venv so nothing touches the tool's own .venv.
set -euo pipefail

WORK="/private/tmp/claude-501/-Users-erik-Documents-Anticlaw/2b2d96e9-bd3d-4051-a05c-e5449ef8777b/scratchpad/kbconv"
OUT="$HOME/.cache/lab37-whisper/kb-whisper-large-mlx"
LOG="$WORK/convert.log"

mkdir -p "$WORK"
exec > >(tee -a "$LOG") 2>&1

echo "=== KB-Whisper -> MLX conversion started $(date) ==="

if [ -f "$OUT/weights.safetensors" ] || [ -f "$OUT/weights.npz" ]; then
    echo "Already converted at $OUT"
    exit 0
fi

# 1. Scratch venv with the conversion deps (torch is needed to read the HF
#    checkpoint; it is NOT a runtime dependency of the tool).
if [ ! -d "$WORK/venv" ]; then
    echo "--- creating venv ---"
    python3 -m venv "$WORK/venv"
fi
PY="$WORK/venv/bin/python"
"$PY" -m pip install --quiet --upgrade pip
echo "--- installing deps (mlx, torch, transformers) ---"
"$PY" -m pip install --quiet mlx numpy tqdm huggingface_hub transformers torch

# 2. mlx-examples carries the whisper converter (not shipped in the pip pkg).
if [ ! -d "$WORK/mlx-examples" ]; then
    echo "--- cloning mlx-examples ---"
    git clone --depth 1 https://github.com/ml-explore/mlx-examples "$WORK/mlx-examples"
fi

# 3. Convert. float16 matches the stock mlx-community large-v3 build.
echo "--- converting (this is the long part) ---"
mkdir -p "$(dirname "$OUT")"
cd "$WORK/mlx-examples/whisper"
"$PY" convert.py \
    --torch-name-or-path KBLab/kb-whisper-large \
    --mlx-path "$OUT" \
    --dtype float16

echo "=== conversion finished $(date) ==="
ls -la "$OUT"
