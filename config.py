import json
import logging
import os
import sys

log = logging.getLogger(__name__)

IS_WIN = sys.platform == "win32"
IS_MAC = sys.platform == "darwin"
IS_LINUX = sys.platform.startswith("linux")

_APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(_APP_DIR, "caption_config.json")

DEFAULT_CONFIG = {
    # Transcription
    "language": "sv",           # None = auto-detect, "sv" = Swedish, "en" = English
    "beam_size": 10,            # Higher = more accurate (5 is default, 10 recommended for GPU)
    "initial_prompt": None,     # Swedish context hint, e.g. "Intervju om ledarskap och projektledning."
    # SRT formatting
    "max_words_per_caption": 0,      # 0 = no limit (use chars/lines), >0 = hard word cap
    "max_chars_per_line": 42,        # 42 = broadcast standard (EBU/Netflix) for 16:9
    "max_lines": 1,                  # 1 = single-line captions (clean for podcast/interview)
    "min_duration_s": 1.0,
    "max_duration_s": 7.0,
    "gap_frames": 2,            # gap between subtitles in frames
    # Output
    "output_dir": None,         # None = same dir as source / project
}


def load_config() -> dict:
    config = dict(DEFAULT_CONFIG)
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                saved = json.load(f)
            config.update(saved)
        except Exception as e:
            log.warning(f"Failed to load config: {e}")
    return config


cfg = load_config()

# Model constants
if IS_MAC:
    MODEL_SIZE = "mlx-community/whisper-large-v3-turbo"
    DEVICE = "mlx"
    COMPUTE_TYPE = None
else:
    MODEL_SIZE = "large-v3"
    try:
        import torch
        _cuda_ok = torch.cuda.is_available()
    except Exception:
        _cuda_ok = False

    if _cuda_ok:
        DEVICE = "cuda"
        COMPUTE_TYPE = "float16"
    else:
        DEVICE = "cpu"
        COMPUTE_TYPE = "int8"
        log.warning("CUDA not available - using CPU mode (slower transcription)")
