import json
import logging
import os
import sys

log = logging.getLogger(__name__)

IS_WIN = sys.platform == "win32"
IS_MAC = sys.platform == "darwin"

_APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(_APP_DIR, "caption_config.json")

DEFAULT_CONFIG = {
    # Transcription
    "language": None,           # None = auto-detect
    "beam_size": 5,
    # SRT formatting
    "max_words_per_caption": 0,      # 0 = no limit (use chars/lines), >0 = hard word cap
    "max_chars_per_line": 42,
    "max_lines": 1,
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
    MODEL_SIZE = "mlx-community/whisper-large-v3"
    DEVICE = "mlx"
    COMPUTE_TYPE = None
else:
    MODEL_SIZE = "large-v3"
    DEVICE = "cuda"
    COMPUTE_TYPE = "float16"
