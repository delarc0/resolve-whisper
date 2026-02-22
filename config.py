import json
import logging
import os
import sys
import tempfile

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
    "caption_style": "traditional",   # "traditional", "social", "karaoke"
    "max_chars_per_line": 42,
    "max_lines": 2,
    "min_duration_s": 1.0,
    "max_duration_s": 7.0,
    "gap_frames": 2,            # gap between subtitles in frames
    # Confidence
    "confidence_threshold": 0.7,
    "highlight_low_confidence": True,
    # Output
    "output_dir": None,         # None = same dir as source / project
    # Version
    "version": "0.1.0",
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


def save_config(config=None):
    if config is None:
        config = cfg
    try:
        fd, tmp = tempfile.mkstemp(
            dir=os.path.dirname(CONFIG_PATH), suffix=".tmp"
        )
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        os.replace(tmp, CONFIG_PATH)
    except Exception as e:
        log.warning(f"Failed to save config: {e}")
        try:
            os.unlink(tmp)
        except Exception:
            pass


cfg = load_config()

# Model constants
if IS_MAC:
    MODEL_SIZE = "mlx-community/whisper-large-v3-turbo"
    DEVICE = "mlx"
    COMPUTE_TYPE = None
else:
    MODEL_SIZE = "deepdml/faster-whisper-large-v3-turbo-ct2"
    DEVICE = "cuda"
    COMPUTE_TYPE = "float16"
