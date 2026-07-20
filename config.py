import json
import logging
import os

from platforminfo import IS_MAC, IS_WIN, IS_LINUX  # noqa: F401

log = logging.getLogger(__name__)

_APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(_APP_DIR, "caption_config.json")

DEFAULT_CONFIG = {
    # Transcription
    "language": "sv",           # None = auto-detect, "sv" = Swedish, "en" = English
    "beam_size": 10,            # Windows/faster-whisper only; mlx (Mac) decodes greedily
    "initial_prompt": None,     # Context hint, e.g. "Intervju om ledarskap och projektledning."
    # Drop words below this Whisper confidence. 0.0 = off (default): Whisper
    # gives real-but-mumbled words low scores, and deleting a mid-sentence
    # word silently corrupts the caption. Set ~0.3 only on noisy footage
    # where junk words actually appear.
    "min_word_probability": 0.0,
    # SRT formatting
    "max_words_per_caption": 0,      # 0 = no limit (use chars/lines), >0 = hard word cap
    "max_chars_per_line": 42,        # 42 = broadcast standard (EBU/Netflix) for 16:9
    "max_lines": 1,                  # 1 = single-line captions (clean for podcast/interview)
    "min_duration_s": 1.0,
    "max_duration_s": 7.0,
    "gap_frames": 2,            # gap between subtitles in frames
    # ALL CAPS caption text. Resolve's subtitle track style has no caps
    # transform (and no styling API at all), so caps live in the text.
    "uppercase": True,
    # Output
    "output_dir": None,         # None = same dir as source / project
    "keep_srt_days": 30,        # auto-delete SRTs older than this in output dir; 0 disables
}


def _write_default_config():
    """Write DEFAULT_CONFIG to CONFIG_PATH so users can discover the knobs.

    Only writes if CONFIG_PATH doesn't exist; never overwrites a user-edited
    file. Failures are non-fatal -- defaults still apply in-memory.
    """
    if os.path.exists(CONFIG_PATH):
        return
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, indent=2, ensure_ascii=False)
            f.write("\n")
    except OSError as e:
        log.debug(f"Could not write default config: {e}")


def load_config() -> dict:
    _write_default_config()
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
    MODEL_SIZE = "mlx-community/whisper-large-v3-mlx"
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
