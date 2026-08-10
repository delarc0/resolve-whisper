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
    # Use KB-Whisper (KBLab Swedish fine-tune, ~47% lower WER on Swedish) when
    # the language is explicitly "sv". Auto/other languages always use large-v3.
    "use_kb_whisper": True,
    # Path to a converted KB-Whisper MLX model (Mac). null = the default
    # ~/.cache/lab37-whisper/kb-whisper-large-mlx if it exists.
    "kb_whisper_mlx_path": None,
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


# Keys whose default is None but which accept a string when set.
_NULLABLE_STR_KEYS = {"language", "initial_prompt", "output_dir",
                      "kb_whisper_mlx_path"}

# Numeric keys that must stay positive / non-negative to make any sense.
_POSITIVE_KEYS = {"max_chars_per_line", "max_lines", "min_duration_s",
                  "max_duration_s", "beam_size"}
_NON_NEGATIVE_KEYS = {"max_words_per_caption", "gap_frames", "keep_srt_days",
                      "min_word_probability"}


def _coerce(key, value):
    """Return value if it is usable for `key`, else None to reject it.

    A bad type used to reach the chunker and raise mid-run -- AFTER the
    timeline audio had been rendered -- so the user lost the render and got
    a traceback. The generated config file itself teaches `null` as the
    "unset" idiom, so a user copying that pattern onto a numeric key is a
    realistic mistake, not an exotic one.
    """
    import math
    default = DEFAULT_CONFIG[key]
    if key in _NULLABLE_STR_KEYS:
        return value if value is None or isinstance(value, str) else None
    if isinstance(default, bool):
        # bool first: bool is a subclass of int.
        return value if isinstance(value, bool) else None

    if isinstance(default, (int, float)) and not isinstance(default, bool):
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            return None
        # json.load happily produces inf/nan from 1e400 / Infinity / NaN, and
        # isinstance(inf, float) is True -- so a type check alone let them
        # through to crash at SRT-write time, i.e. AFTER the render.
        if not math.isfinite(value):
            return None
        # Zero/negative are type-valid but semantically broken (a negative
        # max_chars_per_line silently yields one word per caption).
        if key in _NON_NEGATIVE_KEYS and value < 0:
            return None
        if key in _POSITIVE_KEYS and value <= 0:
            return None
        if isinstance(default, int):
            if isinstance(value, int):
                return value
            return int(value) if float(value).is_integer() else None
        return float(value)

    if isinstance(default, str):
        return value if isinstance(value, str) else None
    return value


def load_config() -> dict:
    _write_default_config()
    config = dict(DEFAULT_CONFIG)
    if not os.path.exists(CONFIG_PATH):
        return config

    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            saved = json.load(f)
    except Exception as e:
        # Name the file and the reason: silently reverting every setting
        # reads to the user as "my settings do nothing".
        log.warning(f"Could not read {CONFIG_PATH} ({e}); using defaults for "
                    "ALL settings. Fix the JSON or delete the file to "
                    "regenerate it.")
        return config

    if not isinstance(saved, dict):
        log.warning(f"{CONFIG_PATH} is not a JSON object; using defaults.")
        return config

    for key, value in saved.items():
        if key not in DEFAULT_CONFIG:
            log.debug(f"Ignoring unknown config key '{key}'.")
            continue
        coerced = _coerce(key, value)
        # None is only acceptable where None is a meaningful value
        # ("language": null = auto-detect). An explicit null on a numeric
        # key is the exact case that used to crash mid-run.
        if coerced is None and not (value is None and key in _NULLABLE_STR_KEYS):
            log.warning(f"Config '{key}': {value!r} is not a valid "
                        f"{type(DEFAULT_CONFIG[key]).__name__}; "
                        f"using default {DEFAULT_CONFIG[key]!r}.")
            continue
        config[key] = coerced
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

# --- Swedish: KB-Whisper (KBLab, National Library of Sweden) ---
# A large-v3 fine-tune on ~50k h Swedish, ~47% lower WER on Swedish than
# openai large-v3. Used ONLY when the language is explicitly Swedish; auto
# and other languages stay on large-v3 (KB-Whisper is Swedish-specialized
# and degrades other languages).
#
# Windows/Linux (faster-whisper) load the CTranslate2 weights straight from
# the repo. Mac (mlx) has no official MLX build, so KB_WHISPER_MLX names our
# conversion; if it's blank or fails to load, Swedish falls back to large-v3
# (config below stays valid, just without the accuracy bump).
KB_WHISPER_CT2 = "KBLab/kb-whisper-large"

# Mac: KBLab publishes no MLX build, so the model has to be converted once
# (tools/convert_kb_whisper_mlx.sh). Auto-detected here so the conversion
# activates with no code edit; overridable with "kb_whisper_mlx_path" in
# caption_config.json for a shared/network copy.
KB_WHISPER_MLX_DEFAULT = os.path.expanduser(
    "~/.cache/lab37-whisper/kb-whisper-large-mlx")


def _kb_whisper_mlx() -> str:
    path = cfg.get("kb_whisper_mlx_path") or KB_WHISPER_MLX_DEFAULT
    path = os.path.expanduser(str(path))
    # A directory alone isn't enough: a half-finished conversion would make
    # every Swedish run fail instead of falling back.
    if os.path.isdir(path) and (
        os.path.exists(os.path.join(path, "weights.safetensors"))
        or os.path.exists(os.path.join(path, "weights.npz"))
    ):
        return path
    return ""


def model_for_language(lang):
    """Transcription model id for a language code (None/'' = auto-detect).

    Returns the Swedish KB-Whisper model when the user asserted Swedish and
    the knob is on, else the default large-v3 model for this platform.
    """
    if lang and str(lang).lower() == "sv" and cfg.get("use_kb_whisper", True):
        kb = _kb_whisper_mlx() if IS_MAC else KB_WHISPER_CT2
        if kb:
            return kb
    return MODEL_SIZE
