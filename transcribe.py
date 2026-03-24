"""
Whisper transcription engine with word-level timestamps.
Forked from Bark's transcriber.py, adapted for file-based captioning.
"""
import logging
import re
from dataclasses import dataclass

from config import cfg, MODEL_SIZE, DEVICE, COMPUTE_TYPE, IS_MAC

log = logging.getLogger(__name__)

# Filler words to strip (Swedish + English)
# Note: "typ" removed -- it is a real Swedish word ("type/kind of")
# Note: "you know" removed -- fullmatch on single words can never match it
FILLER_WORDS = re.compile(
    r"(um|uh|uhm|hmm|ah|eh|oh|"
    r"liksom|asså|alltså|öh|äh)",
    re.IGNORECASE,
)

# Whisper hallucination patterns (common on silence/noise)
HALLUCINATIONS = {
    "thank you",
    "thanks for watching",
    "subscribe",
    "tack för att ni tittade",
    "tack för att du tittade",
    "undertextning",
}


@dataclass
class Word:
    text: str
    start: float   # seconds
    end: float     # seconds
    probability: float


@dataclass
class Segment:
    text: str
    start: float
    end: float
    words: list  # list of Word


def clean_word(text: str) -> str:
    """Clean a single word - remove filler words and Whisper annotations."""
    stripped = text.strip()
    if not stripped:
        return ""
    # Strip Whisper annotations like [Music], [Applause], [inaudible]
    if stripped.startswith("[") and stripped.endswith("]"):
        return ""
    # Remove stray brackets/question-mark artifacts
    stripped = stripped.strip("[]")
    if not stripped:
        return ""
    lower = stripped.lower().strip(".,!?;:")
    if FILLER_WORDS.fullmatch(lower):
        return ""
    return stripped


def is_hallucination(text: str) -> bool:
    """Check if a full segment is a known hallucination."""
    lower = text.lower().strip().strip(".,!?;:…")
    return lower in HALLUCINATIONS


class Transcriber:
    def __init__(self):
        if IS_MAC:
            import mlx_whisper
            self._mlx = mlx_whisper
            self.model = None
            log.info(f"Loading model '{MODEL_SIZE}' with MLX (Metal)...")
            # Warm up with 1s silence
            try:
                import numpy as np
                self._mlx.transcribe(
                    np.zeros(16000, dtype=np.float32),
                    path_or_hf_repo=MODEL_SIZE,
                )
            except Exception as e:
                log.error(f"MLX model warm-up failed: {e}")
                raise
        else:
            from faster_whisper import WhisperModel
            self._mlx = None
            log.info(f"Loading model '{MODEL_SIZE}' on {DEVICE} ({COMPUTE_TYPE})...")
            try:
                self.model = WhisperModel(
                    MODEL_SIZE,
                    device=DEVICE,
                    compute_type=COMPUTE_TYPE,
                )
            except Exception as e:
                if DEVICE == "cuda":
                    err_str = str(e).lower()
                    if "out of memory" in err_str or "oom" in err_str:
                        log.error("GPU out of memory. Close other GPU apps.")
                    else:
                        log.warning(f"CUDA init failed: {e}")
                    log.info("Falling back to CPU mode (slower but compatible)...")
                    try:
                        self.model = WhisperModel(
                            MODEL_SIZE,
                            device="cpu",
                            compute_type="int8",
                        )
                    except Exception as cpu_err:
                        log.error(f"CPU fallback also failed: {cpu_err}")
                        raise
                else:
                    raise
        log.info("Model loaded.")

    def transcribe(self, audio_path: str, on_progress=None) -> list:
        """
        Transcribe an audio file and return segments with word-level timestamps.

        Args:
            audio_path: Path to WAV/audio file
            on_progress: Optional callback(percent: int) called as segments complete

        Returns:
            List of Segment objects, each containing Word objects with timing.
        """
        try:
            if IS_MAC:
                return self._transcribe_mlx(audio_path, on_progress)
            else:
                return self._transcribe_faster_whisper(audio_path, on_progress)
        except Exception as e:
            log.error(f"Transcription failed: {e}")
            return []

    def _transcribe_faster_whisper(self, audio_path: str, on_progress=None) -> list:
        kwargs = dict(
            beam_size=cfg["beam_size"],
            language=cfg["language"],
            word_timestamps=True,
            condition_on_previous_text=True,
            patience=2.0,
            hallucination_silence_threshold=2.0,
            vad_filter=True,
            vad_parameters=dict(
                min_silence_duration_ms=500,
                speech_pad_ms=400,
            ),
        )

        # Optional: initial prompt for domain/style hints
        if cfg.get("initial_prompt"):
            kwargs["initial_prompt"] = cfg["initial_prompt"]

        raw_segments, info = self.model.transcribe(audio_path, **kwargs)

        duration = info.duration if info else 0.0

        if cfg["language"] is None and info:
            log.info(f"Detected language: {info.language} ({info.language_probability:.0%})")

        segments = []
        last_pct = -1
        for seg in raw_segments:
            # Report progress based on segment position vs total duration
            if on_progress and duration > 0:
                try:
                    pct = min(int(seg.end / duration * 100), 99)
                    if pct > last_pct:
                        last_pct = pct
                        on_progress(pct)
                except Exception:
                    pass

            if is_hallucination(seg.text):
                continue

            words = []
            if seg.words:
                for w in seg.words:
                    cleaned = clean_word(w.word)
                    if cleaned:
                        words.append(Word(
                            text=cleaned,
                            start=w.start,
                            end=w.end,
                            probability=w.probability,
                        ))

            if words:
                text = " ".join(w.text for w in words)
                text = re.sub(r"  +", " ", text).strip()
                segments.append(Segment(
                    text=text,
                    start=words[0].start,
                    end=words[-1].end,
                    words=words,
                ))

        try:
            if on_progress:
                on_progress(100)
        except Exception:
            pass
        return segments

    def _transcribe_mlx(self, audio_path: str, on_progress=None) -> list:
        import numpy as np
        import soundfile as sf

        audio, sr = sf.read(audio_path, dtype="float32")
        if len(audio.shape) > 1:
            audio = audio.mean(axis=1)
        if sr != 16000:
            # Resample to 16kHz
            import resampy
            audio = resampy.resample(audio, sr, 16000)

        result = self._mlx.transcribe(
            audio,
            path_or_hf_repo=MODEL_SIZE,
            language=cfg["language"],
            word_timestamps=True,
        )

        if cfg["language"] is None and result.get("language"):
            log.info(f"Detected language: {result['language']}")

        duration = len(audio) / 16000.0  # audio is already 16kHz at this point

        segments = []
        raw_segs = result.get("segments", [])
        last_pct = -1
        for seg in raw_segs:
            seg_end = seg.get("end", 0.0)
            if on_progress and duration > 0:
                try:
                    pct = min(int(seg_end / duration * 100), 99)
                    if pct > last_pct:
                        last_pct = pct
                        on_progress(pct)
                except Exception:
                    pass

            seg_text = seg.get("text", "").strip()
            if is_hallucination(seg_text):
                continue

            words = []
            for w in seg.get("words", []):
                cleaned = clean_word(w.get("word", "") or w.get("text", ""))
                if cleaned:
                    words.append(Word(
                        text=cleaned,
                        start=w.get("start", 0.0),
                        end=w.get("end", 0.0),
                        probability=w.get("probability", 1.0),
                    ))

            if words:
                text = " ".join(w.text for w in words)
                text = re.sub(r"  +", " ", text).strip()
                segments.append(Segment(
                    text=text,
                    start=words[0].start,
                    end=words[-1].end,
                    words=words,
                ))

        try:
            if on_progress:
                on_progress(100)
        except Exception:
            pass
        return segments
