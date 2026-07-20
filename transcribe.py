"""
Whisper transcription engine with word-level timestamps.
Forked from Bark's transcriber.py, adapted for file-based captioning.
"""
import logging
import os
import re
from dataclasses import dataclass

from config import cfg
import config as _config

log = logging.getLogger(__name__)

# Filler words to strip (Swedish + English)
# Note: "typ" removed -- it is a real Swedish word ("type/kind of")
# Note: "you know" removed -- fullmatch on single words can never match it
# Note: "oh" removed -- it carries meaning in English ("Oh no", "Oh really?")
FILLER_WORDS = re.compile(
    r"(um|uh|uhm|hmm|ah|eh|"
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

# Only treat a HALLUCINATIONS match as fake when word confidence is ALSO low.
# A genuinely spoken closing "Thank you." carries high word probabilities;
# silence-hallucinated text does not. Without this gate, real speech that
# happens to match the list is silently deleted -- and on Mac the VAD
# pre-pass already removes the silence that breeds hallucinations, so a
# surviving match is more likely real than fake.
_HALLUCINATION_MAX_AVG_PROB = 0.5


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


def is_hallucination(text: str, avg_word_probability: float = 0.0) -> bool:
    """Check if a full segment is a known hallucination.

    Matches the known-phrase list AND requires low word confidence
    (avg_word_probability <= _HALLUCINATION_MAX_AVG_PROB). Callers that have
    no word probabilities can pass 0.0 to keep the old text-only behavior.
    """
    lower = text.lower().strip().strip(".,!?;:…")
    if lower not in HALLUCINATIONS:
        return False
    return avg_word_probability <= _HALLUCINATION_MAX_AVG_PROB


def _avg_probability(probs) -> float:
    """Mean of an iterable of probabilities; 0.0 when empty/unavailable."""
    probs = [p for p in probs if p is not None]
    if not probs:
        return 0.0
    return sum(probs) / len(probs)


class Transcriber:
    def __init__(self):
        self._vad = None
        if _config.IS_MAC:
            import mlx_whisper
            self._mlx = mlx_whisper
            self.model = None
            log.info(f"Loading model '{_config.MODEL_SIZE}' with MLX (Metal)...")
            # Warm up: forces model load + Metal kernel compilation at __init__
            # so the first real transcribe() call doesn't pay the cost. Also
            # surfaces broken downloads / configs before we render audio.
            try:
                import numpy as np
                self._mlx.transcribe(
                    np.zeros(16000, dtype=np.float32),
                    path_or_hf_repo=_config.MODEL_SIZE,
                )
            except Exception as e:
                log.error(f"MLX model warm-up failed: {e}")
                raise

            # Load silero-vad (ONNX). Used to find speech regions before
            # transcription -- mlx_whisper's word timestamps drift after long
            # pauses, so feeding only continuous speech eliminates the
            # clustered/zero-duration timestamp bug entirely.
            try:
                from silero_vad import load_silero_vad
                self._vad = load_silero_vad(onnx=True)
                log.info("VAD (silero-vad ONNX) loaded.")
            except Exception as e:
                log.warning(f"VAD model load failed: {e}; falling back to whole-file transcription.")
                self._vad = None
        else:
            from faster_whisper import WhisperModel
            self._mlx = None
            # On Windows CUDA, make CTranslate2 find torch's bundled
            # cuBLAS/cuDNN DLLs before the model loads.
            if _config.DEVICE == "cuda":
                import platforminfo
                platforminfo.add_cuda_dll_dir()
            log.info(f"Loading model '{_config.MODEL_SIZE}' on {_config.DEVICE} ({_config.COMPUTE_TYPE})...")
            try:
                self.model = WhisperModel(
                    _config.MODEL_SIZE,
                    device=_config.DEVICE,
                    compute_type=_config.COMPUTE_TYPE,
                )
            except Exception as e:
                if _config.DEVICE == "cuda":
                    err_str = str(e).lower()
                    if "out of memory" in err_str or "oom" in err_str:
                        log.error("GPU out of memory. Close other GPU apps.")
                    else:
                        log.warning(f"CUDA init failed: {e}")
                    log.info("Falling back to CPU mode (slower but compatible)...")
                    try:
                        self.model = WhisperModel(
                            _config.MODEL_SIZE,
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
            if _config.IS_MAC:
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

            seg_avg_prob = _avg_probability(
                w.probability for w in (seg.words or [])
            )
            if is_hallucination(seg.text, seg_avg_prob):
                log.debug(f"Dropping hallucination (avg prob {seg_avg_prob:.2f}): {seg.text!r}")
                continue

            min_prob = float(cfg.get("min_word_probability", 0.0) or 0.0)
            words = []
            if seg.words:
                for w in seg.words:
                    cleaned = clean_word(w.word)
                    if not cleaned:
                        continue
                    if min_prob > 0 and w.probability < min_prob:
                        log.debug(f"Drop low-confidence word ({w.probability:.2f}): {cleaned!r}")
                        continue
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

    @staticmethod
    def _find_tool(name: str):
        """Locate ffmpeg/ffprobe even when Resolve launches us with a bare
        PATH (Homebrew/ffmpeg live outside the default). Delegates to the
        platform boundary."""
        import platforminfo
        return platforminfo.find_tool(name)

    @classmethod
    def _load_mono_16k_ffmpeg(cls, audio_path: str):
        """Decode any codec to mono 16kHz float32 via ffmpeg.

        Fallback for formats libsndfile can't open - notably the AAC/.mp4
        that Resolve 21's factory 'Audio Only' preset renders. ffmpeg is a
        hard dependency of mlx_whisper, so it's always present when the
        pipeline can run at all.
        """
        import subprocess
        import numpy as np

        ffmpeg = cls._find_tool("ffmpeg")
        if not ffmpeg:
            raise RuntimeError("ffmpeg not found for audio decode fallback")
        proc = subprocess.run(
            [ffmpeg, "-v", "error", "-i", audio_path,
             "-f", "s16le", "-ac", "1", "-ar", "16000", "-"],
            capture_output=True, timeout=600,
        )
        if proc.returncode != 0 or not proc.stdout:
            raise RuntimeError(
                f"ffmpeg decode failed: {proc.stderr.decode(errors='replace')[:200]}")
        pcm = np.frombuffer(proc.stdout, dtype=np.int16)
        return pcm.astype(np.float32) / 32768.0

    @classmethod
    def _load_mono_16k(cls, audio_path: str):
        """Read audio in 30-second blocks, downmix to mono, resample to 16kHz.

        Block-wise loading keeps peak memory bounded by one block at the
        source sample rate (~10MB at 48kHz stereo) instead of materialising
        the whole file twice (raw + resampled). Important for hour-long
        interviews on 16GB machines.

        Codecs libsndfile can't read (AAC etc.) fall back to an ffmpeg
        decode so the VAD pre-pass never silently drops out.
        """
        import soundfile as sf
        import numpy as np

        try:
            info = sf.info(audio_path)
        except Exception:
            log.info("libsndfile can't read this format; decoding via ffmpeg.")
            return cls._load_mono_16k_ffmpeg(audio_path)
        sr = info.samplerate
        target_sr = 16000
        block = sr * 30
        ratio = target_sr / float(sr) if sr != target_sr else 1.0

        chunks = []
        with sf.SoundFile(audio_path, "r") as f:
            while True:
                data = f.read(block, dtype="float32")
                if data.size == 0:
                    break
                if data.ndim > 1:
                    data = data.mean(axis=1)
                if sr != target_sr:
                    new_len = int(round(len(data) * ratio))
                    if new_len == 0:
                        # Sub-sample-length tail block: appending it raw would
                        # inject source-rate samples into a 16kHz stream and
                        # shift every later VAD timestamp. Drop it (<1ms).
                        continue
                    xp = np.arange(len(data), dtype=np.float32)
                    x = np.linspace(0, len(data), num=new_len, endpoint=False, dtype=np.float32)
                    data = np.interp(x, xp, data).astype(np.float32)
                chunks.append(data)

        if not chunks:
            return np.array([], dtype=np.float32)
        return np.concatenate(chunks)

    def get_audio_duration(self, audio_path: str) -> float:
        """Return audio duration in seconds, or 0.0 on failure.

        soundfile first (cheap header read), ffprobe as the fallback for
        codecs libsndfile can't open (AAC etc.) - a 0.0 here collapses the
        progress estimate downstream, so fight for a real answer.
        """
        try:
            import soundfile as sf
            info = sf.info(audio_path)
            if info.samplerate:
                return info.frames / float(info.samplerate)
        except Exception:
            pass
        try:
            import subprocess
            ffprobe = self._find_tool("ffprobe")
            if not ffprobe:
                return 0.0
            proc = subprocess.run(
                [ffprobe, "-v", "error", "-show_entries", "format=duration",
                 "-of", "default=noprint_wrappers=1:nokey=1", audio_path],
                capture_output=True, text=True, timeout=60,
            )
            return max(float(proc.stdout.strip()), 0.0)
        except Exception:
            return 0.0

    def _vad_clip_timestamps(self, audio_path: str):
        """Run silero-vad on the audio and return mlx_whisper clip_timestamps.

        Returns a flat list [s1, e1, s2, e2, ...] (seconds) covering only the
        speech regions. Returns the string "0" if VAD is unavailable, finds
        no speech, or anything fails -- mlx_whisper interprets that as
        "transcribe the whole file" so we degrade gracefully.
        """
        if self._vad is None:
            return "0"
        try:
            from silero_vad import get_speech_timestamps

            audio = self._load_mono_16k(audio_path)
            if audio.size == 0:
                log.warning("Empty audio buffer for VAD; using whole file.")
                return "0"

            speech = get_speech_timestamps(
                audio, self._vad,
                return_seconds=True,
                # 200ms padding so we don't clip word starts/ends. silero's
                # 30ms default is too tight for caption-grade word timing.
                speech_pad_ms=200,
                # Keep VAD greedy on speech; a missed region means a word
                # never gets transcribed at all, which is worse than including
                # a borderline-quiet region.
                threshold=0.4,
                min_silence_duration_ms=300,
            )

            if not speech:
                log.warning("VAD found no speech regions; using whole file.")
                return "0"

            total_speech = sum(s["end"] - s["start"] for s in speech)
            log.info(
                f"VAD: {len(speech)} speech regions, {total_speech:.1f}s "
                f"(of {len(audio)/16000:.1f}s audio)"
            )

            flat = []
            for s in speech:
                flat.append(round(float(s["start"]), 2))
                flat.append(round(float(s["end"]), 2))
            return flat
        except Exception as e:
            log.warning(f"VAD pre-pass failed ({e}); using whole file.")
            return "0"

    def _transcribe_mlx(self, audio_path: str, on_progress=None) -> list:
        # mlx_whisper handles decoding via ffmpeg when given a path string,
        # which covers MP4/MOV/WAV/etc. without us needing soundfile+resampy.
        clip_timestamps = self._vad_clip_timestamps(audio_path)
        kwargs = dict(
            path_or_hf_repo=_config.MODEL_SIZE,
            language=cfg["language"],
            word_timestamps=True,
            # Each VAD clip is a continuous speech burst, so we don't want
            # the decoder carrying state across clips -- it can hallucinate
            # filler/repeats. Whole-file fallback also benefits from this
            # because long pauses are exactly where condition_on_previous_text
            # introduces drift.
            condition_on_previous_text=False,
            clip_timestamps=clip_timestamps,
        )
        # initial_prompt is supported by mlx_whisper too; beam_size is not
        # (mlx decodes greedily), so that config key is Windows-only.
        if cfg.get("initial_prompt"):
            kwargs["initial_prompt"] = cfg["initial_prompt"]
        result = self._mlx.transcribe(audio_path, **kwargs)

        if cfg["language"] is None and result.get("language"):
            log.info(f"Detected language: {result['language']}")

        # mlx_whisper.transcribe is synchronous -- no per-segment progress.
        segments = []
        raw_segs = result.get("segments", [])
        for seg in raw_segs:
            seg_text = seg.get("text", "").strip()
            raw_words = seg.get("words", []) or []
            seg_avg_prob = _avg_probability(w.get("probability") for w in raw_words)
            if is_hallucination(seg_text, seg_avg_prob):
                log.debug(f"Dropping hallucination (avg prob {seg_avg_prob:.2f}): {seg_text!r}")
                continue

            min_prob = float(cfg.get("min_word_probability", 0.0) or 0.0)
            words = []
            for w in raw_words:
                cleaned = clean_word(w.get("word", "") or w.get("text", ""))
                if not cleaned:
                    continue
                # No `or 1.0` here: a genuine 0.0 probability must stay 0.0,
                # only a MISSING key defaults to 1.0 (assume good faith).
                prob = w.get("probability")
                prob = 1.0 if prob is None else float(prob)
                if min_prob > 0 and prob < min_prob:
                    log.debug(f"Drop low-confidence word ({prob:.2f}): {cleaned!r}")
                    continue
                start = w.get("start")
                end = w.get("end")
                if start is None or end is None:
                    # A word with no timing would land at t=0 and corrupt
                    # ordering. Malformed and rare: drop it, loudly.
                    log.warning(f"Word missing timestamps, dropping: {cleaned!r}")
                    continue
                words.append(Word(
                    text=cleaned,
                    start=float(start),
                    end=float(end),
                    probability=prob,
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
