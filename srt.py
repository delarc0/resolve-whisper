"""
SRT subtitle file generation from word-level timestamps.
Groups words into readable caption blocks with proper timing.
"""
import json
import logging
from config import cfg

log = logging.getLogger(__name__)

# Words that shouldn't END a caption -- ending here makes the line read as a
# fragment ("Thomas Sjögren i" / "Mästerkocken..."). When the chunker is
# about to break and the last word matches this set, it suppresses the break
# so the connector sticks to the next word. Hard punctuation still wins.
_NOBREAK_TRAILING = {
    # Swedish prepositions / particles
    "i", "av", "till", "för", "på", "om", "med", "från", "vid", "som",
    "över", "under", "mot", "genom", "mellan", "utan",
    # Conjunctions / discourse markers
    "att", "och", "eller", "men", "så", "när", "om", "innan", "efter",
    "fast", "trots", "ifall",
    # Articles / determiners
    "den", "det", "de", "en", "ett",
    # Auxiliaries
    "är", "var", "har", "hade", "ska", "skulle", "kan", "kunde", "får",
    "fick", "vill", "ville", "måste", "borde",
    # Pronouns
    "jag", "du", "han", "hon", "vi", "ni", "min", "din", "sin", "vår", "er",
    # English equivalents (for mixed/auto-detect)
    "a", "an", "the", "of", "to", "in", "on", "at", "for", "with", "by",
    "and", "or", "but", "so", "as", "is", "are", "was", "were",
}


def _strip_trailing_punct(s: str) -> str:
    return s.rstrip(".,!?;:…")


def _format_timestamp(seconds: float) -> str:
    """Convert seconds to SRT timestamp format: HH:MM:SS,mmm"""
    if seconds < 0:
        seconds = 0.0
    total_ms = round(seconds * 1000)
    h = total_ms // 3_600_000
    total_ms %= 3_600_000
    m = total_ms // 60_000
    total_ms %= 60_000
    s = total_ms // 1000
    ms = total_ms % 1000
    return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"


def _split_into_lines(text: str, max_chars: int) -> list:
    """Split text into lines respecting max character width."""
    words = text.split()
    lines = []
    current_line = ""

    for word in words:
        test = f"{current_line} {word}".strip()
        if len(test) <= max_chars:
            current_line = test
        else:
            if current_line:
                lines.append(current_line)
            current_line = word

    if current_line:
        lines.append(current_line)

    return lines


_SOLO_DURATION = 0.45   # word held this long = emphasized -> own caption
_SOLO_PAUSE = 0.25      # silence this long on both sides = isolated -> own caption


def _mark_emphasized(words: list) -> list:
    """Return a list[bool] of which words deserve solo captions.

    A word is "emphasized" when:
      - it lasts >= _SOLO_DURATION seconds (held / drawn out by the speaker)
      - OR it has >= _SOLO_PAUSE silence on both sides (isolated)
    """
    n = len(words)
    flags = [False] * n
    for i, w in enumerate(words):
        duration = w.end - w.start
        if duration >= _SOLO_DURATION:
            flags[i] = True
            continue
        pause_before = w.start - words[i - 1].end if i > 0 else _SOLO_PAUSE
        pause_after = words[i + 1].start - w.end if i < n - 1 else _SOLO_PAUSE
        if pause_before >= _SOLO_PAUSE and pause_after >= _SOLO_PAUSE:
            flags[i] = True
    return flags


def words_to_captions(segments: list, fps: float = 24.0) -> list:
    """
    Group transcription segments into caption blocks.

    Returns list of dicts: [{"start": float, "end": float, "text": str}, ...]
    Each dict represents one caption with timing in seconds.
    """
    max_words = cfg["max_words_per_caption"]
    max_chars = cfg["max_chars_per_line"]
    max_lines = cfg["max_lines"]
    min_dur = cfg["min_duration_s"]
    max_dur = cfg["max_duration_s"]
    gap_frames = cfg["gap_frames"]

    gap_s = gap_frames / fps if fps > 0 else gap_frames / 24.0

    # Flatten all words from all segments
    all_words = []
    for seg in segments:
        all_words.extend(seg.words)

    if not all_words:
        return []

    emphasized = _mark_emphasized(all_words)

    # Group words into caption blocks
    captions = []
    block_words = []
    block_text = ""

    for idx, word in enumerate(all_words):
        w_text = word.text
        is_emphasized = emphasized[idx]

        # Check if adding this word exceeds our limits
        test_text = f"{block_text} {w_text}".strip()
        test_lines = _split_into_lines(test_text, max_chars)

        # Check word count limit
        too_many_words = max_words > 0 and len(block_words) >= max_words

        # Check timing: would this block exceed max duration?
        if block_words:
            block_dur = word.end - block_words[0].start
            too_long = block_dur > max_dur
        else:
            too_long = False

        # Check if too many lines
        too_many_lines = len(test_lines) > max_lines

        # Check pause between this word and the previous one
        if block_words:
            pause = word.start - block_words[-1].end
        else:
            pause = 0.0

        # Tiered pause handling: a hard pause always breaks; a smaller pause
        # only breaks once the caption has been on screen long enough that
        # the viewer needs a refresh. This stops the chunker from greedily
        # filling to max_words on every line.
        if block_words:
            block_dur_so_far = block_words[-1].end - block_words[0].start
        else:
            block_dur_so_far = 0.0

        # >0.18s pause = clear breath pause -> always break
        hard_pause = block_words and pause > 0.18
        # any micro-pause (>0.06s) breaks if caption has been up for ~0.6s+
        micro_pause = block_words and pause > 0.06 and block_dur_so_far > 0.6
        natural_break = hard_pause or micro_pause

        # Sentence-ending punctuation always breaks (. ! ?). Soft punctuation
        # (, ; :) still needs a >0.2s pause since Whisper puts those unreliably.
        if block_words:
            prev = block_words[-1].text.rstrip()
            hard_punct = prev.endswith((".", "!", "?"))
            soft_punct = prev.endswith((",", ";", ":")) and pause > 0.2
            punct_break = hard_punct or soft_punct
        else:
            punct_break = False

        # Suppress soft breaks if the current block would end on a connector
        # word ("Thomas Sjögren i" / "alla tre av de" -> bad). Hard punctuation
        # and word/char/duration limits still force a break.
        if block_words:
            last_lc = _strip_trailing_punct(block_words[-1].text.lower())
            ends_on_connector = last_lc in _NOBREAK_TRAILING
        else:
            ends_on_connector = False

        # Flush current block if needed. Emphasized words always get their
        # own caption: flush whatever's pending, write the word as a solo
        # block, and reset.
        should_flush = block_words and (
            too_many_words or too_many_lines or too_long or
            punct_break or
            is_emphasized or
            (natural_break and not ends_on_connector)
        )

        if should_flush:
            captions.append({
                "start": block_words[0].start,
                "end": block_words[-1].end,
                "text": block_text,
            })
            block_words = []
            block_text = ""

        block_words.append(word)
        block_text = f"{block_text} {w_text}".strip() if block_text else w_text

        # If this word is emphasized, close the block right after adding it
        # so it stands alone on screen.
        if is_emphasized:
            captions.append({
                "start": block_words[0].start,
                "end": block_words[-1].end,
                "text": block_text,
            })
            block_words = []
            block_text = ""
            continue

    # Flush remaining
    if block_words:
        captions.append({
            "start": block_words[0].start,
            "end": block_words[-1].end,
            "text": block_text,
        })

    # Enforce minimum duration and add gaps
    for i, cap in enumerate(captions):
        dur = cap["end"] - cap["start"]
        if dur < min_dur:
            cap["end"] = cap["start"] + min_dur

        # Ensure end never overlaps with next caption's start
        if i < len(captions) - 1:
            next_start = captions[i + 1]["start"]
            if cap["end"] + gap_s > next_start:
                cap["end"] = max(cap["start"] + 0.1, next_start - gap_s)
            # Hard clamp: never exceed next caption's start
            if cap["end"] > next_start:
                cap["end"] = next_start

    return captions


def strip_punct_text(text: str) -> str:
    import re
    return re.sub(r" +", " ", re.sub(r"[^\w\s]", "", text, flags=re.UNICODE)).strip()


def words_to_srt(segments: list, fps: float = 24.0, strip_punctuation: bool = False) -> str:
    """Convert transcription segments into SRT subtitle format."""
    captions = words_to_captions(segments, fps)
    if not captions:
        return ""

    max_chars = cfg["max_chars_per_line"]
    max_lines = cfg["max_lines"]

    srt_lines = []
    for i, cap in enumerate(captions, 1):
        text = strip_punct_text(cap["text"]) if strip_punctuation else cap["text"]
        lines = _split_into_lines(text, max_chars)
        display = lines[:max_lines]

        srt_lines.append(str(i))
        srt_lines.append(
            f"{_format_timestamp(cap['start'])} --> {_format_timestamp(cap['end'])}"
        )
        srt_lines.append("\n".join(display))
        srt_lines.append("")  # blank line separator

    return "\n".join(srt_lines)


def write_srt(segments: list, output_path: str, fps: float = 24.0, strip_punctuation: bool = False):
    """Generate SRT and write to file."""
    content = words_to_srt(segments, fps, strip_punctuation=strip_punctuation)
    if not content:
        log.warning("No captions generated - empty transcription.")
        return False

    with open(output_path, "w", encoding="utf-8-sig") as f:
        f.write(content)
        if not content.endswith("\n"):
            f.write("\n")

    # Count captions (count SRT sequence numbers: lines that are just a number)
    count = sum(1 for line in content.split("\n") if line.strip().isdigit())
    log.info(f"Wrote {count} captions to {output_path}")
    return True


def write_captions_json(segments: list, output_path: str, fps: float = 24.0, strip_punctuation: bool = False):
    """Generate structured caption data and write to JSON file."""
    captions = words_to_captions(segments, fps)
    if not captions:
        log.warning("No captions generated - empty transcription.")
        return False

    if strip_punctuation:
        for cap in captions:
            cap["text"] = strip_punct_text(cap["text"])

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(captions, f, ensure_ascii=False)

    log.info(f"Wrote {len(captions)} captions to {output_path}")
    return True
