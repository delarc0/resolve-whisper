"""
SRT subtitle file generation from word-level timestamps.
Groups words into readable caption blocks with proper timing.
"""
import json
import logging
from config import cfg

log = logging.getLogger(__name__)


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

    # Group words into caption blocks
    captions = []
    block_words = []
    block_text = ""

    for word in all_words:
        w_text = word.text

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

        # Natural break: big pause in speech (>0.5s) = always split
        natural_break = block_words and pause > 0.5

        # Punctuation break: only trust it when there's also a pause (>0.2s)
        # Whisper often puts periods where commas belong, so punctuation
        # alone isn't reliable enough to split on
        if block_words and pause > 0.2:
            prev = block_words[-1].text
            punct_break = prev.rstrip().endswith((".", "!", "?", ":", ";"))
        else:
            punct_break = False

        # Flush current block if needed
        should_flush = block_words and (
            too_many_words or too_many_lines or too_long or
            natural_break or
            punct_break
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


def words_to_srt(segments: list, fps: float = 24.0) -> str:
    """Convert transcription segments into SRT subtitle format."""
    captions = words_to_captions(segments, fps)
    if not captions:
        return ""

    max_chars = cfg["max_chars_per_line"]
    max_lines = cfg["max_lines"]

    srt_lines = []
    for i, cap in enumerate(captions, 1):
        lines = _split_into_lines(cap["text"], max_chars)
        display = lines[:max_lines]

        srt_lines.append(str(i))
        srt_lines.append(
            f"{_format_timestamp(cap['start'])} --> {_format_timestamp(cap['end'])}"
        )
        srt_lines.append("\n".join(display))
        srt_lines.append("")  # blank line separator

    return "\n".join(srt_lines)


def write_srt(segments: list, output_path: str, fps: float = 24.0):
    """Generate SRT and write to file."""
    content = words_to_srt(segments, fps)
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


def write_captions_json(segments: list, output_path: str, fps: float = 24.0):
    """Generate structured caption data and write to JSON file."""
    captions = words_to_captions(segments, fps)
    if not captions:
        log.warning("No captions generated - empty transcription.")
        return False

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(captions, f, ensure_ascii=False)

    log.info(f"Wrote {len(captions)} captions to {output_path}")
    return True
