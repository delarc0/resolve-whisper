"""
SRT subtitle file generation from word-level timestamps.
Groups words into readable caption blocks with proper timing.
"""
import logging
from config import cfg

log = logging.getLogger(__name__)


def _format_timestamp(seconds: float) -> str:
    """Convert seconds to SRT timestamp format: HH:MM:SS,mmm"""
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = int(seconds % 60)
    ms = int(round((seconds % 1) * 1000))
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


def words_to_srt(segments: list, fps: float = 24.0) -> str:
    """
    Convert transcription segments into SRT subtitle format.

    Groups words into caption blocks based on:
    - Max characters per line
    - Max lines per caption
    - Min/max duration
    - Natural pause points (punctuation, timing gaps)

    Args:
        segments: List of Segment objects from transcribe.py
        fps: Timeline frame rate (for gap calculation)

    Returns:
        SRT file content as string
    """
    max_chars = cfg["max_chars_per_line"]
    max_lines = cfg["max_lines"]
    min_dur = cfg["min_duration_s"]
    max_dur = cfg["max_duration_s"]
    gap_frames = cfg["gap_frames"]
    confidence_threshold = cfg["confidence_threshold"]
    highlight = cfg["highlight_low_confidence"]

    gap_s = gap_frames / fps

    # Flatten all words from all segments
    all_words = []
    for seg in segments:
        all_words.extend(seg.words)

    if not all_words:
        return ""

    # Group words into caption blocks
    captions = []
    block_words = []
    block_text = ""

    for word in all_words:
        w_text = word.text

        # Mark low-confidence words
        if highlight and word.probability < confidence_threshold:
            w_text = f"[{w_text}?]"

        # Check if adding this word exceeds our limits
        test_text = f"{block_text} {w_text}".strip()
        test_lines = _split_into_lines(test_text, max_chars)

        # Check timing: would this block exceed max duration?
        if block_words:
            block_dur = word.end - block_words[0].start
            too_long = block_dur > max_dur
        else:
            too_long = False

        # Check if too many lines
        too_many_lines = len(test_lines) > max_lines

        # Check for natural break: big pause between words (>0.5s)
        if block_words:
            pause = word.start - block_words[-1].end
            natural_break = pause > 0.5
        else:
            natural_break = False

        # Check for punctuation break at end of previous word
        if block_words:
            prev = block_words[-1].text
            punct_break = prev.rstrip().endswith((".", "!", "?", ":", ";"))
        else:
            punct_break = False

        # Flush current block if needed
        should_flush = block_words and (
            too_many_lines or too_long or
            (natural_break and len(block_text) > 20) or
            (punct_break and len(block_text) > 20)
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

        # Add gap before next caption
        if i < len(captions) - 1:
            next_start = captions[i + 1]["start"]
            if cap["end"] + gap_s > next_start:
                cap["end"] = max(cap["start"] + 0.1, next_start - gap_s)

    # Format as SRT
    srt_lines = []
    for i, cap in enumerate(captions, 1):
        lines = _split_into_lines(cap["text"], max_chars)
        # Limit to max_lines
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

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(content)

    # Count captions
    count = content.count("\n\n")
    log.info(f"Wrote {count} captions to {output_path}")
    return True
