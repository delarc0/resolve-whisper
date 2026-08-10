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


# Abbreviations ending in '.' that must NOT trigger a sentence break.
# Lowercased, exact match on the full word.
_ABBREVIATIONS = {
    # Swedish
    "t.ex.", "bl.a.", "m.m.", "osv.", "dvs.", "d.v.s.", "p.g.a.", "s.k.",
    "fr.o.m.", "t.o.m.", "ca.", "st.", "kr.", "nr.", "tel.", "kl.",
    # English
    "e.g.", "i.e.", "etc.", "vs.", "mr.", "mrs.", "dr.", "st.",
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


class _TimedWord:
    """Lightweight stand-in used when word timings need rewriting.

    The chunkers only ever read .text/.start/.end, so this is enough and
    avoids mutating the caller's Whisper objects.
    """
    __slots__ = ("text", "start", "end", "probability")

    def __init__(self, text, start, end, probability=1.0):
        self.text = text
        self.start = start
        self.end = end
        self.probability = probability


def _normalise_word_times(words: list) -> list:
    """Force word stamps to be monotonic and non-inverted.

    Whisper (both backends) occasionally returns a word that starts before
    the previous one ended, or start == end. Left alone these produce
    overlapping cues, zero-duration cues, and -- when the finished captions
    were sorted to compensate -- a reordered transcript.
    """
    import math
    out = []
    prev_end = None
    dropped = 0
    for w in words:
        try:
            start = float(w.start)
            end = float(w.end)
        except (TypeError, ValueError):
            dropped += 1
            continue
        # NaN/inf would survive the float() call and then blow up in the
        # millisecond rounding further down.
        if not (math.isfinite(start) and math.isfinite(end)):
            dropped += 1
            continue
        if prev_end is not None and start < prev_end:
            start = prev_end
        if end < start:
            end = start
        out.append(_TimedWord(w.text, start, end,
                              getattr(w, "probability", 1.0)))
        prev_end = end
    if dropped:
        # Never drop words silently: that is the failure mode this module
        # exists to prevent.
        log.warning(f"Dropped {dropped} word(s) with unusable timestamps.")
    return out


def _merge_zero_duration(captions: list) -> list:
    """Remove cues that would still write as zero-duration.

    Only reachable when several words share one timestamp so tightly that
    the clamps cannot separate them at millisecond resolution. Merging the
    text into a neighbour keeps every word instead of emitting a cue that
    strict SRT parsers reject.
    """
    if len(captions) < 2:
        return captions
    out = []
    for i, cap in enumerate(captions):
        if round(cap["end"] * 1000) > round(cap["start"] * 1000):
            out.append(cap)
            continue
        if out:
            out[-1]["text"] = f"{out[-1]['text']} {cap['text']}".strip()
        elif i + 1 < len(captions):
            # No previous cue to absorb it: hand the text to the next one.
            nxt = captions[i + 1]
            nxt["text"] = f"{cap['text']} {nxt['text']}".strip()
        else:
            out.append(cap)
    return out or captions


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


# --- Balanced chunking (plain-SRT mode: Auto/Podcast, max_words == 0) -------
#
# The greedy chunker breaks the moment a rule fires, which structurally
# produces a packed caption followed by a short orphan at the end of a
# sentence ("...full 42-char line" + "och nämns ofta"). Plain mode instead
# splits the words at hard boundaries (sentence punctuation, real silence)
# and then picks break points inside each stretch by minimizing a cost over
# all feasible partitions:
#   - captions should be evenly filled (squared slack vs capacity)
#   - breaking after a connector word costs extra
#   - breaking on a pause or after soft punctuation is rewarded
# Char/line/duration limits stay hard constraints.

_STRETCH_PAUSE = 0.6          # silence >= this always separates captions
_CONNECTOR_PENALTY = 1000.0   # breaking after "och"/"i"/"att"/...
_SOFT_PUNCT_BONUS = 300.0     # breaking after , ; :
_PAUSE_BONUS_MAX = 600.0      # breaking on a pause, scaled up to _STRETCH_PAUSE


def _split_stretches(words: list) -> list:
    """Split words at hard boundaries: sentence punctuation or real silence."""
    stretches = []
    current = []
    for i, w in enumerate(words):
        current.append(w)
        text = w.text.rstrip()
        hard_punct = (
            text.endswith((".", "!", "?"))
            and text.lower() not in _ABBREVIATIONS
        )
        pause = words[i + 1].start - w.end if i + 1 < len(words) else 0.0
        if hard_punct or pause >= _STRETCH_PAUSE:
            stretches.append(current)
            current = []
    if current:
        stretches.append(current)
    return stretches


def _break_cost(prev, cur) -> float:
    """Cost of placing a caption break between two adjacent words."""
    cost = 0.0
    if _strip_trailing_punct(prev.text.lower()) in _NOBREAK_TRAILING:
        cost += _CONNECTOR_PENALTY
    if prev.text.rstrip().endswith((",", ";", ":")):
        cost -= _SOFT_PUNCT_BONUS
    pause = max(0.0, cur.start - prev.end)
    cost -= min(pause, _STRETCH_PAUSE) / _STRETCH_PAUSE * _PAUSE_BONUS_MAX
    return cost


def _partition_stretch(words: list, max_chars: int, max_lines: int,
                       max_dur: float) -> list:
    """Find the minimum-cost partition of one stretch into captions."""
    n = len(words)
    capacity = max_chars * max_lines
    # Prefix sums of word lengths for O(1) chunk-text length.
    plen = [0] * (n + 1)
    for i, w in enumerate(words):
        plen[i + 1] = plen[i] + len(w.text)

    def chunk_len(j, i):
        return plen[i] - plen[j] + (i - j - 1)

    def chunk_fits(j, i):
        if i - j == 1:
            return True  # a single word can't be split further
        if words[i - 1].end - words[j].start > max_dur:
            return False
        text = " ".join(w.text for w in words[j:i])
        lines = _split_into_lines(text, max_chars)
        return len(lines) <= max_lines and all(len(l) <= max_chars for l in lines)

    best = [float("inf")] * (n + 1)
    back = [0] * (n + 1)
    best[0] = 0.0
    for i in range(1, n + 1):
        # Grow the candidate chunk backwards; once it stops fitting, longer
        # chunks can't fit either (length and duration are monotonic).
        for j in range(i - 1, -1, -1):
            if i - j > 1 and not chunk_fits(j, i):
                break
            slack = max(0, capacity - chunk_len(j, i))
            cost = best[j] + slack * slack
            if j > 0:
                cost += _break_cost(words[j - 1], words[j])
            if cost < best[i]:
                best[i] = cost
                back[i] = j

    chunks = []
    i = n
    while i > 0:
        j = back[i]
        chunks.append(words[j:i])
        i = j
    chunks.reverse()
    return [
        {
            "start": c[0].start,
            "end": c[-1].end,
            "text": " ".join(w.text for w in c),
        }
        for c in chunks
    ]


def _balanced_captions(all_words: list, max_chars: int, max_lines: int,
                       max_dur: float) -> list:
    captions = []
    for stretch in _split_stretches(all_words):
        captions.extend(_partition_stretch(stretch, max_chars, max_lines, max_dur))
    return captions


def words_to_captions(segments: list, fps: float = 24.0,
                      offset_s: float = 0.0) -> list:
    """
    Group transcription segments into caption blocks.

    offset_s shifts every caption later by that many seconds. Used when the
    audio was rendered from an in/out range: Whisper timestamps are zero-based
    at the IN point, but Resolve maps an imported SRT's 00:00:00 to the
    timeline start, so without the shift captions land at the top of the
    timeline instead of over the speech.

    Reels mode (max_words > 0) chunks greedily: short groups timed tightly to
    speech, emphasized words solo. Plain-SRT mode (max_words == 0) uses
    balanced segmentation so sentences split at the globally best points
    instead of packing greedily until the char limit hits.

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

    # Normalise word timing BEFORE chunking. Whisper occasionally emits
    # duplicated or inverted stamps; sorting the finished captions (the old
    # approach) reordered the transcript and still let overlapping cues
    # through. Clamping here fixes overlap, zero-duration and reordering at
    # the source, and restores the monotonicity the balanced chunker's
    # feasibility pruning assumes.
    all_words = _normalise_word_times(all_words)

    if max_words > 0:
        captions = _greedy_captions(all_words, max_words, max_chars,
                                    max_lines, max_dur)
    else:
        captions = _balanced_captions(all_words, max_chars, max_lines, max_dur)

    # Enforce minimum duration and add gaps
    for i, cap in enumerate(captions):
        dur = cap["end"] - cap["start"]
        if dur < min_dur:
            cap["end"] = cap["start"] + min_dur
        # A single unsplittable word with a runaway Whisper timestamp would
        # otherwise sit on screen for its full (bogus) duration.
        if cap["end"] - cap["start"] > max_dur:
            cap["end"] = cap["start"] + max_dur

        # Ensure end never overlaps with next caption's start
        if i < len(captions) - 1:
            next_start = captions[i + 1]["start"]
            if cap["end"] + gap_s > next_start:
                cap["end"] = max(cap["start"] + 0.1, next_start - gap_s)
            # Hard clamp: never exceed next caption's start
            if cap["end"] > next_start:
                cap["end"] = next_start
        # Never emit start >= end. The comparison is in MILLISECONDS because
        # that is the SRT's resolution: a 0.3ms sliver passes a float
        # comparison but writes as a zero-duration cue, which makes strict
        # importers reject the whole file. The bump must not run past the
        # next cue either -- otherwise fixing a zero-duration cue creates an
        # overlapping one. If there is genuinely no room, the cue is left
        # zero-length and _merge_zero_duration folds its text into a
        # neighbour so no words are lost.
        if round(cap["end"] * 1000) <= round(cap["start"] * 1000):
            bumped = cap["start"] + 0.001
            if i < len(captions) - 1:
                bumped = min(bumped, captions[i + 1]["start"])
            cap["end"] = bumped

    captions = _merge_zero_duration(captions)

    # Shift into timeline time last: a uniform shift preserves ordering, so
    # every clamp above stays valid.
    if offset_s:
        for cap in captions:
            cap["start"] += offset_s
            cap["end"] += offset_s

    return captions


def _greedy_captions(all_words: list, max_words: int, max_chars: int,
                     max_lines: int, max_dur: float) -> list:
    """Greedy chunker for Reels mode (max_words > 0)."""
    # Held / isolated words land on their own caption for emphasis.
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
        too_many_words = len(block_words) >= max_words

        # Check timing: would this block exceed max duration?
        if block_words:
            block_dur = word.end - block_words[0].start
            too_long = block_dur > max_dur
        else:
            too_long = False

        # Check if too many lines
        # Width matters as well as line count: a word longer than max_chars
        # forms a single over-wide line, which passes a count-only check and
        # lets more words pile onto the same caption.
        too_many_lines = (
            len(test_lines) > max_lines
            or (len(block_words) > 0
                and any(len(l) > max_chars for l in test_lines))
        )

        # Check pause between this word and the previous one
        if block_words:
            pause = word.start - block_words[-1].end
        else:
            pause = 0.0

        # Tiered pause handling: a hard pause always breaks; a smaller pause
        # only breaks once the caption has been on screen long enough that
        # the viewer needs a refresh.
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
        # Abbreviations ("t.ex.", "bl.a.") end in '.' without ending a sentence.
        if block_words:
            prev = block_words[-1].text.rstrip()
            is_abbrev = prev.lower() in _ABBREVIATIONS
            hard_punct = prev.endswith((".", "!", "?")) and not is_abbrev
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

    return captions


# Sentinels that survive the punctuation delete (they are not \w, so they
# are listed explicitly in the keep-set below).
_DEC_COMMA = ""
_DEC_POINT = ""
_SEP_COLON = ""
_SEP_SLASH = ""


def strip_punct_text(text: str) -> str:
    """Remove sentence punctuation but keep word-internal characters.

    Three things this must NOT do:
      - corrupt words: "don't" -> "dont", "e-post" -> "epost" (apostrophes
        and hyphens are kept inside words, trimmed only at the edges);
      - fuse tokens: deleting a separator without leaving a space turned
        "går—nu" into "gårnu", so removals become spaces;
      - corrupt numbers: "3,5 miljoner" must not become "35 miljoner".
        Decimal separators BETWEEN DIGITS are protected, along with '%',
        because these end up burned into client deliverables where a wrong
        figure is both invisible in review and materially wrong.
    """
    import re
    # Sentinels are private-use codepoints; strip any that arrived in the
    # input so they can't be mistaken for our own markers.
    text = text.replace(_DEC_COMMA, "").replace(_DEC_POINT, "")
    text = text.replace(_SEP_COLON, "").replace(_SEP_SLASH, "")
    # En/em dashes separate words ("går—nu").
    text = re.sub(r"[–—]", " ", text)
    # Protect separators that live INSIDE a token. Removing them (or even
    # replacing them with a space) changes the meaning:
    #   "3,5 miljoner" / "1.500 kr"  decimals and thousands
    #   "kl 14:30"                   times
    #   "S:t Eriksgatan", "5:e"      Swedish colon abbreviations/ordinals
    #   "50/50"                      ratios
    _PROTECT = {",": _DEC_COMMA, ".": _DEC_POINT,
                ":": _SEP_COLON, "/": _SEP_SLASH}
    text = re.sub(r"(?<=\d)[.,](?=\d)", lambda m: _PROTECT[m.group(0)], text)
    text = re.sub(r"(?<=\w)[:/](?=\w)", lambda m: _PROTECT[m.group(0)], text,
                  flags=re.UNICODE)
    # Replace (not delete) everything except word chars, whitespace,
    # apostrophes, hyphens, percent and the sentinels.
    text = re.sub(
        rf"[^\w\s'’\-%{_DEC_COMMA}{_DEC_POINT}{_SEP_COLON}{_SEP_SLASH}]",
        " ", text, flags=re.UNICODE)
    text = (text.replace(_DEC_COMMA, ",").replace(_DEC_POINT, ".")
                .replace(_SEP_COLON, ":").replace(_SEP_SLASH, "/"))
    # Trim the keepers when they sit at word edges ("'hello-" -> "hello").
    words = [w.strip("'’-") for w in text.split()]
    return " ".join(w for w in words if w)


def words_to_srt(segments: list, fps: float = 24.0, strip_punctuation: bool = False,
                 offset_s: float = 0.0) -> str:
    """Convert transcription segments into SRT subtitle format.

    offset_s shifts captions into timeline time (see words_to_captions).
    """
    captions = words_to_captions(segments, fps, offset_s=offset_s)
    if not captions:
        return ""

    max_chars = cfg["max_chars_per_line"]
    max_lines = cfg["max_lines"]

    uppercase = bool(cfg.get("uppercase", False))
    srt_lines = []
    index = 0
    for cap in captions:
        text = strip_punct_text(cap["text"]) if strip_punctuation else cap["text"]
        if uppercase:
            text = text.upper()
        text = text.strip()
        # A caption whose text is entirely punctuation (Whisper emits bare
        # "-" / "..." tokens) collapses to "" under --strip-punctuation. An
        # empty body line breaks SRT block framing and every cue after it,
        # so skip the cue and renumber.
        if not text:
            continue

        lines = _split_into_lines(text, max_chars)
        # Deliberately NOT truncated to max_lines. The chunker enforces the
        # limit on the raw text; a transform applied afterwards (uppercase
        # widening "ß" to "SS", a dash becoming a word break) can add a line,
        # and dropping the overflow silently deleted words from the caption.
        # An occasional extra line is a cosmetic issue; losing words is not.
        if len(lines) > max_lines:
            log.debug(f"Caption exceeds {max_lines} line(s) after text "
                      f"transform; keeping all words: {text!r}")

        index += 1
        srt_lines.append(str(index))
        srt_lines.append(
            f"{_format_timestamp(cap['start'])} --> {_format_timestamp(cap['end'])}"
        )
        srt_lines.append("\n".join(lines))
        srt_lines.append("")  # blank line separator

    return "\n".join(srt_lines)


def write_srt(segments: list, output_path: str, fps: float = 24.0,
              strip_punctuation: bool = False, offset_s: float = 0.0):
    """Generate SRT and write to file."""
    content = words_to_srt(segments, fps, strip_punctuation=strip_punctuation,
                           offset_s=offset_s)
    if not content:
        log.warning("No captions generated - empty transcription.")
        return False

    # Write-then-rename: a cancel (SIGTERM) landing mid-write would otherwise
    # leave a truncated file, and in `--file --output existing.srt` mode that
    # file belongs to the user.
    import os
    tmp_path = f"{output_path}.{os.getpid()}.tmp"
    try:
        with open(tmp_path, "w", encoding="utf-8-sig") as f:
            f.write(content)
            if not content.endswith("\n"):
                f.write("\n")
        os.replace(tmp_path, output_path)
    except OSError:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
        raise

    # Count caption blocks (blank-line separated). Counting digit-only lines
    # would also match caption text that happens to be a bare number.
    count = sum(1 for block in content.strip().split("\n\n") if block.strip())
    log.info(f"Wrote {count} captions to {output_path}")
    return True


def write_captions_json(segments: list, output_path: str, fps: float = 24.0,
                        strip_punctuation: bool = False, offset_s: float = 0.0):
    """Generate structured caption data and write to JSON file."""
    captions = words_to_captions(segments, fps, offset_s=offset_s)
    if not captions:
        log.warning("No captions generated - empty transcription.")
        return False

    if cfg.get("uppercase", False):
        for cap in captions:
            cap["text"] = cap["text"].upper()
    if strip_punctuation:
        for cap in captions:
            cap["text"] = strip_punct_text(cap["text"])

    # Same rule as the SRT writer: a caption that transforms to empty text
    # is not a caption.
    captions = [c for c in captions if c["text"].strip()]
    if not captions:
        log.warning("No captions left after text transforms.")
        return False

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(captions, f, ensure_ascii=False)

    log.info(f"Wrote {len(captions)} captions to {output_path}")
    return True
