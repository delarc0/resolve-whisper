"""Tests for srt.py chunking + formatting helpers.

Run from project root:  ./.venv/bin/python -m unittest tests/test_srt.py
"""
import os
import sys
import unittest
from dataclasses import dataclass

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from srt import (  # noqa: E402
    _format_timestamp,
    _split_into_lines,
    strip_punct_text,
    words_to_captions,
)


@dataclass
class W:
    """Minimal Word stand-in for transcribe.Word."""
    text: str
    start: float
    end: float
    probability: float = 1.0


@dataclass
class S:
    """Minimal Segment stand-in."""
    text: str
    start: float
    end: float
    words: list


def seg(words):
    return S(text=" ".join(w.text for w in words),
             start=words[0].start, end=words[-1].end, words=words)


class TestFormatTimestamp(unittest.TestCase):
    def test_zero(self):
        self.assertEqual(_format_timestamp(0), "00:00:00,000")

    def test_negative_clamps_to_zero(self):
        self.assertEqual(_format_timestamp(-5), "00:00:00,000")

    def test_milliseconds(self):
        self.assertEqual(_format_timestamp(1.234), "00:00:01,234")

    def test_minutes_seconds(self):
        self.assertEqual(_format_timestamp(125.5), "00:02:05,500")

    def test_hours(self):
        self.assertEqual(_format_timestamp(3661.001), "01:01:01,001")


class TestSplitIntoLines(unittest.TestCase):
    def test_short_fits_one_line(self):
        self.assertEqual(_split_into_lines("Hej", 42), ["Hej"])

    def test_wraps_at_word_boundary(self):
        self.assertEqual(
            _split_into_lines("Vi ska käka lite frukost", 12),
            ["Vi ska käka", "lite frukost"],
        )

    def test_one_oversized_word_stays_alone(self):
        # Single word longer than max_chars survives as its own line
        self.assertEqual(_split_into_lines("ettjättelångtord", 5), ["ettjättelångtord"])


class TestStripPunctText(unittest.TestCase):
    def test_basic(self):
        self.assertEqual(strip_punct_text("hej, vad? heter. du!"), "hej vad heter du")

    def test_keeps_unicode_letters(self):
        self.assertEqual(strip_punct_text("är åäö."), "är åäö")

    def test_collapses_extra_spaces(self):
        self.assertEqual(strip_punct_text("a,, b!! c"), "a b c")


class TestWordsToCaptions(unittest.TestCase):
    """words_to_captions reads max_words/max_chars/etc from cfg, so we monkey-
    patch the cfg dict for each scenario."""

    def setUp(self):
        from srt import cfg
        self.cfg = cfg
        # Snapshot
        self._saved = dict(cfg)
        cfg.update({
            "max_words_per_caption": 4,
            "max_chars_per_line": 26,
            "max_lines": 1,
            "min_duration_s": 1.0,
            "max_duration_s": 7.0,
            "gap_frames": 2,
        })

    def tearDown(self):
        self.cfg.clear()
        self.cfg.update(self._saved)

    def test_period_forces_split_even_with_no_pause(self):
        # "frukost." 0.0-0.5, "Men" 0.51-0.6 -- pause 0.01s, but period must split
        words = [
            W("Vi", 0.0, 0.1), W("ska", 0.12, 0.2), W("käka", 0.22, 0.3),
            W("frukost.", 0.32, 0.5), W("Men", 0.51, 0.6),
        ]
        caps = words_to_captions([seg(words)], fps=25.0)
        # First caption ends with "frukost.", next starts with "Men"
        self.assertTrue(caps[0]["text"].endswith("frukost."))
        self.assertTrue(caps[1]["text"].startswith("Men"))

    def test_max_words_caps_chunk_size(self):
        words = [W(str(i), i * 0.1, i * 0.1 + 0.05) for i in range(10)]
        caps = words_to_captions([seg(words)], fps=25.0)
        # 10 words / max 4 per caption = at least 3 captions
        self.assertGreaterEqual(len(caps), 3)
        for cap in caps:
            self.assertLessEqual(len(cap["text"].split()), 4)

    def test_micro_pause_breaks_when_caption_old(self):
        # Long block (>0.6s) followed by a small pause (>0.06s) -> should break
        words = [
            W("a", 0.00, 0.10), W("b", 0.11, 0.20), W("c", 0.21, 0.30),
            W("d", 0.31, 0.70),                          # block now 0.7s long
            W("e", 0.79, 0.85),                          # 0.09s gap -- should split
        ]
        caps = words_to_captions([seg(words)], fps=25.0)
        # "e" should be in a new caption, not appended to the first
        self.assertEqual(len(caps), 2)
        self.assertEqual(caps[1]["text"], "e")

    def test_micro_pause_ignored_when_caption_young(self):
        # Same micro-pause but block is too young -- shouldn't split
        words = [
            W("a", 0.00, 0.10), W("b", 0.11, 0.20),     # block only 0.2s
            W("c", 0.29, 0.40),                          # 0.09s gap
        ]
        caps = words_to_captions([seg(words)], fps=25.0)
        self.assertEqual(len(caps), 1)

    def test_emphasis_solo_disabled_in_plain_srt_mode(self):
        # max_words=0 (Auto/Podcast presets): a long-held word ("Stockholms"
        # 0.52s) should NOT force its own caption. In Reels mode it would,
        # but plain SRT wants full sentences.
        from srt import cfg
        cfg["max_words_per_caption"] = 0
        cfg["max_chars_per_line"] = 100
        words = [
            W("hotell", 6.00, 6.40),
            W("är", 6.48, 6.60),
            W("Stockholms", 6.98, 7.50),  # 0.52s -- would be "emphasized"
            W("mest", 7.58, 7.70),
        ]
        caps = words_to_captions([seg(words)], fps=25.0)
        self.assertEqual(len(caps), 1)
        self.assertIn("Stockholms", caps[0]["text"])

    def test_micro_pause_ignored_in_plain_srt_mode(self):
        # max_words=0 (Auto/Podcast presets): micro-pauses should NOT fragment
        # captions. The chunker should fill up to char/line limit instead.
        from srt import cfg
        cfg["max_words_per_caption"] = 0
        cfg["max_chars_per_line"] = 100  # remove char pressure
        words = [
            W("a", 0.00, 0.10), W("b", 0.11, 0.20), W("c", 0.21, 0.30),
            W("d", 0.31, 0.70),                          # block now 0.7s long
            W("e", 0.79, 0.85),                          # 0.09s micro-pause
        ]
        caps = words_to_captions([seg(words)], fps=25.0)
        # In plain-SRT mode the micro-pause should be ignored.
        self.assertEqual(len(caps), 1)

    def test_hard_pause_always_breaks(self):
        # >0.18s pause splits regardless of block age, when neither word is
        # in the connector list ("ja" / "nej" are content words in Swedish)
        words = [
            W("ja", 0.0, 0.05),
            W("nej", 0.30, 0.40),  # 0.25s gap -- hard pause
        ]
        caps = words_to_captions([seg(words)], fps=25.0)
        self.assertEqual(len(caps), 2)

    def test_min_duration_extends_short_caption(self):
        words = [W("a", 0.0, 0.05), W("b", 1.0, 1.05)]
        caps = words_to_captions([seg(words)], fps=25.0)
        # First "a" caption only 0.05s in raw words; should extend to min 1.0s
        first = caps[0]
        self.assertGreaterEqual(first["end"] - first["start"], 0.5)

    def test_no_break_after_swedish_connector(self):
        # "Bob och Tim" -- micro-pause before "Tim" but block ends on "och"
        # (connector), so the chunker should keep going. Test isolates the
        # connector-suppression logic from char/line limits.
        from srt import cfg
        cfg["max_chars_per_line"] = 100  # remove char-limit pressure
        words = [
            W("Bob", 0.00, 0.30),
            W("och", 0.40, 0.70),  # block dur reaches 0.7s by end of "och"
            W("Tim", 0.85, 1.10),  # 0.15s gap -- micro_pause would normally fire
        ]
        caps = words_to_captions([seg(words)], fps=25.0)
        # Connector suppresses the break, so it stays as one caption
        self.assertEqual(len(caps), 1)
        self.assertIn("Tim", caps[0]["text"])

    def test_break_after_connector_still_works_on_period(self):
        # Hard punctuation always wins, even after a connector
        words = [
            W("Thomas",  0.00, 0.30),
            W("Sjögren.", 0.32, 0.60),  # period -> hard break
            W("Den",      0.65, 0.80),
        ]
        caps = words_to_captions([seg(words)], fps=25.0)
        self.assertEqual(len(caps), 2)

    def test_long_word_gets_solo_caption(self):
        # A word held >= 0.45s = emphasized, should stand alone
        words = [
            W("ja",  0.00, 0.20),
            W("absolut", 0.25, 0.85),  # 0.6s duration -- emphasized
            W("inte", 0.90, 1.10),
        ]
        caps = words_to_captions([seg(words)], fps=25.0)
        # Expect: ["ja", "absolut", "inte"]  -- 3 separate captions
        self.assertEqual(len(caps), 3)
        self.assertEqual(caps[1]["text"], "absolut")

    def test_isolated_word_gets_solo_caption(self):
        # Word with >= 0.25s pause on both sides = isolated, solo
        words = [
            W("hej", 0.00, 0.20),
            W("vad",  0.60, 0.75),   # 0.40s pause before, 0.30s after -> solo
            W("säger", 1.05, 1.30),
            W("du",    1.32, 1.45),
        ]
        caps = words_to_captions([seg(words)], fps=25.0)
        # "vad" is isolated -> own caption
        texts = [c["text"] for c in caps]
        self.assertIn("vad", texts)

    def test_normal_words_dont_get_solo(self):
        # Continuous speech, no long words, no big pauses -> normal grouping
        words = [
            W("ett", 0.0, 0.10), W("två", 0.11, 0.20),
            W("tre", 0.21, 0.30), W("fyra", 0.31, 0.40),
        ]
        caps = words_to_captions([seg(words)], fps=25.0)
        # Should be a single caption since nothing is emphasized
        self.assertEqual(len(caps), 1)


if __name__ == "__main__":
    unittest.main()
