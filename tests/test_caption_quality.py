"""Regression tests for caption-quality defects found in the 2026-07 audit.

These are output-corruption bugs (wrong text in a client deliverable),
not crashes, so each test pins the exact bad output that was observed.
"""
import os
import sys
import unittest
from dataclasses import dataclass

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from srt import (  # noqa: E402
    strip_punct_text,
    words_to_captions,
    words_to_srt,
)


@dataclass
class W:
    text: str
    start: float
    end: float
    probability: float = 1.0


@dataclass
class S:
    text: str
    start: float
    end: float
    words: list


def seg(words):
    return S(text=" ".join(w.text for w in words),
             start=words[0].start, end=words[-1].end, words=words)


class TestNumbersSurvivePunctuationStripping(unittest.TestCase):
    """The Reels preset ships --strip-punctuation. Deleting a decimal
    separator turned 3,5 million into 35 million: a wrong figure burned into
    a deliverable, invisible in review."""

    def test_swedish_decimal_comma(self):
        self.assertEqual(strip_punct_text("3,5 miljoner"), "3,5 miljoner")

    def test_decimal_point(self):
        self.assertEqual(strip_punct_text("1.500 kr"), "1.500 kr")

    def test_multiplier(self):
        self.assertEqual(strip_punct_text("1,5x"), "1,5x")

    def test_percent_kept(self):
        self.assertEqual(strip_punct_text("50% rabatt"), "50% rabatt")

    def test_sentence_punctuation_still_removed(self):
        self.assertEqual(strip_punct_text("hej, vad? heter. du!"),
                         "hej vad heter du")

    def test_time_is_not_fused(self):
        # "kl. 14.30" -> the abbreviation dot goes, the time's dot stays
        self.assertEqual(strip_punct_text("kl. 14.30"), "kl 14.30")

    def test_words_never_fuse(self):
        self.assertEqual(strip_punct_text("går—nu"), "går nu")
        self.assertEqual(strip_punct_text("a,, b!! c"), "a b c")

    def test_word_internal_chars_kept(self):
        self.assertEqual(strip_punct_text("don't e-post"), "don't e-post")

    def test_time_colon_survives(self):
        # The first fix protected commas but split on ':' instead --
        # "kl 14:30" became "kl 14 30". Same bug class, new separator.
        self.assertEqual(strip_punct_text("Vi ses kl 14:30 i lokalen"),
                         "Vi ses kl 14:30 i lokalen")

    def test_swedish_colon_abbreviations(self):
        self.assertEqual(strip_punct_text("S:t Eriksgatan 5"),
                         "S:t Eriksgatan 5")
        self.assertEqual(strip_punct_text("det var 5:e gången"),
                         "det var 5:e gången")

    def test_ratio_slash_survives(self):
        self.assertEqual(strip_punct_text("resultatet blev 50/50"),
                         "resultatet blev 50/50")

    def test_input_sentinels_cannot_leak(self):
        # The implementation uses private-use codepoints as markers; input
        # containing them must not be able to forge a separator.
        self.assertNotIn(",", strip_punct_text("ab"))
        self.assertNotIn(".", strip_punct_text("xy"))


class TestBoundTokenMerging(unittest.TestCase):
    """KB-Whisper emits "3,5" as the tokens "3" and ",5". Joining them only
    at render time is not enough: the chunker could put them on separate
    caption cards, i.e. a wrong number on screen."""

    def setUp(self):
        from srt import cfg
        self.cfg = cfg
        self._saved = dict(cfg)
        cfg.update({
            "max_words_per_caption": 3,
            "max_chars_per_line": 22,
            "max_lines": 1,
            "min_duration_s": 1.0,
            "max_duration_s": 7.0,
            "gap_frames": 2,
            "uppercase": False,
        })

    def tearDown(self):
        self.cfg.clear()
        self.cfg.update(self._saved)

    def test_decimal_token_never_splits_across_captions(self):
        words = [W("Vi", 0.0, 0.1), W("omsatte", 0.18, 0.66),
                 W("3", 0.74, 1.0), W(",5", 1.08, 1.64),
                 W("miljoner", 1.72, 2.30)]
        caps = words_to_captions([seg(words)], fps=25.0)
        texts = [c["text"] for c in caps]
        self.assertTrue(any("3,5" in t for t in texts),
                        f"decimal was split across captions: {texts}")
        for t in texts:
            self.assertNotEqual(t.strip(), "3")
            self.assertNotEqual(t.strip(), ",5")

    def test_reels_output_keeps_the_number(self):
        words = [W("3", 0.0, 0.3), W(",5", 0.35, 0.7),
                 W("miljoner", 0.8, 1.3)]
        out = words_to_srt([seg(words)], fps=25.0, strip_punctuation=True)
        self.assertIn("3,5", out)

    def test_caption_never_ends_on_dangling_comma(self):
        words = [W("hej", 0.0, 0.3), W(",", 0.31, 0.33),
                 W("då", 0.4, 0.8)]
        caps = words_to_captions([seg(words)], fps=25.0)
        for cap in caps:
            self.assertFalse(cap["text"].strip().startswith(","),
                             f"caption starts with a comma: {cap}")

    def test_trailing_period_token_merges(self):
        words = [W("slut", 0.0, 0.4), W(".", 0.41, 0.43)]
        caps = words_to_captions([seg(words)], fps=25.0)
        self.assertEqual(caps[0]["text"], "slut.")


class TestNoMalformedCues(unittest.TestCase):
    def setUp(self):
        from srt import cfg
        self.cfg = cfg
        self._saved = dict(cfg)
        cfg.update({
            "max_words_per_caption": 3,
            "max_chars_per_line": 22,
            "max_lines": 1,
            "min_duration_s": 1.0,
            "max_duration_s": 7.0,
            "gap_frames": 2,
            "uppercase": True,
        })

    def tearDown(self):
        self.cfg.clear()
        self.cfg.update(self._saved)

    def test_punctuation_only_word_does_not_emit_empty_cue(self):
        # Whisper emits bare "-" / "..." tokens; under --strip-punctuation
        # they collapse to "", which used to write an empty body line and
        # break block framing for every cue after it.
        words = [W("hej", 0.0, 0.4), W("-", 0.5, 0.55),
                 W("ungefär", 1.0, 1.6)]
        out = words_to_srt([seg(words)], fps=25.0, strip_punctuation=True)
        self.assertNotIn("\n\n\n", out)
        for block in out.strip().split("\n\n"):
            lines = block.split("\n")
            self.assertGreaterEqual(len(lines), 3, f"malformed block: {block!r}")
            self.assertTrue(lines[2].strip(), f"empty caption text: {block!r}")

    def test_cue_numbering_is_contiguous(self):
        words = [W("hej", 0.0, 0.4), W("...", 0.5, 0.55),
                 W("då", 1.0, 1.4), W("!", 1.5, 1.55),
                 W("igen", 2.0, 2.4)]
        out = words_to_srt([seg(words)], fps=25.0, strip_punctuation=True)
        numbers = [int(b.split("\n")[0]) for b in out.strip().split("\n\n") if b.strip()]
        self.assertEqual(numbers, list(range(1, len(numbers) + 1)))

    def test_no_words_dropped_by_line_truncation(self):
        # A transform that widens the text (ß -> SS under uppercase) used to
        # push a word onto a second line, which was then truncated away.
        from srt import cfg
        cfg.update({"max_words_per_caption": 0, "max_chars_per_line": 42,
                    "uppercase": True})
        text = "vi bor pa Straße nara torget i staden idag"
        words, t = [], 0.0
        for w in text.split():
            words.append(W(w, t, t + 0.3))
            t += 0.32
        out = words_to_srt([seg(words)], fps=25.0)
        for w in text.split():
            self.assertIn(w.upper(), out.upper(),
                          f"word {w!r} was dropped from the SRT")


class TestTimingInvariants(unittest.TestCase):
    def setUp(self):
        from srt import cfg
        self.cfg = cfg
        self._saved = dict(cfg)
        cfg.update({
            "max_words_per_caption": 0,
            "max_chars_per_line": 42,
            "max_lines": 1,
            "min_duration_s": 1.0,
            "max_duration_s": 7.0,
            "gap_frames": 2,
            "uppercase": False,
        })

    def tearDown(self):
        self.cfg.clear()
        self.cfg.update(self._saved)

    def _assert_valid(self, caps):
        for i, cap in enumerate(caps):
            self.assertGreater(round(cap["end"] * 1000), round(cap["start"] * 1000),
                               f"zero-duration cue at {i}: {cap}")
            if i:
                self.assertLessEqual(caps[i - 1]["end"], cap["start"] + 1e-9,
                                     f"overlapping cues at {i}: {caps[i-1]} / {cap}")

    def test_duplicate_start_times_do_not_overlap(self):
        words = [W("Vi", 0.0, 0.30), W("ses", 0.35, 0.70), W("snart.", 0.75, 1.20),
                 W("Ja.", 1.40, 1.40), W("Hej.", 1.40, 1.75), W("Da", 2.10, 2.40)]
        self._assert_valid(words_to_captions([seg(words)], fps=25.0))

    def test_sub_millisecond_cluster_has_no_zero_duration(self):
        words = [W("hej", 1.0, 1.0), W("da.", 1.0, 1.0),
                 W("Nästa", 1.0003, 1.0003), W("ord.", 1.0006, 1.0006),
                 W("Sen", 5.0, 5.3)]
        self._assert_valid(words_to_captions([seg(words)], fps=25.0))

    def test_out_of_order_words_keep_transcript_order(self):
        # Sorting finished captions reordered the transcript; normalising the
        # word stamps instead preserves reading order.
        words = [W("Ett.", 10.0, 10.4), W("Tva.", 3.0, 3.4), W("Tre.", 11.0, 11.4)]
        caps = words_to_captions([seg(words)], fps=25.0)
        order = [c["text"] for c in caps]
        self.assertEqual(order, ["Ett.", "Tva.", "Tre."])
        self._assert_valid(caps)

    def test_single_word_respects_max_duration(self):
        # A runaway Whisper stamp parked one word on screen for 24s.
        words = [W("ja", 0.0, 0.3), W("mmm", 1.0, 25.0), W("okej", 26.0, 26.4)]
        caps = words_to_captions([seg(words)], fps=25.0)
        for cap in caps:
            self.assertLessEqual(cap["end"] - cap["start"], 7.0 + 1e-6,
                                 f"cue longer than max_duration_s: {cap}")


if __name__ == "__main__":
    unittest.main()
