"""In/out-range offset: captions must land where the speech is.

Resolve renders only the in/out range, so Whisper timestamps are zero-based
at the IN point, while an imported SRT's 00:00:00 maps to the TIMELINE start.
Without a shift every caption piles up at the top of the timeline.
"""
import os
import sys
import unittest
from dataclasses import dataclass

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import caption  # noqa: E402
from srt import words_to_captions, words_to_srt  # noqa: E402


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


class FakeTimeline:
    """Minimal stand-in for a Resolve Timeline object."""

    def __init__(self, start_frame=86400, marks=None):
        self._start = start_frame
        self._marks = marks

    def GetStartFrame(self):
        return self._start

    def GetMarkInOut(self):
        if self._marks is None:
            raise RuntimeError("no marks set")  # mirrors Resolve raising
        return self._marks


class TestRenderOffset(unittest.TestCase):
    """caption._render_offset_s converts MarkIn into a seconds offset."""

    def test_markin_from_job_settings(self):
        # Timeline starts at 01:00:00:00 (86400 @ 24fps); in-point 10s later.
        tl = FakeTimeline(start_frame=86400)
        off = caption._render_offset_s({"MarkIn": 86400 + 240}, tl, 24.0)
        self.assertAlmostEqual(off, 10.0, places=6)

    def test_no_markin_no_marks_is_zero(self):
        tl = FakeTimeline(start_frame=86400)
        self.assertEqual(caption._render_offset_s({}, tl, 24.0), 0.0)

    def test_markin_equal_start_is_zero(self):
        # Whole-timeline render: MarkIn == timeline start -> no shift.
        tl = FakeTimeline(start_frame=86400)
        self.assertEqual(caption._render_offset_s({"MarkIn": 86400}, tl, 24.0), 0.0)

    def test_negative_offset_clamped_to_zero(self):
        # Defensive: a MarkIn in a different frame base must not shift back.
        tl = FakeTimeline(start_frame=86400)
        self.assertEqual(caption._render_offset_s({"MarkIn": 0}, tl, 24.0), 0.0)

    def test_falls_back_to_timeline_marks(self):
        tl = FakeTimeline(start_frame=0,
                          marks={"video": {"in": 250, "out": 500}})
        off = caption._render_offset_s({}, tl, 25.0)
        self.assertAlmostEqual(off, 10.0, places=6)

    def test_zero_fps_is_safe(self):
        tl = FakeTimeline(start_frame=0)
        self.assertEqual(caption._render_offset_s({"MarkIn": 240}, tl, 0.0), 0.0)

    def test_bad_types_are_safe(self):
        tl = FakeTimeline(start_frame=86400)
        self.assertEqual(caption._render_offset_s({"MarkIn": "abc"}, tl, 24.0), 0.0)

    def test_lowercase_markin_key(self):
        tl = FakeTimeline(start_frame=0)
        off = caption._render_offset_s({"markIn": 48}, tl, 24.0)
        self.assertAlmostEqual(off, 2.0, places=6)


class TestCaptionShift(unittest.TestCase):
    """The offset must move every caption by exactly the same amount."""

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

    def _words(self):
        return [W("hej", 0.0, 0.4), W("där.", 0.45, 0.9),
                W("Hur", 1.6, 1.9), W("går", 1.95, 2.2), W("det?", 2.25, 2.6)]

    def test_shift_moves_all_captions(self):
        base = words_to_captions([seg(self._words())], fps=25.0)
        shifted = words_to_captions([seg(self._words())], fps=25.0, offset_s=600.0)
        self.assertEqual(len(base), len(shifted))
        for b, s in zip(base, shifted):
            self.assertAlmostEqual(s["start"] - b["start"], 600.0, places=6)
            self.assertAlmostEqual(s["end"] - b["end"], 600.0, places=6)
            self.assertEqual(b["text"], s["text"])

    def test_shift_preserves_ordering_and_durations(self):
        shifted = words_to_captions([seg(self._words())], fps=25.0, offset_s=123.456)
        for i, cap in enumerate(shifted):
            self.assertLess(cap["start"], cap["end"])
            if i:
                self.assertLessEqual(shifted[i - 1]["end"], cap["start"] + 1e-9)

    def test_zero_offset_is_identical(self):
        a = words_to_captions([seg(self._words())], fps=25.0)
        b = words_to_captions([seg(self._words())], fps=25.0, offset_s=0.0)
        self.assertEqual(a, b)

    def test_srt_text_shows_shifted_timecode(self):
        # 10 minutes in: first caption must render at 00:10:xx, not 00:00:xx
        srt = words_to_srt([seg(self._words())], fps=25.0, offset_s=600.0)
        first_stamp = srt.split("\n")[1]
        self.assertTrue(first_stamp.startswith("00:10:00"),
                        f"expected shifted timecode, got {first_stamp}")


if __name__ == "__main__":
    unittest.main()
