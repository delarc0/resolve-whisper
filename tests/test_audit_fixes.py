"""Regression tests for the June 2026 production audit fixes.

Run from project root:  ./.venv/bin/python -m unittest tests.test_audit_fixes
"""
import json
import os
import sys
import tempfile
import time
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import caption
from caption import (  # noqa: E402
    _acquire_run_lock,
    _release_run_lock,
    _cleanup_old_srts,
    _is_own_srt,
    _write_status,
)
from srt import (  # noqa: E402
    cfg as srt_cfg,
    strip_punct_text,
    words_to_captions,
    words_to_srt,
    _format_timestamp,
)
from transcribe import (  # noqa: E402
    clean_word,
    is_hallucination,
    _avg_probability,
    Transcriber,
)


class W:
    def __init__(self, text, start, end, probability=0.9):
        self.text = text
        self.start = start
        self.end = end
        self.probability = probability


class Seg:
    def __init__(self, words):
        self.words = words
        self.text = " ".join(w.text for w in words)
        self.start = words[0].start
        self.end = words[-1].end


class TestStripPunctText(unittest.TestCase):
    def test_keeps_intraword_apostrophes(self):
        self.assertEqual(strip_punct_text("don't stop"), "don't stop")
        self.assertEqual(strip_punct_text("it's fine."), "it's fine")

    def test_keeps_intraword_hyphens(self):
        self.assertEqual(strip_punct_text("e-post och Wi-Fi"), "e-post och Wi-Fi")
        self.assertEqual(strip_punct_text("icke-vinstdrivande!"), "icke-vinstdrivande")

    def test_strips_sentence_punctuation(self):
        self.assertEqual(strip_punct_text("Hej, hur mår du?"), "Hej hur mår du")

    def test_trims_edge_apostrophes_and_hyphens(self):
        self.assertEqual(strip_punct_text("'quoted' -dash-"), "quoted dash")

    def test_keeps_digits(self):
        self.assertEqual(strip_punct_text("120 kr, tack!"), "120 kr tack")

    def test_em_dash_between_words_does_not_merge(self):
        # The old [^\w\s] delete merged "går—nu" into "gårnu".
        out = strip_punct_text("går—nu")
        self.assertNotEqual(out, "gårnu")


class TestCleanupOwnFilesOnly(unittest.TestCase):
    def test_pattern_matches_own_files_only(self):
        self.assertTrue(_is_own_srt("My Timeline 20260611-142233.srt"))
        self.assertFalse(_is_own_srt("client_final_v3.srt"))
        self.assertFalse(_is_own_srt("My Timeline.srt"))
        self.assertFalse(_is_own_srt("notes 20260611-142233.txt"))

    def test_foreign_srts_never_deleted(self):
        with tempfile.TemporaryDirectory() as d:
            own_old = os.path.join(d, "Reel 20250101-120000.srt")
            foreign_old = os.path.join(d, "delivered_subs.srt")
            own_new = os.path.join(d, "Reel 20260611-120000.srt")
            for p in (own_old, foreign_old, own_new):
                with open(p, "w") as f:
                    f.write("x")
            old = time.time() - 60 * 86400
            os.utime(own_old, (old, old))
            os.utime(foreign_old, (old, old))

            # Cleanup only ever runs in the tool's OWN output folder.
            orig = caption._default_output_dir
            caption._default_output_dir = lambda: d
            try:
                _cleanup_old_srts(d, keep_days=30)
            finally:
                caption._default_output_dir = orig

            self.assertFalse(os.path.exists(own_old), "own stale file should go")
            self.assertTrue(os.path.exists(foreign_old), "foreign SRT must survive")
            self.assertTrue(os.path.exists(own_new), "fresh own file must survive")

    def test_custom_output_dir_is_never_swept(self):
        # A user pointing output_dir at a delivery folder must not have
        # month-old client subtitles deleted, even if the name happens to
        # match the tool's timestamp pattern.
        with tempfile.TemporaryDirectory() as d:
            victim = os.path.join(d, "Kundleverans FINAL 20250612-093000.srt")
            with open(victim, "w") as f:
                f.write("x")
            old = time.time() - 60 * 86400
            os.utime(victim, (old, old))

            _cleanup_old_srts(d, keep_days=30)  # d != default output dir

            self.assertTrue(os.path.exists(victim),
                            "custom output dir must never be swept")


class TestRunLock(unittest.TestCase):
    def setUp(self):
        self._orig = caption.LOCK_FILE
        fd, self.lock_path = tempfile.mkstemp(suffix=".lock")
        os.close(fd)
        os.unlink(self.lock_path)  # tests want it absent initially
        caption.LOCK_FILE = self.lock_path

    def tearDown(self):
        caption.LOCK_FILE = self._orig
        try:
            os.unlink(self.lock_path)
        except OSError:
            pass

    # The lock is a KERNEL lock held on an open fd, not the existence of the
    # file. That is what makes it survive SIGKILL/TerminateProcess: the OS
    # drops it when the process dies. The file is therefore expected to
    # remain on disk after release; what matters is that it re-acquires.

    def test_acquire_release_roundtrip(self):
        self.assertTrue(_acquire_run_lock())
        with open(self.lock_path) as f:
            self.assertEqual(int(f.read()), os.getpid())
        _release_run_lock()
        # Releasing must make the lock available again...
        self.assertTrue(_acquire_run_lock())
        _release_run_lock()

    def test_second_acquire_in_same_process_is_a_noop(self):
        # Re-acquiring must NOT open a second fd: that would orphan the
        # first one for the process lifetime with no way to release it.
        self.assertTrue(_acquire_run_lock())
        self.assertTrue(_acquire_run_lock())  # same lock, still ours
        _release_run_lock()
        self.assertIsNone(caption._LOCK_FD)

    def test_other_process_is_blocked(self):
        # The contract that matters: a DIFFERENT process cannot run while we
        # hold the lock (flock is per-process, so this needs a real fork).
        import subprocess
        import sys as _sys
        self.assertTrue(_acquire_run_lock())
        try:
            probe = (
                "import sys; sys.path.insert(0, %r); "
                "import caption; caption.LOCK_FILE = %r; "
                "print('ACQUIRED' if caption._acquire_run_lock() else 'BLOCKED')"
                % (os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   self.lock_path)
            )
            out = subprocess.run([_sys.executable, "-c", probe],
                                 capture_output=True, text=True, timeout=60)
            self.assertIn("BLOCKED", out.stdout,
                          f"second process was not blocked: {out.stdout}{out.stderr}")
        finally:
            _release_run_lock()

    def test_lock_frees_when_holder_is_killed(self):
        # A kernel lock dies with the process -- this is what removes the
        # old "crashed run wedges the tool" failure mode entirely.
        import subprocess
        import sys as _sys
        import time as _time
        holder_src = (
            "import sys, time; sys.path.insert(0, %r); "
            "import caption; caption.LOCK_FILE = %r; "
            "caption._acquire_run_lock(); print('HELD', flush=True); "
            "time.sleep(60)"
            % (os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
               self.lock_path)
        )
        holder = subprocess.Popen([_sys.executable, "-c", holder_src],
                                  stdout=subprocess.PIPE, text=True)
        try:
            self.assertEqual(holder.stdout.readline().strip(), "HELD")
            self.assertFalse(_acquire_run_lock(), "lock should be held")
            holder.kill()
            holder.wait(timeout=10)
            _time.sleep(0.2)
            self.assertTrue(_acquire_run_lock(),
                            "SIGKILL must release the lock immediately")
            _release_run_lock()
        finally:
            if holder.poll() is None:
                holder.kill()

    def test_leftover_file_with_dead_pid_is_not_a_lock(self):
        # A file left by a crashed run holds no kernel lock, so the next run
        # takes over immediately -- no 2h Windows lockout, no PID-reuse wedge.
        with open(self.lock_path, "w") as f:
            f.write("999999")
        self.assertTrue(_acquire_run_lock())
        _release_run_lock()

    def test_garbage_lock_taken_over(self):
        with open(self.lock_path, "w") as f:
            f.write("not-a-pid")
        self.assertTrue(_acquire_run_lock())
        _release_run_lock()

    def test_release_without_holding_is_safe(self):
        # Never raises, and must not disturb a file we do not own.
        with open(self.lock_path, "w") as f:
            f.write("999999")
        _release_run_lock()
        self.assertTrue(os.path.exists(self.lock_path))


class TestAtomicStatusWrite(unittest.TestCase):
    def test_status_file_always_valid_json(self):
        orig = caption.STATUS_FILE
        fd, path = tempfile.mkstemp(suffix=".json")
        os.close(fd)
        caption.STATUS_FILE = path
        try:
            for i in range(50):
                _write_status("transcribing", f"tick {i}", progress=i)
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)  # must never be half-written
                self.assertEqual(data["message"], f"tick {i}")
                self.assertEqual(data["pid"], os.getpid())
            # No tmp litter left behind
            litter = [p for p in os.listdir(os.path.dirname(path))
                      if p.startswith(os.path.basename(path)) and p.endswith(".tmp")]
            self.assertEqual(litter, [])
        finally:
            caption.STATUS_FILE = orig
            os.unlink(path)


class TestCleanWord(unittest.TestCase):
    def test_fillers_stripped(self):
        for filler in ("um", "uh", "öh", "asså", "Hmm"):
            self.assertEqual(clean_word(filler), "")

    def test_oh_is_kept(self):
        # "oh" carries meaning in English; removed from the filler list.
        self.assertEqual(clean_word("oh"), "oh")

    def test_alltsa_is_kept(self):
        # "alltså" means "thus/therefore" -- a real word, not a filler.
        # Deleting it also deleted its sentence-ending punctuation, which
        # merged two sentences into one caption stretch.
        self.assertEqual(clean_word("alltså"), "alltså")

    def test_annotations_removed(self):
        self.assertEqual(clean_word("[Music]"), "")
        self.assertEqual(clean_word("[inaudible]"), "")

    def test_annotations_with_trailing_punctuation_removed(self):
        # "[Musik]," used to leak "Musik]," into the caption.
        self.assertEqual(clean_word("[Musik],"), "")
        self.assertEqual(clean_word("[Musik]."), "")
        self.assertEqual(clean_word("(skratt)"), "")

    def test_symbol_only_tokens_removed(self):
        # These become captions whose text is empty once punctuation is
        # stripped, which produced a malformed SRT block.
        for junk in ("-", "...", "♪", "–"):
            self.assertEqual(clean_word(junk), "")

    def test_real_words_kept(self):
        self.assertEqual(clean_word(" hej "), "hej")
        self.assertEqual(clean_word("typ"), "typ")


class TestHallucinationConfidenceGate(unittest.TestCase):
    def test_genuine_thank_you_survives(self):
        # High word confidence = real speech, even if it matches the list.
        self.assertFalse(is_hallucination("Thank you.", avg_word_probability=0.92))

    def test_low_confidence_match_dropped(self):
        self.assertTrue(is_hallucination("Thank you.", avg_word_probability=0.2))

    def test_no_probabilities_keeps_old_behavior(self):
        self.assertTrue(is_hallucination("Tack för att ni tittade", 0.0))

    def test_non_matching_text_never_dropped(self):
        self.assertFalse(is_hallucination("Vi ses imorgon.", 0.0))

    def test_avg_probability_helper(self):
        self.assertEqual(_avg_probability([]), 0.0)
        self.assertEqual(_avg_probability([None, None]), 0.0)
        self.assertAlmostEqual(_avg_probability([0.4, None, 0.8]), 0.6)


class TestOverlapClampMonotonic(unittest.TestCase):
    def setUp(self):
        self._saved = dict(srt_cfg)
        srt_cfg.update({
            "max_words_per_caption": 2,
            "max_chars_per_line": 26,
            "max_lines": 1,
            "min_duration_s": 1.0,
            "max_duration_s": 7.0,
            "gap_frames": 2,
        })

    def tearDown(self):
        srt_cfg.clear()
        srt_cfg.update(self._saved)

    def test_non_monotonic_words_never_yield_start_after_end(self):
        # Whisper clustered timestamps after a pause: a later word starts
        # BEFORE an earlier caption.
        words = [
            W("Spruta", 18.20, 18.64),
            W("lite", 18.72, 18.82),
            W("Och", 18.82, 18.82),   # zero-duration word
            W("så", 18.81, 18.82),    # starts before previous end
            W("är", 25.74, 26.14),
            W("det", 26.14, 26.14),   # zero-duration at the tail
        ]
        caps = words_to_captions([Seg(words)], fps=25.0)
        self.assertTrue(caps)
        for c in caps:
            self.assertGreater(c["end"], c["start"],
                               f"start>=end in {c!r}")

    def test_sorted_output(self):
        words = [W("b", 5.0, 5.4), W("a", 1.0, 1.4)]
        caps = words_to_captions([Seg([words[0]]), Seg([words[1]])], fps=25.0)
        starts = [c["start"] for c in caps]
        self.assertEqual(starts, sorted(starts))


class TestSwedishAbbreviations(unittest.TestCase):
    def setUp(self):
        self._saved = dict(srt_cfg)
        srt_cfg.update({
            "max_words_per_caption": 0,
            "max_chars_per_line": 100,
            "max_lines": 1,
            "min_duration_s": 1.0,
            "max_duration_s": 7.0,
            "gap_frames": 2,
        })

    def tearDown(self):
        srt_cfg.clear()
        srt_cfg.update(self._saved)

    def test_tex_does_not_break_sentence(self):
        words = [
            W("Vi", 0.0, 0.1), W("har", 0.12, 0.2), W("t.ex.", 0.22, 0.4),
            W("bastu", 0.42, 0.6), W("här", 0.62, 0.8),
        ]
        caps = words_to_captions([Seg(words)], fps=25.0)
        self.assertEqual(len(caps), 1)
        self.assertIn("t.ex. bastu", caps[0]["text"])

    def test_real_period_still_breaks(self):
        words = [
            W("klart.", 0.0, 0.3),
            W("Sen", 0.32, 0.5), W("då", 0.52, 0.7),
        ]
        caps = words_to_captions([Seg(words)], fps=25.0)
        self.assertEqual(len(caps), 2)


class TestFormatTimestampRounding(unittest.TestCase):
    def test_ms_rounding_carries_into_seconds(self):
        # 59.9996s rounds to 60000ms == exactly 00:01:00,000 (never 00:00:60)
        self.assertEqual(_format_timestamp(59.9996), "00:01:00,000")

    def test_ms_rounding_carries_into_hours(self):
        self.assertEqual(_format_timestamp(3599.9996), "01:00:00,000")


class TestWordsToSrtStructure(unittest.TestCase):
    def setUp(self):
        self._saved = dict(srt_cfg)
        srt_cfg.update({
            "max_words_per_caption": 0,
            "max_chars_per_line": 42,
            "max_lines": 1,
            "min_duration_s": 1.0,
            "max_duration_s": 7.0,
            "gap_frames": 2,
        })

    def tearDown(self):
        srt_cfg.clear()
        srt_cfg.update(self._saved)

    def test_uppercase_applied_when_enabled(self):
        srt_cfg["uppercase"] = True
        words = [W("hej på er.", 0.0, 0.5)]
        out = words_to_srt([Seg(words)], fps=25.0)
        self.assertIn("HEJ PÅ ER.", out)

    def test_uppercase_off_preserves_case(self):
        srt_cfg["uppercase"] = False
        words = [W("Hej på er.", 0.0, 0.5)]
        out = words_to_srt([Seg(words)], fps=25.0)
        self.assertIn("Hej på er.", out)

    def test_srt_block_structure(self):
        words = [
            W("Hej.", 0.0, 0.5),
            W("Hejdå.", 2.0, 2.5),
        ]
        out = words_to_srt([Seg(words)], fps=25.0)
        blocks = [b for b in out.split("\n\n") if b.strip()]
        self.assertEqual(len(blocks), 2)
        for i, block in enumerate(blocks, 1):
            lines = block.split("\n")
            self.assertEqual(lines[0], str(i))
            self.assertRegex(
                lines[1],
                r"^\d{2}:\d{2}:\d{2},\d{3} --> \d{2}:\d{2}:\d{2},\d{3}$",
            )
            self.assertTrue(lines[2].strip())


class TestLoadMono16k(unittest.TestCase):
    def _write_wav(self, sr, seconds, channels=1, extra_samples=0):
        import numpy as np
        import soundfile as sf
        n = int(sr * seconds) + extra_samples
        t = (440.0 * 2 * 3.141592653589793 / sr) * np.arange(n)
        audio = (0.25 * np.sin(t)).astype(np.float32)
        if channels == 2:
            audio = np.stack([audio, audio], axis=1)
        fd, path = tempfile.mkstemp(suffix=".wav")
        os.close(fd)
        sf.write(path, audio, sr)
        return path, n

    def test_48k_stereo_resamples_to_correct_length(self):
        path, n = self._write_wav(48000, 2.0, channels=2)
        try:
            out = Transcriber._load_mono_16k(path)
            expected = n * 16000 / 48000
            self.assertEqual(out.ndim, 1, "must be mono")
            # within 1ms worth of samples per 30s block boundary
            self.assertLess(abs(len(out) - expected), 32)
        finally:
            os.unlink(path)

    def test_16k_passthrough_exact(self):
        path, n = self._write_wav(16000, 1.0)
        try:
            out = Transcriber._load_mono_16k(path)
            self.assertEqual(len(out), n)
        finally:
            os.unlink(path)


if __name__ == "__main__":
    unittest.main()
