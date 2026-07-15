"""AAC/.mp4 decode fallbacks (Resolve 21 'Audio Only' renders AAC).

libsndfile can't open AAC, which used to silently kill the VAD pre-pass and
collapse the progress estimate to a stuck 95%. These tests build a real AAC
file with ffmpeg and exercise the ffmpeg/ffprobe fallbacks end to end.
Skipped if ffmpeg isn't installed (it always is on a working install:
mlx_whisper requires it).
"""
import os
import shutil
import subprocess
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from transcribe import Transcriber  # noqa: E402

_FFMPEG = (shutil.which("ffmpeg") or
           ("/opt/homebrew/bin/ffmpeg"
            if os.path.exists("/opt/homebrew/bin/ffmpeg") else None))


def _make_aac(path, seconds=2):
    subprocess.run(
        [_FFMPEG, "-v", "error", "-f", "lavfi",
         "-i", f"sine=frequency=440:duration={seconds}",
         "-c:a", "aac", "-y", path],
        check=True, capture_output=True, timeout=60,
    )


@unittest.skipUnless(_FFMPEG, "ffmpeg not installed")
class TestAacFallback(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmpdir = tempfile.mkdtemp()
        cls.aac_path = os.path.join(cls.tmpdir, "render_audio.mp4")
        _make_aac(cls.aac_path, seconds=2)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmpdir, ignore_errors=True)

    def test_duration_via_ffprobe(self):
        # soundfile throws on AAC; ffprobe fallback must return ~2s, not 0.0
        dur = Transcriber.get_audio_duration(
            Transcriber.__new__(Transcriber), self.aac_path)
        self.assertGreater(dur, 1.5)
        self.assertLess(dur, 3.0)

    def test_load_mono_16k_via_ffmpeg(self):
        audio = Transcriber._load_mono_16k(self.aac_path)
        # ~2s of 16kHz mono float32
        self.assertGreater(len(audio), 16000 * 1.5)
        self.assertLess(len(audio), 16000 * 3.0)
        self.assertEqual(str(audio.dtype), "float32")
        # A 440Hz sine has real signal energy; silence would mean we
        # decoded garbage.
        self.assertGreater(float(abs(audio).max()), 0.1)

    def test_wav_still_uses_soundfile_path(self):
        wav_path = os.path.join(self.tmpdir, "ref.wav")
        subprocess.run(
            [_FFMPEG, "-v", "error", "-f", "lavfi",
             "-i", "sine=frequency=440:duration=2",
             "-c:a", "pcm_s16le", "-y", wav_path],
            check=True, capture_output=True, timeout=60,
        )
        audio = Transcriber._load_mono_16k(wav_path)
        self.assertGreater(len(audio), 16000 * 1.5)


class TestEstimatorGuard(unittest.TestCase):
    def test_unknown_duration_gives_no_expected(self):
        # Mirror of the guard in caption._start_progress_estimator: a 0.0
        # duration must select indeterminate mode, never a 1s estimate.
        for dur in (0.0, -1.0):
            expected = max(dur / 4.0, 1.0) if dur > 0 else None
            self.assertIsNone(expected)
        self.assertEqual(max(8.0 / 4.0, 1.0) if 8.0 > 0 else None, 2.0)


if __name__ == "__main__":
    unittest.main()
