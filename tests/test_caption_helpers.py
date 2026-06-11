"""Tests for caption.py pure helpers (no Resolve required).

Run from project root:  ./.venv/bin/python -m unittest tests/test_caption_helpers.py
"""
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# caption.py imports DaVinciResolveScript lazily inside get_resolve(), so the
# top-level import is safe outside Resolve.
from caption import (  # noqa: E402
    _validate_audio_only_settings,
    _expected_output_path,
    _start_rendering_compat,
)


class TestValidateAudioOnlySettings(unittest.TestCase):
    def test_correct_settings_pass(self):
        self.assertEqual(
            _validate_audio_only_settings({
                "ExportVideo": False,
                "ExportAudio": True,
                "AudioCodec": "LinearPCM",
            }),
            [],
        )

    def test_resolve21_is_prefixed_keys_pass(self):
        # Resolve 21 renamed job-dict keys: ExportVideo -> IsExportVideo etc.
        self.assertEqual(
            _validate_audio_only_settings({
                "IsExportVideo": False,
                "IsExportAudio": True,
                "AudioCodec": "lpcm",
            }),
            [],
        )

    def test_resolve21_is_prefixed_video_true_fails(self):
        problems = _validate_audio_only_settings({
            "IsExportVideo": True,
            "IsExportAudio": True,
        })
        self.assertEqual(len(problems), 1)
        self.assertIn("IsExportVideo", problems[0])

    def test_resolve21_is_prefixed_audio_false_fails(self):
        problems = _validate_audio_only_settings({
            "IsExportVideo": False,
            "IsExportAudio": False,
        })
        self.assertEqual(len(problems), 1)
        self.assertIn("IsExportAudio", problems[0])

    def test_resolve21_output_filename_used_first(self):
        # Resolve 21 job dicts carry OutputFilename; it should win over probing.
        with tempfile.TemporaryDirectory() as d:
            target = os.path.join(d, "Timeline 1.wav")
            with open(target, "w") as f:
                f.write("x")
            path = _expected_output_path(
                {"TargetDir": d, "OutputFilename": "Timeline 1.wav"},
                d, "fallback",
            )
            self.assertEqual(path, target)

    def test_export_video_true_fails(self):
        problems = _validate_audio_only_settings({
            "ExportVideo": True,
            "ExportAudio": True,
        })
        self.assertEqual(len(problems), 1)
        self.assertIn("ExportVideo", problems[0])

    def test_export_audio_false_fails(self):
        problems = _validate_audio_only_settings({
            "ExportVideo": False,
            "ExportAudio": False,
        })
        self.assertEqual(len(problems), 1)
        self.assertIn("ExportAudio", problems[0])

    def test_string_truthy_falsy_normalized(self):
        # Resolve sometimes returns "0"/"1" instead of bool; we accept both.
        self.assertEqual(
            _validate_audio_only_settings({
                "ExportVideo": "0",
                "ExportAudio": "1",
            }),
            [],
        )
        self.assertNotEqual(
            _validate_audio_only_settings({
                "ExportVideo": "1",
                "ExportAudio": "1",
            }),
            [],
        )

    def test_missing_keys_treated_as_unknown_not_failing(self):
        # If keys are absent, we don't fail (Resolve may not surface them all)
        self.assertEqual(_validate_audio_only_settings({}), [])

    def test_non_dict_returns_problem(self):
        self.assertEqual(_validate_audio_only_settings(None), ["render settings unreadable"])
        self.assertEqual(_validate_audio_only_settings("oops"), ["render settings unreadable"])

    def test_unusual_codec_does_not_fail(self):
        # We log info but don't fail on non-PCM codecs
        self.assertEqual(
            _validate_audio_only_settings({
                "ExportVideo": False,
                "ExportAudio": True,
                "AudioCodec": "AAC",
            }),
            [],
        )


class TestExpectedOutputPath(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="rw_test_")

    def tearDown(self):
        import shutil
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_finds_wav_at_expected_path(self):
        wav = os.path.join(self.tmp, "myvideo_audio.wav")
        open(wav, "w").close()
        self.assertEqual(
            _expected_output_path(
                {"TargetDir": self.tmp, "CustomName": "myvideo_audio"},
                self.tmp, "myvideo_audio",
            ),
            wav,
        )

    def test_falls_back_to_dir_scan(self):
        # Resolve sometimes appends a suffix or changes case; CustomName-substring scan saves us.
        wav = os.path.join(self.tmp, "myvideo_audio_001.wav")
        open(wav, "w").close()
        result = _expected_output_path(
            {"TargetDir": self.tmp, "CustomName": "myvideo_audio"},
            self.tmp, "myvideo_audio",
        )
        self.assertEqual(result, wav)

    def test_missing_returns_empty_string(self):
        result = _expected_output_path(
            {"TargetDir": self.tmp, "CustomName": "nope"},
            self.tmp, "nope",
        )
        self.assertEqual(result, "")

    def test_uses_output_dir_fallback(self):
        wav = os.path.join(self.tmp, "fallback_audio.wav")
        open(wav, "w").close()
        # job_settings has no TargetDir -> falls back to output_dir arg
        result = _expected_output_path({}, self.tmp, "fallback_audio")
        self.assertEqual(result, wav)

    def test_nonexistent_dir_returns_empty(self):
        result = _expected_output_path(
            {"TargetDir": "/no/such/dir"},
            "/no/such/dir", "x",
        )
        self.assertEqual(result, "")


class FakeProject:
    """Minimal project stand-in for testing StartRendering compatibility."""
    def __init__(self, accepted_signature):
        # accepted_signature: which call form should "succeed"
        # one of: "single", "single_with_bool", "list_with_bool"
        self.accepted = accepted_signature
        self.calls = []

    def StartRendering(self, *args, **kwargs):
        self.calls.append((args, kwargs))
        if self.accepted == "single" and len(args) == 1 and isinstance(args[0], str):
            return True
        if self.accepted == "single_with_bool" and len(args) == 2 and isinstance(args[0], str):
            return True
        if self.accepted == "list_with_bool" and len(args) == 2 and isinstance(args[0], list):
            return True
        # Real Resolve raises TypeError for wrong arg counts in some versions
        raise TypeError(f"unaccepted form: args={args} kwargs={kwargs}")


class TestStartRenderingCompat(unittest.TestCase):
    def test_single_arg_form_succeeds(self):
        p = FakeProject("single")
        self.assertTrue(_start_rendering_compat(p, "job1"))

    def test_two_arg_form_succeeds(self):
        p = FakeProject("single_with_bool")
        self.assertTrue(_start_rendering_compat(p, "job1"))

    def test_list_form_succeeds(self):
        p = FakeProject("list_with_bool")
        self.assertTrue(_start_rendering_compat(p, "job1"))

    def test_all_fail_returns_false(self):
        # No form accepted -> returns False (does not raise)
        class AlwaysFail:
            def StartRendering(self, *a, **kw):
                raise TypeError("never")
        self.assertFalse(_start_rendering_compat(AlwaysFail(), "job1"))

    def test_falsy_return_treated_as_failure(self):
        class FalsyReturn:
            def StartRendering(self, *a, **kw):
                return False
        self.assertFalse(_start_rendering_compat(FalsyReturn(), "job1"))


if __name__ == "__main__":
    unittest.main()
