"""Tests for settings_ui persistence + option handling (no Tk needed)."""
import json
import os
import sys
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import settings_ui  # noqa: E402


class TestSettingsPersistence(unittest.TestCase):
    def setUp(self):
        import tempfile
        self._tmp = tempfile.NamedTemporaryFile(suffix=".json", delete=False)
        self._tmp.close()
        os.unlink(self._tmp.name)
        self._patch = mock.patch.object(
            settings_ui, "SETTINGS_PATH", self._tmp.name)
        self._patch.start()

    def tearDown(self):
        self._patch.stop()
        try:
            os.unlink(self._tmp.name)
        except FileNotFoundError:
            pass

    def test_defaults_when_no_file(self):
        self.assertEqual(settings_ui.load_settings(), settings_ui.DEFAULTS)

    def test_round_trip(self):
        chosen = {"language": "en", "max_words": 3, "max_chars": 22,
                  "uppercase": False, "strip_punctuation": True}
        settings_ui.save_settings(chosen)
        self.assertEqual(settings_ui.load_settings(), chosen)

    def test_corrupt_file_falls_back_to_defaults(self):
        with open(self._tmp.name, "w") as f:
            f.write("{not json")
        self.assertEqual(settings_ui.load_settings(), settings_ui.DEFAULTS)

    def test_wrong_types_ignored_per_key(self):
        settings_ui.save_settings({"language": 42, "max_words": "three",
                                   "max_chars": 37, "uppercase": True,
                                   "strip_punctuation": False})
        loaded = settings_ui.load_settings()
        # Bad-typed keys revert to defaults; good keys survive.
        self.assertEqual(loaded["language"], settings_ui.DEFAULTS["language"])
        self.assertEqual(loaded["max_words"], settings_ui.DEFAULTS["max_words"])
        self.assertEqual(loaded["max_chars"], 37)

    def test_unknown_keys_dropped(self):
        settings_ui.save_settings(dict(settings_ui.DEFAULTS, evil="x"))
        self.assertNotIn("evil", settings_ui.load_settings())


class TestWithStored(unittest.TestCase):
    def test_known_value_unchanged(self):
        opts = settings_ui._with_stored(
            settings_ui.LINE_LENGTHS, 42, "{} characters")
        self.assertEqual(opts, settings_ui.LINE_LENGTHS)

    def test_off_list_value_appended(self):
        opts = settings_ui._with_stored(
            settings_ui.LINE_LENGTHS, 45, "{} characters")
        self.assertIn(("45 characters", 45), opts)


class TestDialogProtocol(unittest.TestCase):
    """caption.py parses the SETTINGS: stdout line; lock it down."""

    def test_settings_line_is_json(self):
        chosen = {"language": "auto", "max_words": 0, "max_chars": 42,
                  "uppercase": True, "strip_punctuation": False}
        line = "SETTINGS:" + json.dumps(chosen)
        self.assertTrue(line.startswith("SETTINGS:"))
        self.assertEqual(json.loads(line[len("SETTINGS:"):]), chosen)


if __name__ == "__main__":
    unittest.main()
