"""A bad value in caption_config.json must never reach the pipeline.

Every one of these used to be accepted by the type check and crash later,
at SRT-write time, which is AFTER the timeline audio has been rendered: the
user lost minutes of render and got a traceback instead of captions.
"""
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, APP_DIR)

import config  # noqa: E402


def load_with(raw):
    """Run load_config() against a config file holding `raw` (text or dict)."""
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "caption_config.json")
        Path(path).write_text(
            raw if isinstance(raw, str) else json.dumps(raw),
            encoding="utf-8")
        original = config.CONFIG_PATH
        config.CONFIG_PATH = path
        try:
            return config.load_config()
        finally:
            config.CONFIG_PATH = original


class TestGoodValuesSurvive(unittest.TestCase):
    def test_a_valid_override_is_applied(self):
        self.assertEqual(load_with({"max_chars_per_line": 30})
                         ["max_chars_per_line"], 30)

    def test_null_means_auto_detect_on_a_nullable_key(self):
        self.assertIsNone(load_with({"language": None})["language"])

    def test_an_int_is_accepted_for_a_float_key(self):
        self.assertEqual(load_with({"min_duration_s": 2})["min_duration_s"], 2.0)


class TestBadValuesFallBackToDefaults(unittest.TestCase):
    def assert_default(self, key, bad_value):
        self.assertEqual(load_with({key: bad_value})[key],
                         config.DEFAULT_CONFIG[key],
                         f"{key}={bad_value!r} should have been rejected")

    def test_wrong_type(self):
        self.assert_default("max_chars_per_line", "forty-two")

    def test_null_on_a_numeric_key(self):
        # The generated file teaches `null` as the "unset" idiom, so copying
        # that pattern onto a number is a realistic mistake, not an exotic one.
        self.assert_default("max_chars_per_line", None)

    def test_bool_is_not_a_number(self):
        # bool subclasses int, so a plain isinstance check let True through
        # as the value 1.
        self.assert_default("max_lines", True)

    def test_number_is_not_a_bool(self):
        self.assert_default("uppercase", 1)

    def test_infinity_from_a_huge_literal(self):
        # json.load turns 1e400 into float("inf"), which IS a float.
        self.assertEqual(
            load_with('{"max_duration_s": 1e400}')["max_duration_s"],
            config.DEFAULT_CONFIG["max_duration_s"])

    def test_nan(self):
        self.assertEqual(
            load_with('{"max_duration_s": NaN}')["max_duration_s"],
            config.DEFAULT_CONFIG["max_duration_s"])

    def test_zero_where_only_positive_makes_sense(self):
        self.assert_default("max_chars_per_line", 0)

    def test_negative_duration(self):
        self.assert_default("min_duration_s", -1.0)

    def test_negative_where_zero_is_allowed(self):
        # keep_srt_days accepts 0 (cleanup off) but not -5.
        self.assertEqual(load_with({"keep_srt_days": 0})["keep_srt_days"], 0)
        self.assert_default("keep_srt_days", -5)

    def test_fractional_value_on_an_int_key(self):
        self.assert_default("max_lines", 1.5)

    def test_whole_float_on_an_int_key_is_accepted(self):
        self.assertEqual(load_with({"max_lines": 2.0})["max_lines"], 2)


class TestABrokenFileNeverStopsTheRun(unittest.TestCase):
    def test_invalid_json_falls_back_to_every_default(self):
        self.assertEqual(load_with("{not json at all"), config.DEFAULT_CONFIG)

    def test_a_json_list_is_not_a_config(self):
        self.assertEqual(load_with("[1, 2, 3]"), config.DEFAULT_CONFIG)

    def test_unknown_keys_are_ignored_not_fatal(self):
        loaded = load_with({"colour_scheme": "neon", "max_lines": 2})
        self.assertNotIn("colour_scheme", loaded)
        self.assertEqual(loaded["max_lines"], 2)

    def test_one_bad_key_does_not_discard_the_good_ones(self):
        loaded = load_with({"max_lines": "two", "max_chars_per_line": 30})
        self.assertEqual(loaded["max_lines"], config.DEFAULT_CONFIG["max_lines"])
        self.assertEqual(loaded["max_chars_per_line"], 30)


if __name__ == "__main__":
    unittest.main()
