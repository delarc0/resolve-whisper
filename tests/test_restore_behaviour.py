"""Execute _restore_deliver_state against a fake project.

The other tests in this area read source text. This one runs the function
and asserts what it actually does to the project, which is the only way to
prove the three cases stay distinct:

  snapshot read      -> put the exact values back
  changed, unreadable-> clear, so our path can't redirect their next export
  unchanged          -> touch nothing, even if the snapshot failed
"""
import os
import sys
import unittest

APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, APP_DIR)

import caption  # noqa: E402


class FakeProject:
    """Records the calls the restore makes."""

    def __init__(self):
        self.render_settings = []
        self.fmt = None
        self.mode = None

    def SetRenderSettings(self, settings):
        self.render_settings.append(settings)
        return True

    def SetCurrentRenderFormatAndCodec(self, fmt, codec):
        self.fmt = (fmt, codec)
        return True

    def SetCurrentRenderMode(self, mode):
        self.mode = mode
        return True

    def output_writes(self):
        return [s for s in self.render_settings
                if "TargetDir" in s or "CustomName" in s]


class TestRestoreDeliverState(unittest.TestCase):
    def test_exact_snapshot_is_put_back(self):
        p = FakeProject()
        caption._restore_deliver_state(
            p, {"format": "mp4", "codec": "H264"}, "single",
            saved_target_dir="/Volumes/BASED/EXPORT",
            saved_custom_name="reel_01")
        self.assertEqual(p.output_writes(),
                         [{"TargetDir": "/Volumes/BASED/EXPORT",
                           "CustomName": "reel_01"}])
        self.assertEqual(p.fmt, ("mp4", "H264"))
        self.assertEqual(p.mode, "single")

    def test_changed_but_unreadable_is_cleared(self):
        # Leaving the tool's own path would silently redirect the next export.
        p = FakeProject()
        caption._restore_deliver_state(p, {}, None, clear_unknown_output=True)
        self.assertEqual(p.output_writes(),
                         [{"TargetDir": "", "CustomName": ""}])

    def test_untouched_settings_are_left_alone(self):
        # The failure this guards: probe fails, then LoadRenderPreset fails
        # before anything wrote TargetDir. Clearing here would destroy a
        # setting this run never touched.
        p = FakeProject()
        caption._restore_deliver_state(p, {}, None, clear_unknown_output=False)
        self.assertEqual(p.output_writes(), [],
                         "cleared settings the run never modified")

    def test_partial_snapshot_still_restores(self):
        p = FakeProject()
        caption._restore_deliver_state(p, {}, None,
                                       saved_target_dir="/tmp/x",
                                       saved_custom_name=None)
        self.assertEqual(p.output_writes(),
                         [{"TargetDir": "/tmp/x", "CustomName": ""}])

    def test_incomplete_format_snapshot_is_not_applied(self):
        # A half-read format would set a codec that doesn't match the format.
        p = FakeProject()
        caption._restore_deliver_state(p, {"format": "mp4", "codec": ""}, None)
        self.assertIsNone(p.fmt)


if __name__ == "__main__":
    unittest.main()
